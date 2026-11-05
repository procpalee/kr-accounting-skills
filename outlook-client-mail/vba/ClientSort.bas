Attribute VB_Name = "ClientSort"
' ClientSort - 거래처 메일 자동 분류(outlook-client-mail 스킬과 한 묶음, setup.py 가 이 파일을 만든다)
'
' 하는 일(모두 정해진 규칙대로 - AI 판단 없음):
'   1) 분류: 받은 편지함 · 보낸 편지함 · #watch 폴더에 메일이 들어오면 그 거래처 Outlook 폴더로(받음 이동 · 보냄 사본)
'   2) 저장: 거래처 Outlook 폴더에 메일이 들어오면(자동 분류 · 끌어다 넣기 · [거래처로 분류] 모두) mail_sync.py 를 불러
'            원본 .msg 와 받은 메일 첨부를 거래처 폴더에 저장
'   3) 연결: Outlook 을 켤 때 · [폴더 연결 맞추기] 버튼 - mail_sync_all.py 가 거래처 폴더 ↔ Outlook 폴더(mail.json outlook.entry_id)를
'            맞추고 거래처 목록(Client_Map.tsv)을 다시 만든다. 켤 때는 모든 거래처 수집도 한 번
' 분류 판단 순서(고정) - 제목이 #exclude(그룹 가입 · 공유 알림 등)에 걸리면 분류 안 함:
'   ① 제목에 거래처명 · 키워드 - 긴 이름 우선('A홀딩스' 가 걸리면 'A' 는 버림), 여러 거래처면 분류 안 함
'   ② 보낸 사람 · 받는 사람 · 참조 주소가 거래처 주소/도메인 - 여러 거래처면 분류 안 함, 단체 메일(받는 사람 많음)은 건너뜀
'   ③ 같은 대화(ConversationID)의 메일이 이미 한 거래처 폴더에 있으면 그 폴더
'   ④ 셋 다 아니면 그대로 둔다 - 사용자가 폴더에 끌어다 넣거나 [거래처로 분류] 로 지정
' 보낸 메일 사본은 이미 있거나 나도 받는 사람이면 만들지 않는다. 지우기 · 발송은 하지 않는다. 기록: _메일분류\Client_Sort_log.tsv
'
' 빠른 실행 도구 모음 버튼용 매크로:
'   ClientSort_File         [거래처로 분류]      선택한 메일(받은 · 보낸 · 잘못 들어간 메일 모두)을 거래처 폴더로 - 규칙 추측을 기본값으로 보여 주고 바꿀 수 있음
'   ClientSort_LinkFolders  [폴더 연결 맞추기]   새 거래처 폴더 만들기 · 연결 확인 + 거래처 목록 다시 읽기, 결과 요약 표시
'   ClientSort_Help         [사용법]             규칙 · 버튼 · 문제 해결 안내 페이지 열기
Option Explicit

Private Const MAP_DIR As String = "{{CONF_DIR}}\"
Private Const PR_MSGID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"

Private gN As Long
Private gName() As String
Private gPath() As String
Private gKw() As Variant
Private gAddr() As Variant
Private gEng() As String
Private gEid() As String
Private gPy As String
Private gSync As String
Private gAll As String
Private gRoot As String
Private gWatch As Collection
Private gExcl As Collection
Private gMaxR As Long
Private gStamp As Date
Private gConv As Object
Private gDone As Object
Private gWatchers As Collection                                       ' 분류 감시(받은 · 보낸 · #watch)
Private gCaseW As Collection                                          ' 거래처 Outlook 폴더 감시(들어오면 저장)
Private gMe As String
Private gMyDomain As String
Private gReady As Boolean
Private gHelp As String
Private gPending As Collection                                        ' 켤 때 못 찾은 감시 폴더(Outlook 이 폴더 목록을 덜 읽었을 때) - 30초마다 다시
Private gTimer As LongPtr
Private gTries As Long
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long

' ───────────────────────── 시작 · 종료 (ThisOutlookSession 에서 부름) ─────────────────────────
' 시작 - ThisOutlookSession 의 Startup · MAPILogonComplete · NewMailEx 와 버튼이 부른다. 이미 시작했으면 아무것도 안 함(여러 번 불려도 한 번)
Public Sub Start(Optional ByVal src As String = "")
    On Error GoTo Fail
    If gReady Then Exit Sub
    Set gDone = CreateObject("Scripting.Dictionary")
    LoadMap
    BuildConv
    gMe = LCase$(SmtpOfEntry(Application.Session.CurrentUser.AddressEntry))
    If InStr(gMe, "@") > 0 Then gMyDomain = Mid$(gMe, InStr(gMe, "@") + 1)
    Set gWatchers = New Collection
    AddWatch Application.Session.GetDefaultFolder(6), False          ' 받은 편지함
    AddWatch Application.Session.GetDefaultFolder(5), True           ' 보낸 편지함
    Dim w As Variant, f As Object
    Set gPending = New Collection
    For Each w In gWatch
        Set f = FolderByPath(CStr(w))
        If f Is Nothing Then
            gPending.Add CStr(w)                                      ' 켜는 중이라 폴더 목록이 아직 없을 수 있음 - 다시 시도
        Else
            AddWatch f, False
        End If
    Next
    WireCases
    gReady = True
    If gPending.Count > 0 Or gCaseW.Count < CaseCount() Then StartRetry
    CatchUp
    SaveLastRun
    WriteLog "시작", "", IIf(Len(src) > 0, src, "버튼"), ""
    RunPy gAll, "", False                                             ' 폴더 연결 맞추기 + 거래처 목록 + 모든 거래처 수집(창 없이)
    Exit Sub
Fail:
    WriteLog "오류", "", "Start: " & Err.Description, ""
End Sub

Public Sub SaveLastRun()
    On Error Resume Next
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(MAP_DIR & "last_run.txt", True, False)
    ts.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ts.Close
End Sub

Private Sub AddWatch(ByVal f As Object, ByVal isSent As Boolean)
    Dim w As ClientSortWatch
    Set w = New ClientSortWatch
    Set w.Items = f.Items
    w.IsSent = isSent
    gWatchers.Add w
End Sub

' 거래처 Outlook 폴더마다 감시 - 거래처 목록이 바뀌면 다시 건다
Private Sub WireCases()
    Dim i As Long, f As Object, w As ClientSortWatch
    Set gCaseW = New Collection
    For i = 0 To gN - 1
        Set f = ClientFolder(i)
        If Not f Is Nothing Then
            Set w = New ClientSortWatch
            Set w.Items = f.Items
            w.CaseIdx = i
            gCaseW.Add w
        End If
    Next
End Sub

' 분류 감시 폴더에 메일이 들어올 때(ClientSortWatch 가 부름)
Public Sub HandleItem(ByVal it As Object, ByVal isSent As Boolean)
    On Error GoTo Fail
    If Not gReady Then Exit Sub
    If it.Class <> 43 Then Exit Sub                                   ' 메일만(회의 요청 · 보고서 제외)
    EnsureMap
    If isSent Then
        If gDone.Exists(MsgId(it)) Then Exit Sub                     ' 방금 만든 사본이 다시 들어온 것
    End If
    Dim why As String, i As Long
    i = Classify(it, why)
    If i >= 0 Then FileTo it, i, isSent, why
    Exit Sub
Fail:
    If isSent Then Exit Sub                                           ' 방금 옮긴 사본의 이벤트 - 이미 처리됨
    WriteLog "오류", "", "HandleItem: " & Err.Description, ""
End Sub

' 거래처 Outlook 폴더에 메일이 들어올 때 - 누가 넣었든 그 거래처 메일 저장(원본 · 첨부)
Public Sub CaseAdded(ByVal i As Long)
    On Error Resume Next
    If Not gReady Then Exit Sub
    If i < 0 Or i >= gN Then Exit Sub
    If Len(gEng(i)) = 0 Then Exit Sub
    RunPy gSync, gEng(i), False
End Sub

' 켤 때 못 찾은 감시 폴더 · 거래처 폴더를 30초마다 다시(최대 20번 = 10분)
Private Sub StartRetry()
    If gTimer <> 0 Then Exit Sub
    gTries = 0
    gTimer = SetTimer(0, 0, 30000, AddressOf RetryTimerProc)
End Sub

Public Sub RetryTimerProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
    On Error Resume Next
    gTries = gTries + 1
    RetryWatch
    If (gPending.Count = 0 And gCaseW.Count >= CaseCount()) Or gTries >= 20 Then
        KillTimer 0, gTimer
        gTimer = 0
        If gPending.Count > 0 Then WriteLog "설정", "", "감시 폴더를 찾지 못함(이름이 바뀌었으면 mail_sort.json 의 watch 를 고치기)", JoinColl(gPending)
    End If
End Sub

Private Sub RetryWatch()
    On Error Resume Next
    Dim k As Long, f As Object
    If gPending Is Nothing Then Exit Sub
    For k = gPending.Count To 1 Step -1
        Set f = FolderByPath(CStr(gPending(k)))
        If Not f Is Nothing Then
            AddWatch f, False
            WriteLog "설정", "", "감시 폴더 연결(다시 시도)", CStr(gPending(k))
            gPending.Remove k
        End If
    Next
    If gCaseW.Count < CaseCount() Then WireCases
End Sub

Private Function CaseCount() As Long
    Dim i As Long
    For i = 0 To gN - 1
        If Len(gEng(i)) > 0 Then CaseCount = CaseCount + 1
    Next
End Function

Private Function JoinColl(ByVal c As Collection) As String
    Dim x As Variant
    For Each x In c
        JoinColl = JoinColl & IIf(Len(JoinColl) > 0, ", ", "") & CStr(x)
    Next
End Function

' Outlook 이 꺼져 있던 동안 받은 · 보낸 메일 분류
Private Sub CatchUp()
    On Error GoTo Fail
    Dim fso As Object, p As String, since As Date, ts As Object, s As String
    p = MAP_DIR & "last_run.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(p) Then Exit Sub                            ' 처음 설치: 기존 메일은 건드리지 않음
    Set ts = fso.OpenTextFile(p, 1)
    s = Trim$(ts.ReadLine)
    ts.Close
    since = CDate(s) - TimeSerial(0, 10, 0)
    Dim w As Variant, todo As New Collection, x As Variant
    For Each w In gWatchers
        CollectSince w.Items, w.IsSent, since, todo
    Next
    For Each x In todo
        HandleItem x(0), x(1)
    Next
    Exit Sub
Fail:
    WriteLog "오류", "", "CatchUp: " & Err.Description, ""
End Sub

Private Sub CollectSince(ByVal its As Object, ByVal isSent As Boolean, ByVal since As Date, ByRef todo As Collection)
    On Error Resume Next
    Dim key As String, it As Object, t As Date
    key = IIf(isSent, "[SentOn]", "[ReceivedTime]")
    its.Sort key, True
    For Each it In its
        If it.Class = 43 Then
            t = IIf(isSent, it.SentOn, it.ReceivedTime)
            If t < since Then Exit For
            todo.Add Array(it, isSent)
        End If
    Next
End Sub

' 창 없이 Python 도구 실행(경로는 거래처 목록 #python · #mail_sync · #mail_all)
Private Sub RunPy(ByVal script As String, ByVal arg As String, ByVal wait As Boolean)
    On Error Resume Next
    If Len(gPy) = 0 Or Len(script) = 0 Then Exit Sub
    Dim cmd As String
    cmd = """" & gPy & """ -X utf8 """ & script & """"
    If Len(arg) > 0 Then cmd = cmd & " " & IIf(Left$(arg, 2) = "--", arg, """" & arg & """")
    CreateObject("WScript.Shell").Run cmd, 0, wait
End Sub

' ───────────────────────── 판단 ─────────────────────────
Private Function Classify(ByVal m As Object, ByRef why As String) As Long
    Dim i As Long, k As Variant, n As Long, a As Long, b As Long, res As Long, dropped As Boolean
    Dim mi() As Long, mk() As String, subj As String
    why = ""
    subj = m.Subject
    For Each k In gExcl
        If InStr(1, subj, CStr(k), vbTextCompare) > 0 Then why = "제외 제목": Classify = -1: Exit Function
    Next
    ' ① 제목
    ReDim mi(0 To 0): ReDim mk(0 To 0)
    For i = 0 To gN - 1
        For Each k In gKw(i)
            If Len(k) > 0 Then
                If KwIn(subj, CStr(k)) Then
                    n = n + 1
                    ReDim Preserve mi(0 To n): ReDim Preserve mk(0 To n)
                    mi(n) = i: mk(n) = CStr(k)
                End If
            End If
        Next
    Next
    res = -1
    For a = 1 To n
        dropped = False
        For b = 1 To n
            If mi(b) <> mi(a) And Len(mk(b)) > Len(mk(a)) And InStr(1, mk(b), mk(a), vbTextCompare) > 0 Then dropped = True
        Next
        If Not dropped Then
            If res = -1 Then
                res = mi(a)
            ElseIf res <> mi(a) Then
                res = -2
            End If
        End If
    Next
    If res >= 0 Then why = "제목": Classify = res: Exit Function
    If res = -2 Then why = "제목에 여러 거래처": Classify = -2: Exit Function
    ' ② 주소
    Dim addrs As Collection, ad As Variant, aa As Variant
    Set addrs = Participants(m)
    If addrs.Count - 1 > gMaxR Then why = "단체 메일": Classify = -1: Exit Function
    res = -1
    For Each ad In addrs
        If ad <> gMe Then
            For i = 0 To gN - 1
                For Each aa In gAddr(i)
                    If AddrMatch(CStr(ad), CStr(aa)) Then
                        If res = -1 Then
                            res = i
                        ElseIf res <> i Then
                            res = -2
                        End If
                    End If
                Next
            Next
        End If
    Next
    If res >= 0 Then why = "주소": Classify = res: Exit Function
    If res = -2 Then why = "주소에 여러 거래처": Classify = -2: Exit Function
    ' ③ 같은 대화
    Dim cid As String
    On Error Resume Next
    cid = m.ConversationID
    On Error GoTo 0
    If Len(cid) > 0 Then
        If gConv.Exists(cid) Then
            If gConv(cid) >= 0 Then why = "대화": Classify = gConv(cid): Exit Function
        End If
    End If
    Classify = -1
End Function

' 제목에 키워드가 있는지 - 2자 이하 키워드는 앞 글자가 한글이 아닐 때만(다른 단어 속 우연 일치 방지)
Private Function KwIn(ByVal s As String, ByVal kw As String) As Boolean
    Dim p As Long, c As Long
    p = InStr(1, s, kw, vbTextCompare)
    Do While p > 0
        If Len(kw) > 2 Or p = 1 Then KwIn = True: Exit Function
        c = AscW(Mid$(s, p - 1, 1))
        If c < 0 Then c = c + 65536
        If c < &HAC00& Or c > &HD7A3& Then KwIn = True: Exit Function
        p = InStr(p + 1, s, kw, vbTextCompare)
    Loop
End Function

Private Function AddrMatch(ByVal ad As String, ByVal pat As String) As Boolean
    If Len(pat) = 0 Or Len(ad) = 0 Then Exit Function
    If InStr(pat, "@") > 0 Then
        AddrMatch = (ad = pat)
    Else
        AddrMatch = (Right$(ad, Len(pat) + 1) = "@" & pat) Or (Right$(ad, Len(pat) + 1) = "." & pat)
    End If
End Function

Private Function Participants(ByVal m As Object) As Collection
    Dim c As New Collection, r As Object
    On Error Resume Next
    c.Add SenderSmtp(m)
    For Each r In m.Recipients
        c.Add LCase$(SmtpOfEntry(r.AddressEntry))
    Next
    Set Participants = c
End Function

Private Function SenderSmtp(ByVal m As Object) As String
    On Error Resume Next
    If m.SenderEmailType = "EX" Then
        SenderSmtp = LCase$(m.Sender.GetExchangeUser.PrimarySmtpAddress)
    Else
        SenderSmtp = LCase$(m.SenderEmailAddress)
    End If
End Function

Private Function SmtpOfEntry(ByVal ae As Object) As String
    On Error Resume Next
    If ae.Type = "EX" Then
        SmtpOfEntry = ae.GetExchangeUser.PrimarySmtpAddress
    Else
        SmtpOfEntry = ae.Address
    End If
End Function

Private Function MsgId(ByVal m As Object) As String
    On Error Resume Next
    MsgId = m.PropertyAccessor.GetProperty(PR_MSGID)
    If Len(MsgId) = 0 Then MsgId = m.EntryID
End Function

' ───────────────────────── 옮기기 ─────────────────────────
' 폴더에 들어가면 그 폴더 감시(CaseAdded)가 저장을 부른다
' isSent = 보낸 편지함에 있는 메일(원본은 두고 사본), 그 밖(받은 편지함 · 다른 거래처 폴더 등)은 이동. manual = [거래처로 분류]
Private Function FileTo(ByVal it As Object, ByVal i As Long, ByVal isSent As Boolean, ByVal why As String, Optional ByVal manual As Boolean = False) As Boolean
    Dim f As Object, subj As String, cid As String, id As String, c As Object
    Set f = ClientFolder(i)
    subj = it.Subject
    If f Is Nothing Then WriteLog "폴더 없음", gName(i), why, subj: Exit Function
    On Error Resume Next
    cid = it.ConversationID
    On Error GoTo 0
    If isSent Then
        id = MsgId(it)
        gDone(id) = True
        If AlreadyIn(f, id) Then FileTo = True: Exit Function
        If ToMe(it) And Not manual Then WriteLog "보냄 사본 생략", gName(i), "나에게 보낸 메일 - 받은 메일로 분류", subj: Exit Function
        Set c = it.Copy
        c.Move f
        WriteLog "보냄 사본", gName(i), why, subj
    Else
        If it.Parent.EntryID = f.EntryID Then FileTo = True: Exit Function
        it.Move f
        WriteLog IIf(manual, "직접 이동", "받음 이동"), gName(i), why, subj
    End If
    If Len(cid) > 0 Then
        If Not gConv.Exists(cid) Or manual Then gConv(cid) = i
    End If
    FileTo = True
End Function

' 받는 사람 · 참조에 내가 있으면 받은 메일이 따로 분류되므로 보낸 사본은 만들지 않는다(같은 메일 두 통 방지)
Private Function ToMe(ByVal m As Object) As Boolean
    Dim r As Object
    On Error Resume Next
    For Each r In m.Recipients
        If LCase$(SmtpOfEntry(r.AddressEntry)) = gMe Then ToMe = True: Exit Function
    Next
End Function

Private Function AlreadyIn(ByVal f As Object, ByVal id As String) As Boolean
    On Error Resume Next
    If Left$(id, 1) <> "<" Then Exit Function
    AlreadyIn = (f.Items.Restrict("@SQL=""" & PR_MSGID & """ = '" & Replace(id, "'", "''") & "'").Count > 0)
End Function

' ───────────────────────── 거래처 목록 · 폴더 ─────────────────────────
Private Sub EnsureMap()
    On Error Resume Next
    If FileDateTime(MAP_DIR & "Client_Map.tsv") <> gStamp Then
        LoadMap
        WireCases
    End If
End Sub

Private Sub LoadMap()
    Dim txt As String, lines As Variant, ln As Variant, s As String, p As Variant
    txt = ReadUtf8(MAP_DIR & "Client_Map.tsv")
    gStamp = FileDateTime(MAP_DIR & "Client_Map.tsv")
    gRoot = "거래처": gMaxR = 15: gN = 0
    Set gWatch = New Collection
    Set gExcl = New Collection
    ReDim gName(0 To 0): ReDim gPath(0 To 0): ReDim gKw(0 To 0): ReDim gAddr(0 To 0): ReDim gEng(0 To 0): ReDim gEid(0 To 0)
    gPy = "": gSync = "": gAll = "": gHelp = ""
    lines = Split(Replace(txt, vbCr, ""), vbLf)
    For Each ln In lines
        s = CStr(ln)
        If Len(Trim$(s)) > 0 Then
            p = Split(s, vbTab)
            If Left$(s, 1) = "#" Then
                Select Case p(0)
                    Case "#client_root": gRoot = p(1)
                    Case "#watch": gWatch.Add CStr(p(1))
                    Case "#exclude": gExcl.Add CStr(p(1))
                    Case "#python": gPy = p(1)
                    Case "#mail_sync": gSync = p(1)
                    Case "#mail_all": gAll = p(1)
                    Case "#help": gHelp = p(1)
                    Case "#max_recipients": gMaxR = CLng(p(1))
                End Select
            ElseIf UBound(p) >= 1 Then
                ReDim Preserve gName(0 To gN): ReDim Preserve gPath(0 To gN): ReDim Preserve gKw(0 To gN)
                ReDim Preserve gAddr(0 To gN): ReDim Preserve gEng(0 To gN): ReDim Preserve gEid(0 To gN)
                gName(gN) = p(0)
                gPath(gN) = p(1)
                gKw(gN) = Split(Col(p, 2), ";")
                gAddr(gN) = Split(LCase$(Col(p, 3)), ";")
                gEng(gN) = Col(p, 5)
                gEid(gN) = Col(p, 6)
                gN = gN + 1
            End If
        End If
    Next
End Sub

Private Function Col(ByVal p As Variant, ByVal k As Long) As String
    If UBound(p) >= k Then Col = Trim$(CStr(p(k)))
End Function

Private Function ReadUtf8(ByVal path As String) As String
    Dim s As Object
    Set s = CreateObject("ADODB.Stream")
    s.Type = 2: s.Charset = "utf-8": s.Open
    s.LoadFromFile path
    ReadUtf8 = s.ReadText
    s.Close
End Function

Private Sub BuildConv()
    Dim i As Long, f As Object, it As Object, cid As String
    Set gConv = CreateObject("Scripting.Dictionary")
    For i = 0 To gN - 1
        Set f = ClientFolder(i)
        If Not f Is Nothing Then
            For Each it In f.Items
                cid = ""
                On Error Resume Next
                cid = it.ConversationID
                On Error GoTo 0
                If Len(cid) > 0 Then
                    If Not gConv.Exists(cid) Then
                        gConv(cid) = i
                    ElseIf gConv(cid) <> i Then
                        gConv(cid) = -2                               ' 두 거래처 폴더에 같은 대화 - 대화 기준은 쓰지 않음
                    End If
                End If
            Next
        End If
    Next
End Sub

Private Function FolderByPath(ByVal rel As String) As Object
    On Error GoTo Nope
    Dim f As Object, p As Variant
    Set f = Application.Session.DefaultStore.GetRootFolder
    For Each p In Split(rel, "\")
        Set f = f.Folders.Item(CStr(p))
    Next
    Set FolderByPath = f
    Exit Function
Nope:
    Set FolderByPath = Nothing
End Function

' 연결 ID(거래처 mail.json 의 outlook.entry_id)로 찾고, 없으면 경로로
Private Function ClientFolder(ByVal i As Long) As Object
    Dim f As Object
    If Len(gEid(i)) > 0 Then
        On Error Resume Next
        Set f = Application.Session.GetFolderFromID(gEid(i))
        On Error GoTo 0
    End If
    If f Is Nothing Then Set f = FolderByPath(gRoot & "\" & gPath(i))
    Set ClientFolder = f
End Function

Private Sub WriteLog(ByVal kind As String, ByVal comp As String, ByVal why As String, ByVal subj As String)
    On Error Resume Next
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(MAP_DIR & "Client_Sort_log.tsv", 8, True, -1)    ' 덧붙이기 · UTF-16
    ts.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & kind & vbTab & comp & vbTab & why & vbTab & Replace(subj, vbTab, " ")
    ts.Close
End Sub

Private Sub Remember(ByVal i As Long, ByVal addr As String)
    On Error GoTo Fail
    Dim s As Object, fso As Object, p As String
    p = MAP_DIR & "Client_Map_추가.tsv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set s = CreateObject("ADODB.Stream")
    s.Type = 2: s.Charset = "utf-8": s.Open
    If fso.FileExists(p) Then
        s.LoadFromFile p
        s.Position = s.Size
    End If
    s.WriteText gName(i) & vbTab & "주소" & vbTab & addr & vbTab & Format$(Date, "yyyy-mm-dd") & vbCrLf
    s.SaveToFile p, 2
    s.Close
    gAddr(i) = Split(Join(gAddr(i), ";") & ";" & LCase$(addr), ";")
    Exit Sub
Fail:
    WriteLog "오류", gName(i), "Remember: " & Err.Description, addr
End Sub

Private Function PickCompany(ByVal ask As String, Optional ByVal def As Long = -1) As Long
    Dim i As Long, s As String, v As String
    PickCompany = -1
    For i = 0 To gN - 1
        s = s & Right$(" " & (i + 1), 2) & ". " & gName(i) & IIf((i + 1) Mod 2 = 0, vbCrLf, "        ")
    Next
    v = InputBox(ask & vbCrLf & vbCrLf & s & vbCrLf & vbCrLf & "번호를 입력하세요(취소 = 그대로 둠)", "거래처로 분류", IIf(def >= 0, CStr(def + 1), ""))
    If IsNumeric(v) Then
        If CLng(v) >= 1 And CLng(v) <= gN Then PickCompany = CLng(v) - 1
    End If
End Function

Private Function IsMine(ByVal m As Object) As Boolean
    IsMine = (SenderSmtp(m) = gMe)
End Function

Private Function FirstExternal(ByVal m As Object) As String
    Dim ad As Variant
    For Each ad In Participants(m)
        If Len(ad) > 0 And ad <> gMe And Right$(ad, Len(gMyDomain) + 1) <> "@" & gMyDomain Then
            FirstExternal = CStr(ad): Exit Function
        End If
    Next
End Function

Private Sub EnsureStarted()
    If Not gReady Then Start
    EnsureMap
    RetryWatch
End Sub

Private Function ReadResult() As String
    On Error Resume Next
    ReadResult = ReadUtf8(MAP_DIR & "last_sync.txt")
End Function

' ───────────────────────── 빠른 실행 도구 모음 버튼 ─────────────────────────
' [거래처로 분류] 선택한 메일(여러 통 가능)을 거래처 폴더로 - 규칙이 정한 거래처를 기본값으로 보여 주고, 다른 번호를 넣으면 그 거래처로
'   보낸 편지함의 메일 = 사본, 그 밖(받은 편지함 · 다른 거래처 폴더에 잘못 들어간 메일) = 이동. 규칙으로 못 정했거나 바꿨으면 상대 주소를 기억할지 묻는다
Public Sub ClientSort_File()
    EnsureStarted
    Dim sel As New Collection, it As Object, i As Long, why As String, guess As Long, first As Boolean, ad As String
    Dim sentId As String, n As Long, ask As String
    For Each it In Application.ActiveExplorer.Selection
        If it.Class = 43 Then sel.Add it
    Next
    If sel.Count = 0 Then MsgBox "분류할 메일을 먼저 선택하세요(여러 통 선택 가능).", vbInformation, "거래처로 분류": Exit Sub
    guess = -1: first = True
    For Each it In sel                                                ' 모두 같은 거래처로 판단될 때만 기본값
        i = Classify(it, why)
        If first Then
            guess = i: first = False
        ElseIf i <> guess Then
            guess = -1
        End If
    Next
    If guess < 0 Then guess = -1
    If sel.Count = 1 Then
        ask = "'" & Left$(sel(1).Subject, 60) & "'"
    Else
        ask = sel.Count & "통을 한 거래처 폴더로"
    End If
    If guess >= 0 Then
        ask = ask & vbCrLf & "규칙 판단: " & gName(guess) & " - 맞으면 그대로 확인, 아니면 번호를 바꾸세요"
    ElseIf Len(why) > 0 Then
        ask = ask & vbCrLf & "규칙으로 정하지 못했습니다(" & why & ") - 거래처 번호를 넣으세요"
    Else
        ask = ask & vbCrLf & "규칙으로 정하지 못했습니다(제목 · 주소 · 대화에 단서 없음) - 거래처 번호를 넣으세요"
    End If
    i = PickCompany(ask, guess)
    If i < 0 Then Exit Sub
    If sel.Count = 1 And i <> guess Then                              ' 규칙이 못 맞혔으면 주소 기억 제안
        ad = FirstExternal(sel(1))
        If Len(ad) > 0 Then
            If MsgBox(ad & vbCrLf & vbCrLf & "이 주소를 '" & gName(i) & "' 로 기억할까요?" & vbCrLf & "다음부터 이 주소가 낀 메일은 자동으로 분류됩니다.", vbYesNo + vbQuestion, "거래처로 분류") = vbYes Then Remember i, ad
        End If
    End If
    sentId = Application.Session.GetDefaultFolder(5).EntryID
    For Each it In sel
        If FileTo(it, i, (it.Parent.EntryID = sentId), IIf(i = guess, "규칙 확인", "직접 지정"), True) Then n = n + 1
    Next
    MsgBox n & "통을 '" & gName(i) & "' 폴더에 넣었습니다." & vbCrLf & "(보낸 편지함 메일은 사본, 나머지는 이동)" & vbCrLf & vbCrLf & "원본 · 첨부는 거래처 폴더에 자동 저장됩니다.", vbInformation, "거래처로 분류"
End Sub

' [폴더 연결 맞추기] 새 거래처 폴더 만들기 · 거래처 폴더 ↔ Outlook 폴더 연결 확인 + 거래처 목록 다시 읽기 - 결과 요약 표시
Public Sub ClientSort_LinkFolders()
    EnsureStarted
    RunPy gAll, "--link-only", True
    LoadMap
    BuildConv
    WireCases
    RetryWatch
    MsgBox Left$(ReadResult(), 1800), vbInformation, "폴더 연결 맞추기"
End Sub

' [사용법] 규칙 · 버튼 · 문제 해결 안내 페이지
Public Sub ClientSort_Help()
    EnsureStarted
    If Len(gHelp) = 0 Then MsgBox "사용법 파일 경로가 거래처 목록에 없습니다. [폴더 연결 맞추기]를 한 번 누르세요.", vbExclamation, "사용법": Exit Sub
    CreateObject("WScript.Shell").Run Chr(34) & gHelp & Chr(34), 1, False
End Sub
