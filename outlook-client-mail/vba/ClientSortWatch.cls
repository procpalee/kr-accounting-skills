VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClientSortWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ClientSort - 폴더 하나의 '메일 들어옴' 이벤트를 받아 넘긴다
'   CaseIdx = -1 : 분류 감시 폴더(받은 · 보낸 · #watch) → ClientSort.HandleItem(분류)
'   CaseIdx >= 0 : 거래처 Outlook 폴더 → ClientSort.CaseAdded(그 거래처 메일 저장)
Option Explicit

Public WithEvents Items As Outlook.Items
Public IsSent As Boolean
Public CaseIdx As Long

Private Sub Class_Initialize()
    CaseIdx = -1
End Sub

Private Sub Items_ItemAdd(ByVal Item As Object)
    On Error Resume Next
    If CaseIdx >= 0 Then
        ClientSort.CaseAdded CaseIdx
    Else
        ClientSort.HandleItem Item, IsSent
    End If
End Sub
