# 클래식 Outlook 자동 시작(메일 자동 분류 ClientSort 용) - setup.py 가 경로를 채워 _메일분류\ 에 만든다
# '최소화로 실행' 바로가기로 켜면 Outlook 이 VBA 를 불러오지 않아 분류가 안 돈다.
# 그래서 보통으로 켜고, ClientSort 가 시작 기록(last_run.txt)을 남기면(최대 3분) 창을 최소화한다.
# 바로가기 대상: powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "<이 파일>"
$ErrorActionPreference = 'SilentlyContinue'
$outlook = '{{OUTLOOK_EXE}}'
$lastRun = '{{CONF_DIR}}\last_run.txt'

if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) { exit }       # 이미 켜져 있으면 아무것도 안 함
$before = if (Test-Path $lastRun) { Get-Content $lastRun -Raw } else { '' }
Start-Process $outlook

Add-Type -Namespace OlAuto -Name W -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr h, int c); [DllImport("user32.dll")] public static extern bool IsIconic(System.IntPtr h);'
for ($i = 0; $i -lt 90; $i++) {                                      # 2초 x 90 = 3분
    Start-Sleep -Seconds 2
    $now = if (Test-Path $lastRun) { Get-Content $lastRun -Raw } else { '' }
    if ($now -ne $before) { break }
}
# 시작 매크로는 Outlook 이 창을 다 띄우기 전에 돈다 - 창이 자리 잡을 때까지 기다렸다가 최소화하고, 다시 펴지면 한 번 더(최대 30초)
Start-Sleep -Seconds 8
for ($k = 0; $k -lt 6; $k++) {
    $p = Get-Process OUTLOOK -ErrorAction SilentlyContinue | Where-Object MainWindowHandle -ne 0 | Select-Object -First 1
    if ($p -and -not [OlAuto.W]::IsIconic($p.MainWindowHandle)) { [OlAuto.W]::ShowWindow($p.MainWindowHandle, 6) | Out-Null }   # 6 = 최소화
    Start-Sleep -Seconds 5
}
