param(
  [Parameter(Mandatory = $true)][string]$File,
  [int]$MaxRows = 3000
)
# autofit_rows.ps1 - openpyxl 로 만든 조서의 줄바꿈 칸 행을 엑셀이 실제로 맞추게 한다(Rows.AutoFit).
# 줄이 MaxRows 보다 많은 시트(Data_분개장 등)는 건너뛴다. 병합 칸은 이어서 excel-theme tools/fit_merged_rows.ps1 이 맞춘다.
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.ScreenUpdating = $false
try {
  $wb = $xl.Workbooks.Open((Resolve-Path -LiteralPath $File).Path)
  if ($wb.ReadOnly) { Write-Output ("SKIP (open elsewhere): " + $File); $wb.Close($false); return }
  $n = 0
  foreach ($ws in $wb.Worksheets) {
    $ur = $ws.UsedRange
    if ($ur.Rows.Count -gt $MaxRows) { continue }
    [void]$ur.Rows.AutoFit()
    $n++
  }
  $wb.Save()
  $wb.Close($false)
  Write-Output ("OK  autofit " + $n + " sheets - " + $File)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
