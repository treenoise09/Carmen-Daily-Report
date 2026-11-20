param([string]$ReportPath)
$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null
try {
  $wb = $excel.Workbooks.Open($ReportPath, 0, $true)
  $wb.RefreshAll()
  $excel.CalculateUntilAsyncQueriesDone()
  foreach ($ws in $wb.Worksheets) {
    foreach ($pt in $ws.PivotTables()) {
      if (!$pt.RefreshTable()) { throw "Refresh failed: $($pt.Name)" }
      $actual = $pt.GetPivotData($pt.DataFields().Item(1).Name).Value2
      if ($actual -ne $pt.PivotCache().RecordCount) { throw "Total mismatch: $($pt.Name)" }
      Write-Host "Refresh again OK: $($pt.Name) = $actual"
    }
  }
  $excel.CalculateFullRebuild()
} finally {
  if ($wb) { $wb.Close($false) }
  $excel.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
