param([string]$ReportPath, [string]$ExportDir, [string]$AuditPath, [string]$SnapshotPath)
$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$result = @()
try {
  foreach ($file in @($ReportPath, (Join-Path $ExportDir 'TicketToday_Support.xls'), (Join-Path $ExportDir 'OpenAllNew.xls'))) {
    $wb = $excel.Workbooks.Open($file, 0, $true)
    try {
      $sheets = @()
      foreach ($ws in $wb.Worksheets) {
        $used = $ws.UsedRange
        $values = $used.Value2
        $rows = @()
        for ($r = 1; $r -le $used.Rows.Count; $r++) {
          $row = @()
          for ($c = 1; $c -le $used.Columns.Count; $c++) { $row += $values[$r,$c] }
          $rows += ,$row
        }
        $pivots = @()
        foreach ($pt in $ws.PivotTables()) {
          $fields = @()
          foreach ($field in $pt.PivotFields()) {
            if ($field.Orientation -ne 0) {
              $items = @()
              try { foreach ($item in $field.PivotItems()) { $items += @{name=$item.Name; visible=$item.Visible} } } catch {}
              $fields += @{name=$field.Name; source=$field.SourceName; orientation=$field.Orientation; items=$items}
            }
          }
          $pivots += @{name=$pt.Name; source=[string]$pt.SourceData; range=$pt.TableRange2.Address(); records=$pt.PivotCache().RecordCount; fields=$fields}
        }
        $merges = @()
        if ($ws.Name -eq 'Open_All') {
          foreach ($cell in $used.Cells) {
            if ($cell.MergeCells -and $cell.Address() -eq $cell.MergeArea.Cells.Item(1,1).Address()) { $merges += @{range=$cell.MergeArea.Address(); text=$cell.Text} }
          }
        }
        $sheets += @{name=$ws.Name; address=$used.Address(); rows=$rows; pivots=$pivots; merges=$merges}
      }
      $result += @{file=$file; sheets=$sheets}
      if ($SnapshotPath -and $file -eq $ReportPath) { $wb.SaveAs($SnapshotPath, 51) }
    } finally { $wb.Close($false) }
  }
  $result | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $AuditPath -Encoding UTF8
} finally { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
