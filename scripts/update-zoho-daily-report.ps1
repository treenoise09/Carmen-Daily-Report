param(
  [Parameter(Mandatory = $true)]
  [string]$TemplatePath,

  [Parameter(Mandatory = $true)]
  [string]$InputDir,

  [Parameter(Mandatory = $true)]
  [string]$OutputDir,

  [string]$TodayExport,
  [string]$OpenAllExport
)

$ErrorActionPreference = "Stop"

$xlSrcRange = 1
$xlYes = 1
$xlDatabase = 1
$xlFormulas = -4123
$xlPart = 2
$xlByRows = 1
$xlByColumns = 2
$xlPrevious = 2

function Release-Com($obj) {
  if ($null -ne $obj) {
    try {
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj)
    } catch {
      # ignore release error
    }
  }
}

function Get-LastUsedRow($ws) {
  $cell = $ws.Cells.Find("*", $ws.Cells.Item(1,1), $xlFormulas, $xlPart, $xlByRows, $xlPrevious)
  if ($null -eq $cell) { return 1 }
  return $cell.Row
}

function Get-LastUsedCol($ws) {
  $cell = $ws.Cells.Find("*", $ws.Cells.Item(1,1), $xlFormulas, $xlPart, $xlByColumns, $xlPrevious)
  if ($null -eq $cell) { return 1 }
  return $cell.Column
}

function Find-HeaderCell($ws) {
  $maxRows = [Math]::Min((Get-LastUsedRow $ws), 40)
  $maxCols = [Math]::Min((Get-LastUsedCol $ws), 40)

  for ($r = 1; $r -le $maxRows; $r++) {
    for ($c = 1; $c -le $maxCols; $c++) {
      $text = ([string]$ws.Cells.Item($r, $c).Text).Trim()

      if ($text -eq "Support Member Assigned") {
        return @{
          Row = $r
          Col = $c
        }
      }
    }
  }

  throw "Cannot find header 'Support Member Assigned' in worksheet '$($ws.Name)'"
}

function Find-FooterRow($ws, $startRow, $lastRow, $lastCol) {
  for ($r = $startRow; $r -le $lastRow; $r++) {
    for ($c = 1; $c -le $lastCol; $c++) {
      $text = ([string]$ws.Cells.Item($r, $c).Text).Trim()

      if ($text -like "Total records in this page*") {
        return $r
      }
    }
  }

  return $null
}

function Read-ZohoExport($excel, $filePath) {
  if (!(Test-Path $filePath)) {
    throw "Export file not found: $filePath"
  }

  Write-Host "Opening export: $filePath"

  $wb = $null
  $ws = $null

  try {
    $wb = $excel.Workbooks.Open($filePath)
    $ws = $wb.Worksheets.Item(1)

    $header = Find-HeaderCell $ws

    $lastUsedRow = Get-LastUsedRow $ws
    $lastUsedCol = Get-LastUsedCol $ws

    $footerRow = Find-FooterRow $ws $header.Row $lastUsedRow $lastUsedCol

    if ($null -ne $footerRow) {
      $lastDataRow = $footerRow - 1
    } else {
      $lastDataRow = $lastUsedRow
    }

    $rowCount = $lastDataRow - $header.Row + 1
    $colCount = $lastUsedCol - $header.Col + 1

    if ($rowCount -lt 2) {
      throw "No data rows found in: $filePath"
    }

    $range = $ws.Range(
      $ws.Cells.Item($header.Row, $header.Col),
      $ws.Cells.Item($lastDataRow, $lastUsedCol)
    )

    $values = $range.Value2

    return @{
      Values = $values
      Rows = $rowCount
      Cols = $colCount
    }
  }
  finally {
    if ($null -ne $wb) {
      try { $wb.Close($false) } catch {}
    }

    Release-Com $ws
    Release-Com $wb
  }
}

function Write-DataSheet($wb, $sheetName, $data, $preferredTableName) {
  Write-Host "Writing sheet: $sheetName"

  $ws = $wb.Worksheets.Item($sheetName)

  $existingTableName = $preferredTableName

  if ($ws.ListObjects.Count -gt 0) {
    try {
      $existingTableName = $ws.ListObjects.Item(1).Name
      $ws.ListObjects.Item(1).Unlist()
    } catch {
      Write-Host "Could not unlist existing table on $sheetName. Continue with clear range."
    }
  }

  $ws.Cells.Clear()

  $target = $ws.Range(
    $ws.Cells.Item(1, 1),
    $ws.Cells.Item($data.Rows, $data.Cols)
  )

  $target.Value2 = $data.Values

  try {
    $table = $ws.ListObjects.Add($xlSrcRange, $target, $null, $xlYes)
    $table.Name = $existingTableName
    $table.TableStyle = "TableStyleMedium2"
  } catch {
    Write-Host "Could not create Excel table on $sheetName. Continue as plain range."
  }

  $ws.Columns.AutoFit() | Out-Null

  Release-Com $target
  Release-Com $ws
}

function Refresh-Pivots($wb, $openData, $todayData) {
  Write-Host "Refreshing and validating pivot tables..."
  foreach ($ws in $wb.Worksheets) {
    $pivotTables = $ws.PivotTables()
    for ($i = 1; $i -le $pivotTables.Count; $i++) {
      $pt = $pivotTables.Item($i)
      $cache = $null
      try {
        $sourceText = [string]$pt.SourceData
        if ($sourceText -match 'Ticket_Today') {
          $data = $todayData
          $sourceSheet = 'Ticket_Today'
        } elseif ($sourceText -match 'Ticket_OpenAll') {
          $data = $openData
          $sourceSheet = 'Ticket_OpenAll'
        } else {
          throw "Unknown source for pivot '$($pt.Name)': $sourceText"
        }
        # The template has section headings immediately below its pivots.
        # Reserve extra rows BEFORE changing the cache, which can resize the pivot.
        if ($pt.RowFields().Count -eq 1) {
          $fieldName = [string]$pt.RowFields().Item(1).SourceName
          $fieldCol = 0
          for ($c = 1; $c -le $data.Cols; $c++) {
            if ([string]$data.Values[1,$c] -eq $fieldName) { $fieldCol = $c; break }
          }
          if (!$fieldCol) { throw "Missing pivot row field: $fieldName" }
          $labels = @{}
          for ($r = 2; $r -le $data.Rows; $r++) { $labels[[string]$data.Values[$r,$fieldCol]] = $true }
          $requiredRows = $labels.Count + 3
          $extraRows = $requiredRows - $pt.TableRange2.Rows.Count
          if ($extraRows -gt 0) {
            $insertAt = $pt.TableRange2.Row + $pt.TableRange2.Rows.Count
            Write-Host "Reserving $extraRows extra rows below '$($pt.Name)'"
            $ws.Rows.Item("${insertAt}:$($insertAt + $extraRows - 1)").Insert() | Out-Null
          }
        }
        $source = "'$sourceSheet'!R1C1:R$($data.Rows)C$($data.Cols)"
        $cache = $wb.PivotCaches().Create($xlDatabase, $source)
        $cache.MissingItemsLimit = 0
        $pt.ChangePivotCache($cache)
        if (!$pt.RefreshTable()) { throw "RefreshTable returned false" }
        $expected = $data.Rows - 1
        $actual = $pt.GetPivotData($pt.DataFields().Item(1).Name).Value2
        if ($cache.RecordCount -ne $expected -or $actual -ne $expected) {
          throw "Pivot total mismatch: expected $expected, actual $actual, cache $($cache.RecordCount)"
        }
        Write-Host "Validated '$($pt.Name)' on '$($ws.Name)': $actual tickets from $sourceSheet"
      } catch {
        throw "Pivot '$($pt.Name)' on '$($ws.Name)' failed: $($_.Exception.Message)"
      } finally {
        Release-Com $cache
        Release-Com $pt
      }
    }
    Release-Com $pivotTables
  }
}

function Get-InputExports($inputDir) {
  if (!(Test-Path $inputDir)) {
    throw "Input directory not found: $inputDir"
  }

  $files = @(
    Get-ChildItem -Path $inputDir -File |
      Where-Object {
        $_.Extension -in ".xls", ".xlsx" -and
        $_.Name -notlike "~$*"
      } |
      Sort-Object LastWriteTime -Descending
  )

  if ($files.Count -lt 2) {
    throw "Need at least 2 Excel export files in: $inputDir"
  }

  $sortedBySize = @($files | Sort-Object Length -Descending)

  $openAllExport = $sortedBySize[0].FullName
  $todayExport = $sortedBySize[1].FullName

  Write-Host "Auto selected OpenAll export: $openAllExport"
  Write-Host "Auto selected Today export:   $todayExport"

  return @{
    OpenAll = $openAllExport
    Today = $todayExport
  }
}

function Clear-InputFolder($inputDir) {
  Write-Host "Cleaning input folder..."

  $files = @(
    Get-ChildItem -Path $inputDir -File |
      Where-Object {
        $_.Extension -in ".xls", ".xlsx" -and
        $_.Name -notlike "~$*"
      }
  )

  foreach ($file in $files) {
    try {
      Remove-Item $file.FullName -Force
      Write-Host "Deleted input file: $($file.Name)"
    } catch {
      Write-Host "WARN: Cannot delete input file: $($file.FullName)"
      Write-Host "      $($_.Exception.Message)"
    }
  }
}

$excel = $null
$templateWb = $null

try {
  if ($TodayExport -or $OpenAllExport) {
    if (!$TodayExport -or !$OpenAllExport) {
      throw "Both TodayExport and OpenAllExport are required together."
    }
  } else {
    $exports = Get-InputExports $InputDir
    $OpenAllExport = $exports.OpenAll
    $TodayExport = $exports.Today
  }

  if (!(Test-Path $TemplatePath)) {
    throw "Template not found: $TemplatePath"
  }

  if (!(Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
  }

  $dateText = Get-Date -Format "dd_MM_yyyy"
  $outputPath = Join-Path $OutputDir "Tickets_AllZoho$dateText.xls"

  $templateFullPath = [System.IO.Path]::GetFullPath($TemplatePath)
  $outputFullPath = [System.IO.Path]::GetFullPath($outputPath)

  # Build separately, so failures never replace a working report and same-day reruns work.
  $stagingPath = Join-Path $OutputDir (".zoho-" + [guid]::NewGuid().ToString('N') + '.xls')

  Write-Host "Copy template:"
  Write-Host "  From: $TemplatePath"
  Write-Host "  To:   $outputPath"

  Copy-Item -LiteralPath $TemplatePath -Destination $stagingPath

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false

  Write-Host "Reading Today export..."
  $todayData = Read-ZohoExport $excel $TodayExport

  Write-Host "Reading OpenAll export..."
  $openAllData = Read-ZohoExport $excel $OpenAllExport

  Write-Host "Opening output workbook..."
  $templateWb = $excel.Workbooks.Open($stagingPath)

  Write-DataSheet $templateWb "Ticket_Today" $todayData "tbl_Ticket_Today"
  Write-DataSheet $templateWb "Ticket_OpenAll" $openAllData "tbl_Ticket_OpenAll"

  Refresh-Pivots $templateWb $openAllData $todayData

  Write-Host "Calculating workbook..."
  $excel.CalculateFullRebuild()

Write-Host "Saving: $outputPath"
$templateWb.Save()
$templateWb.Close($true)
$templateWb = $null

if (Test-Path -LiteralPath $outputPath) {
  $backupPath = "$outputPath.backup-$(Get-Date -Format 'yyyyMMdd-HHmmss-fff')"
  Copy-Item -LiteralPath $outputPath -Destination $backupPath
  Write-Host "Previous report backed up: $backupPath"
}
Move-Item -LiteralPath $stagingPath -Destination $outputPath -Force
Write-Host "Saved report: $outputPath"

if (!$PSBoundParameters.ContainsKey('TodayExport')) {
  Clear-InputFolder $InputDir
}
}
finally {
  if ($null -ne $templateWb) {
    try { $templateWb.Close($false) } catch {}
    Release-Com $templateWb
  }

  if ($null -ne $excel) {
    try { $excel.Quit() } catch {}
    Release-Com $excel
  }

  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

