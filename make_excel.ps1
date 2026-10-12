$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$csvPath = Join-Path $pwd.Path "live_sheet_clean.csv"
$workbook = $excel.Workbooks.Open($csvPath)
$sheet = $workbook.Worksheets.Item(1)

$fixesJsonText = [System.IO.File]::ReadAllText((Join-Path $pwd.Path "fixes.json"), [System.Text.Encoding]::UTF8)
$fixes = $fixesJsonText | ConvertFrom-Json

$usedRows = $sheet.UsedRange.Rows.Count
Write-Host "Excel rows count: $usedRows"

# Header column 8
$headerCell = $sheet.Cells.Item(1, 8)
$headerCell.Value2 = "Status / Notes"
$headerRange = $sheet.Range("A1", "H1")
$headerRange.Font.Bold = $true
$headerRange.Interior.Color = 0xE0E0E0

for ($r = 2; $r -le $usedRows; $r++) {
    $idCell = $sheet.Cells.Item($r, 1)
    $id = [string]$idCell.Text
    $id = $id.Trim()
    
    $fix = $fixes.$id
    if ($fix) {
        $cellLat = $sheet.Cells.Item($r, 6)
        $cellLng = $sheet.Cells.Item($r, 7)
        $cellNote = $sheet.Cells.Item($r, 8)
        
        $cellLat.Value2 = [string]$fix.lat
        $cellLng.Value2 = [string]$fix.lng
        $cellNote.Value2 = [string]($fix.reason)
        
        # Color RED (0x0000FF in BGR)
        $cellLat.Font.Color = 0x0000FF
        $cellLat.Font.Bold = $true
        $cellLat.Interior.Color = 0xD8E8FF # Light pink/red highlight
        
        $cellLng.Font.Color = 0x0000FF
        $cellLng.Font.Bold = $true
        $cellLng.Interior.Color = 0xD8E8FF
        
        $cellNote.Font.Color = 0x0000FF
        $cellNote.Font.Bold = $true
        
        Write-Host "Row $r (ID $id): Updated coordinates with RED font"
    }
}

$sheet.Columns.AutoFit() | Out-Null

$outXlsx = Join-Path $pwd.Path "Mekong_Villages_Fixed_RedText.xlsx"
if (Test-Path $outXlsx) { Remove-Item $outXlsx -Force }

$workbook.SaveAs($outXlsx, 51)
$workbook.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Successfully generated Excel file: $outXlsx"