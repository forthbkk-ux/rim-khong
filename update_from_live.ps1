$csvPath = Join-Path $pwd.Path "current_live.csv"
$jsonPath = Join-Path $pwd.Path "villages.json"
$jsPath = Join-Path $pwd.Path "villages-data.js"

$lines = [System.IO.File]::ReadAllLines($csvPath, [System.Text.Encoding]::UTF8)
$header = $lines[0].Split(',')

$list = New-Object System.Collections.Generic.List[Object]
for ($i = 1; $i -lt $lines.Length; $i++) {
    $line = $lines[$i]
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $cols = $line.Split(',')
    if ($cols.Length -ge 7) {
        $obj = [ordered]@{
            $header[0] = $cols[0]
            $header[1] = $cols[1]
            $header[2] = $cols[2]
            $header[3] = $cols[3]
            $header[4] = $cols[4]
            $header[5] = $cols[5]
            $header[6] = $cols[6]
        }
        $list.Add($obj)
    }
}

$json = $list | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($jsonPath, $json, [System.Text.Encoding]::UTF8)
$jsContent = "const VILLAGES_DATA = " + $json + ";"
[System.IO.File]::WriteAllText($jsPath, $jsContent, [System.Text.Encoding]::UTF8)
Write-Host "Success: $($list.Count) villages written from latest Google Sheet."
