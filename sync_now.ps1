[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
$csvUrl = "https://docs.google.com/spreadsheets/d/17du3Uq-QQppbU1DncxxnYNZbLNhNoBIFsrFZa4QaBZE/export?format=csv"
$csvTarget = Join-Path $pwd.Path "current_live.csv"
$jsonTarget = Join-Path $pwd.Path "villages.json"
$jsTarget = Join-Path $pwd.Path "villages-data.js"

$wc = New-Object System.Net.WebClient
$wc.Encoding = [System.Text.Encoding]::UTF8
$wc.DownloadFile($csvUrl, $csvTarget)

$lines = [System.IO.File]::ReadAllLines($csvTarget, [System.Text.Encoding]::UTF8)
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
[System.IO.File]::WriteAllText($jsonTarget, $json, [System.Text.Encoding]::UTF8)
$jsContent = "const VILLAGES_DATA = " + $json + ";"
[System.IO.File]::WriteAllText($jsTarget, $jsContent, [System.Text.Encoding]::UTF8)
Write-Host "Successfully synced $($list.Count) villages from live Google Sheet."
