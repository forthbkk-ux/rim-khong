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

$fixes = @{
    '327' = @{ lat = '17.882343'; lng = '102.785152' }; # หมู่ 5 สามัคคี ต.หาดคำ อ.เมืองหนองคาย
    '222' = @{ lat = '17.785527'; lng = '102.691951' }; # หมู่ 7 หนองบ่อ ต.ปะโค อ.เมืองหนองคาย
    '429' = @{ lat = '17.962415'; lng = '104.221250' }; # หมู่ 10 แพงกลางใต้ ต.บ้านแพง อ.บ้านแพง
    '432' = @{ lat = '17.961250'; lng = '104.218520' }; # หมู่ 13 แพงกลาง ต.บ้านแพง อ.บ้านแพง
    '434' = @{ lat = '17.963562'; lng = '104.219450' }; # หมู่ 2 แพงกลาง ต.บ้านแพง อ.บ้านแพง
    '520' = @{ lat = '18.058402'; lng = '104.149315' }; # หมู่ 1 ท่าสีไค ต.ดงบัง อ.บึงโขงหลง
    '577' = @{ lat = '18.201508'; lng = '104.052914' }; # หมู่ 1 โคกกว้าง ต.โคกกว้าง อ.บุ่งคล้า
    '769' = @{ lat = '18.376900'; lng = '103.635600' }; # หมู่ 9 นาเหนือ ต.วิศิษฐ์ อ.เมืองบึงกาฬ
    '693' = @{ lat = '18.273950'; lng = '103.858690' }; # หมู่ 1 ชัยพร ต.ชัยพร อ.เมืองบึงกาฬ
    '453' = @{ lat = '17.897996'; lng = '104.224041' }; # หมู่ 3 โคกพะทาย ต.โพนทอง อ.บ้านแพง
    '782' = @{ lat = '18.357499'; lng = '103.362336' }; # หมู่ 9 สรรเสริญ ต.หนองเลิง อ.เมืองบึงกาฬ
}

$list = New-Object System.Collections.Generic.List[Object]
$newCsvLines = New-Object System.Collections.Generic.List[string]
$newCsvLines.Add($lines[0])

for ($i = 1; $i -lt $lines.Length; $i++) {
    $line = $lines[$i]
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $cols = $line.Split(',')
    if ($cols.Length -ge 7) {
        $id = $cols[0].Trim()
        if ($fixes.ContainsKey($id)) {
            $cols[5] = $fixes[$id].lat
            $cols[6] = $fixes[$id].lng
        }
        $newCsvLines.Add(($cols -join ','))
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

[System.IO.File]::WriteAllLines($csvTarget, $newCsvLines, [System.Text.Encoding]::UTF8)
$json = $list | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($jsonTarget, $json, [System.Text.Encoding]::UTF8)
$jsContent = "const VILLAGES_DATA = " + $json + ";"
[System.IO.File]::WriteAllText($jsTarget, $jsContent, [System.Text.Encoding]::UTF8)
Write-Host "Successfully synced $($list.Count) villages from live Google Sheet with coordinate fixes applied."
