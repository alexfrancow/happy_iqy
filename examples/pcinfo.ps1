# Get Basic information
$whoami = whoami
$date = Get-Date -UFormat "%Y/%m/%d-%T"
$os = Get-CimInstance Win32_OperatingSystem | Select -ExpandProperty  Caption
$ip = Test-Connection -ComputerName (hostname) -Count 1  | Select -ExpandProperty IPV4Address
$ip = $ip.IPAddressToString

$user = @{
    whoami=$whoami
    ip=$ip
    os=$os
    date=$date
}

$json = $user | ConvertTo-Json
Invoke-RestMethod 'http://localhost:3000/users/create' -Method Post -Body $json -ContentType 'application/json'