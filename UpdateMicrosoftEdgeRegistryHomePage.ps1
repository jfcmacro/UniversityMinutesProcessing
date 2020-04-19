# $registryPath = 'HKEY_CURRENT_USER\SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppContainer\Storage\microsoft.microsoftedge_8wekyb3d8bbwe\MicrosoftEdge'
$registryPath = 'HKCU:\Software\Microsoft\MicrosoftEdge'
$name = "HomeButtonPage"
$value = "http://boca.poligran.edu.co/"
# New-Item-Property -Path $registryPath -Name $name
Test-Path $registryPath
New-ItemProperty -Path $registryPath -Name $name -Value $value -PropertyType String -Force | Out-Null
