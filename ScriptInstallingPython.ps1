$disk=gwmi win32_diskdrive | ?{$_.interfacetype -eq "USB"} | %{gwmi -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID=`"$($_.DeviceID.replace('\','\\'))`"} WHERE AssocClass = Win32_DiskDriveToDiskPartition"} |  %{gwmi -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID=`"$($_.DeviceID)`"} WHERE AssocClass = Win32_LogicalDiskToPartition"} | %{$_.deviceid}
$marathonDoc=Join-Path -Path "$disk" -ChildPath "maraton"
Set-Location (Get-ChildItem Env:\HOMEDRIVE).Value

# Installing Python 3.6.2
$execDir='python-software'
$pythonExec='python-3.6.2-amd64.exe'
$pythonInstallPath=Join-Path -Path $marathonDoc -ChildPath $execDir | Join-Path -ChildPath $pythonExec

Write-Host ('Installing python: ' + $pythonInstallPath)
& $pythonInstallPath /quiet 'InstallAllUser=0' 'PrependPath=1' 