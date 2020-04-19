function set-shortcut {
param ( [string]$SourceLnk, [string]$DestinationPath )
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($SourceLnk)
    $Shortcut.TargetPath = $DestinationPath
    $Shortcut.Save()
    }

$disk=gwmi win32_diskdrive | ?{$_.interfacetype -eq "USB"} | %{gwmi -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID=`"$($_.DeviceID.replace('\','\\'))`"} WHERE AssocClass = Win32_DiskDriveToDiskPartition"} |  %{gwmi -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID=`"$($_.DeviceID)`"} WHERE AssocClass = Win32_LogicalDiskToPartition"} | %{$_.deviceid}
$marathonDoc=Join-Path -Path "$disk" -ChildPath "maraton"
Set-Location (Get-ChildItem Env:\HOMEDRIVE).Value
$userDir=[Io.Path]::Combine('C:\',(Get-ChildItem Env:\HOMEPATH).Value, "Desktop")
Set-Location $userDir

# C++ Documentation
$docDir='cpp-documentation'
$docName='html-book-20180819.zip' 

$zipFile=(Join-Path -Path $marathonDoc -ChildPath $docDir | Join-Path -ChildPath $docName)
$destDir=Join-Path -Path (Get-Location) -ChildPath $docDir
$shortcutFile=Join-Path -Path (Get-Location) -ChildPath 'Cpp-Docs.lnk'
$shortcutDest=Join-Path -Path (Get-Location) -ChildPath $docDir | Join-Path -ChildPath 'reference' | Join-Path -ChildPath 'en' | Join-Path -ChildPath 'Main_Page.html'

if (-not(Test-Path -Path $destDir)) {
    New-Item -ItemType Directory -Path $destDir
    Expand-Archive $zipFile -DestinationPath $destDir
}

if (-not(Test-Path -Path $shortcutFile)) {
    set-shortcut $shortcutFile $shortcutDest
}

# Java Documentation
$docDir='java-documentation'
$docName='jdk-8u221-docs-all.zip'

$zipFile=(Join-Path -Path $marathonDoc -ChildPath $docDir | Join-Path -ChildPath $docName)
$destDir=Join-Path -Path (Get-Location) -ChildPath $docDir
$shortcutFile=Join-Path -Path (Get-Location) -ChildPath 'Java-Docs.lnk'
$shortcutDest=Join-Path -Path (Get-Location) -ChildPath $docDir | Join-Path -ChildPath 'docs' | Join-Path -ChildPath 'api' | Join-Path -ChildPath 'index.html'

if (-not(Test-Path -Path $destDir)) {
    New-Item -ItemType Directory -Path $destDir
    Expand-Archive $zipFile -DestinationPath $destDir
}

if (-not(Test-Path -Path $shortcutFile)) {
    set-shortcut $shortcutFile $shortcutDest
}

# Python Documentation
$docDir='python-doc'
$docName='python-3.6.2-docs-html.zip' 

$zipFile=(Join-Path -Path $marathonDoc -ChildPath $docDir | Join-Path -ChildPath $docName)
$destDir=Join-Path -Path (Get-Location) -ChildPath $docDir
$shortcutFile=Join-Path -Path (Get-Location) -ChildPath 'Python-Docs.lnk'
$shortcutDest=Join-Path -Path (Get-Location) -ChildPath $docDir | Join-Path -ChildPath 'python-3.6.2-docs-html' | Join-Path -ChildPath 'index.html'

if (-not(Test-Path -Path $destDir)) {
    New-Item -ItemType Directory -Path $destDir
    Expand-Archive $zipFile -DestinationPath $destDir
}

if (-not(Test-Path -Path $shortcutFile)) {
    set-shortcut $shortcutFile $shortcutDest
}

## Installing Python 3.6.2
#$execDir='python-software'
#$pythonExec='python-3.6.2-amd64.exe'
#$pythonInstallPath=Join-Path -Path $marathonDoc -ChildPath $execDir | Join-Path -ChildPath $pythonExec

#Write-Host ('Installing python: ' + $pythonInstallPath)
#& $pythonInstallPath /quiet 'InstallAllUser=0' 'PrependPath=1' 


