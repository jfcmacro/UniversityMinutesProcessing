# Author: Juan Francisco Cardona McCormick
# email: jfcmacro@gmail.com
#
# date: 2020/03/10
# Purpose: Show the arguments
param(
    [switch]$Activity,
    [switch]$CreateDirectories,
    [string]$DirectoryName = "$env:userprofile\AppData\Local\Temp"
)

function ShowOutput($switch,$value) {
    if ($value) {
	Write-Host "$switch is on"
    }
    else {
	Write-Host "$switch is off"
    }
}

ShowOutput "Activity" $Activity
ShowOutput "CreateDirectories" $CreateDirectories
Write-Host "Directory name: $DirectoryName"

