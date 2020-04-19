function Clear-Null-Files($dir) {
    Push-Location $dir
    Get-ChildItem -Recurse | ForEach-Object  {
        if ($_.Name -eq "0" -or $_.Name -eq "9") {
            Write-Host "Removing: ",$_.FullName
            Remove-Item $_.FullName
        }
    }
    Pop-Location
}

Set-Location ~
Clear-Null-Files "Documents\Beatriz\ACTAS CONSEJO ESCUELA"