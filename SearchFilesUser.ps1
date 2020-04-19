﻿[String]$username = "CINFO\fcardona"
[String]$outfile = "C:\Users\fcardona\output.txt"
$path = Get-ChildItem "C:\Users\fcardona" -Recurse
Foreach( $file in $path ) {
    $f = Get-Acl $file.FullName
    if ( $f.Owner -eq $username ) {
        Write-Host( "{0}"-f $file.FullName | Out-File -Encoding utf8 -FilePath $outfile -Append)
    }
}
& notepad $outfile