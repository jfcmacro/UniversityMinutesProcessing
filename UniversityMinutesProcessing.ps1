# Author: Juan Francisco Cardona Mc'Cormick
# email: jfcmacro@gmail.com
# date: 2019/09/13
# Purpose: 

function ParserDate($date) {
    if ($date -match "(.*)( *)(Ho|Pr|Al|As|Co).*$") {
        $date = $Matches[1]
    }
    return $date
}

function ParserDateWith($date, $withText) {
    if ($date -match "(.*)( *)$($withText).*$") {
       $date = $Matches[1]
    }
    return $date
}

function ParserDate2($date) {
    if ($date -match "(.*)( *)LUG.*$") {
       $date = $Matches[1]
    }
    return $date
}

function ParserDate3($date) {
    if ($date -match "(.*)( *)Esta.*$") {
       $date = $Matches[1]
    }
    return $date
}

function ParserDate4($date) {
    if ($date -match "(.*)( *)Presi.*$") {
       $date = $Matches[1]
    }
    return $date
}

function ParserDate5($date) {
    if ($date -match "(.*)( *)\..*$") {
       $date = $Matches[1]
    }
    return $date
}

function ParserDate6($date) {
    if ($date -match "(.*)( *)Reuni.*$") {
       $date = $Matches[1]
    }
    return $date
}

function ParserDate7($date) {
    if ($date -match "(.*)( *)Consu.*$") {
       $date = $Matches[1]
    }
    return $date
}

function ParserDate8($date) {
    if ($date -match "(.*)( *)Dici$") {
       $date = $Matches[1]
    }
    return $date
}

function ParserDate9($date) {
    if ($date -match "(.*)( *)H$") {
       $date = $Matches[1]
    }
    return $date
}

function ReduceFromYear($date) {
    return [string]$date.Replace("del año", "de")
}

function AllDateParse($date) {
    $date = ParserDate9(ParserDate8(ParserDate7(ParserDate6(ParserDate5(ParserDate4(ParserDate3(ParserDate2(ParserDate($dateActa)))))))))
    $date = ReduceFromYear($date)
    return FinalDateTrans($date)
}

function FinalDateTrans($date) {
    switch([string]$date.Trim()) {
        "Agosto 13 y 14  de 1999"        { $res = "Agosto 13 de 1999"; break }
        "Septiembre 10 y Octubre 9 de 2" { $res = "Septiembre 10 de 2003"; break }
        "Noviembre  13 y 20 de 2008"     { $res = "Noviembre 13 de 2008"; break }
        "Bloque 26- 106 S"               { $res = "Enero 14 de 2009"; break }
        "Jueves 16 de Julio de 2009"     { $res = "Julio 16 de 2009"; break }
        "Jueves 30 Julio de 2009"        { $res = "Julio 30 de 2009"; break }
        "Martes 17 de Noviembre de 2009" { $res = "Noviembre 17 de 2009"; break }
        "Jueves 19 de Febrero de 2010"   { $res = "Febrero 19 de 2010"; break }
        "Jueves 19 de Noviembre 23 de 2" { $res = "Noviembre 19 de 2010"; break }
        "Lunes 6 de Diciembre  de 2010"  { $res = "Diciembre 6 de 2010"; break }
        "Mayo 28  y Junio 4 de 2012"     { $res = "Mayo 29 de 2012"; break }
        "Abril 4"                        { $res = "Abril 4 de 2013"; break }
        "Junio 5"                        { $res = "Junio 5 de 2013"; break }
        "Julio 15"                       { $res = "Julio 15 de 2013"; break }
        "Abril 21"                       { $res = "Abril 21 de 2014"; break }
        "Mayo 65 de 2014"                { $res = "Mayo 7 de 2014"; break }
        "Julio 21 y Agosto 11 de 2016"   { $res = "Julio 21 de 2016"; break }
        "Agosto 1° de 2017"              { $res = "Agosto 1 de 2017"; break }
        default                          { $res = $date; break }
    }
    return $res
}

function getStringNameFromFileName($filename) {
    $strName = ([string]$filename).Substring(4)

    if ($strName.EndsWith(".doc")) {
        $strName = $strName.Substring(0, $strName.Length - 4)
    }
    
    if ($strName.EndsWith(".docx")) {
        $strName = $strName.Substring(0, $strName.Length - 5)
    }

    if ($strName -match " *([0-9]+)?(.*)") {
        $strName = $Matches[2]
    }

    if (([string]$strName).Length -gt 0) {
        if ($strName -match " *(\( *[0-9]+ *\))(.*)") {
            $strName = $Matches[2]
        }
    }

    if (([string]$strName).Length -gt 0) {
        if ($strName -match "(.*)(\(? *[0-9]+ *\)?)") {
            $strName = $Matches[1]
        }
    }

    if (([string]$strName).Length -gt 0) {
        if ($strName -match "(.*)(\(?[0-9]+)") {
            $strName = $Matches[1]
        }
    }

    if (([string]$strName).Length -gt 0) {
        if ($strName -match "(.*)([0-9]$)") {
            $strName = $Matches[1]
        }
    }

    if (([string]$strName).Length -gt 0) {
        $strName = ([string]$strName).Trim().ToLower().Replace(' ', '-')
    }

    return $strName
}

function Generate-New-Name ($stringDate, $seqFormatNumber, $str) {
    $newName = $stringDate + "-consejo-de-escuela-" + $seqFormatNumber

    if (([string]$str).Length -gt 0) {
        $newName = $newName + "-" + $str
    }

    return $newName
}

function Get-Date-From-Minute-File ($wordApp, $year, $obj,[ref]$mla) {
    
    $findText = "Fecha:(\t| )*"
    $charactersAround = 30
    $fileName = $obj.Name
    if (($obj.Name -match "^(a|A)(c|C)(t|T)(a|A).*(d|D)(o|O)(c|C)(x|X)?$") -and ($obj.Name -notmatch "(a|A)(n|N)(e|E)(x|X)(o|O)")) {

        $document = $wordApp.Documents.open($obj.FullName)
        $range = $document.content

        if ($range.Text -match "$($findText)(.{$($charactersAround)})") {
            $dateActa = $Matches[2]
            # $dateActa2 = ParserDate9(ParserDate8(ParserDate7(ParserDate6(ParserDate5(ParserDate4(ParserDate3(ParserDate2(ParserDate($dateActa)))))))))
            # $dateActa2 = ReduceFromYear($dateActa2)
            # $dateActa2 = FinalDateTrans($dateActa2)
            $dateActa2 = AllDateParse($dateActa2)
            Write-Host("FileName: " + $fileName)
            $fileStringName = getStringNameFromFileName($fileName)
            # Write-Host -NoNewline ('Year: ' + $year + ' Found Date: ' + $dateActa +' Date reduced: ' + $dateActa2)  
            $formatDate = [datetime]::Parse($dateActa2)
            $stringDate = $formatDate.toString("yyyy-MM-dd")
            # Write-Host (' Parsed: ' + $stringDate)
            $formatNumber = ""
            if (($year -ge 1998) -and ($year -lt 2007)) {
                if ($fileName -match "[0-9]+") {
                    $actaNumberStr = $Matches[0]
                    $actaNumber = [int]$actaNumberStr
                    $mla.Value.Add($actaNumber) | Out-Null
                    $formatNumber = '{0:d3}' -f $actaNumber
                    # $seqFormatNumber = '{0:d3}' -f $minuteNumber.value
                    # Write-Host ('FormatNumber: ' + $formatNumber)
                }
            }
            else {
                if ($fileName -match " *\( *([0-9]+)\)") {
                    $actaNumberStr = $Matches[1]
                    $actaNumber = [int]$actaNumberStr
                    $formatNumber = '{0:d3}' -f $actaNumber
                    # $seqFormatNumber = '{0:d3}' -f $minuteNumber.value
                    # Write-Host ('FormatNumber: ' + $formatNumber)
                }
            }
            $newName = Generate-New-Name $stringDate  $formatNumber $fileStringName
            Write-Host ('New name: ' + $newName)
        }
        else {
            ## Write-Host ('Fichero no reconocido: ' + $fileName)
        }
        # $formatDate = [datetime]::Parse($dateActa)
        # $stringDate = $formatDate.toString("yyyy-MM-dd")
        # $pdf_filename = "$($_.DirectoryName)\$($stringDate)-consejo-de-escuela-$($formatNumber).pdf"
        # Write-Host $pdf_filename
        # $document.SaveAs([ref] $pdf_filename, [ref] 17)
        $document.Close()
    }
}

function ContainsActa ($fileName) {
    return ($fileName -match "(a|A)(c|C)(t|T)(a|A)")
}


function GetYear ($dirName) {
   $year = ""
    if ($dirName -match "([0-9]+)") {
        $year = $Matches[0]
    }
    return [int]$year
}

# main
$wordApp = New-Object -ComObject Word.Application
cd ~
cd "Documents\Beatriz\ACTAS CONSEJO ESCUELA"
$minuteNumber = 0
$minuteListArray = New-Object System.Collections.ArrayList

Get-ChildItem -Directory | ForEach-Object {

    $dirName = $_.Name
    if (ContainsActa($dirName)) {

        $year = GetYear($dirName)
        
        Push-Location $dirName
        #- Write-Host (" Year: " + $year)

        Get-ChildItem -File | ForEach-Object {
            Get-Date-From-Minute-File $wordApp $year $_ ([ref] $minuteListArray)
        }

        Get-ChildItem -Directory | ForEach-Object {

            $dirName2 = $_.Name
            if (ContainsActa($dirName2)) {

                Push-Location $dirName2
                ## Write-Host ("Directory2: " + $dirName2)

                Get-ChildItem -File | ForEach-Object {
                    Get-Date-From-Minute-File $wordApp $year $_ ([ref] $minuteListArray)
                }

                Get-ChildItem -Directory | ForEach-Object {

                    $dirName3 = $_.Name
                    if (ContainsActa($dirName3)) {

                        Push-Location $dirName3
                   ##     Write-Host ("Directory3: " + $dirName3)
                        Get-ChildItem -File | ForEach-Object {
                            Get-Date-From-Minute-File $wordApp $year $_ ([ref] $minuteListArray)
                        }
                        Pop-Location
                    }
                }
                Pop-Location
            }
        }
        Pop-Location
    }
}

for ($i = 0; $i -eq 130; $i++) {
   if (-Not ($minuteListArray.Contains($i))) {
       Write-Host ('Minute ' + $i + "doesn't exist")
   }
}
$wordApp.Quit()
cd ~ 
# Clear-Host