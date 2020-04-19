# newSeqInfo
# purpose: create a new custom object with three attributes: Seq, Year, Date
function newSeqInfo($seq,$seqYear,$date,$name,$path,$fileorpath) {
    
    return [pscustomobject]@{Seq=$seq;SeqYear=$seqYear;Date=$date;Name=$name;Path=$path;FileOrPath=$fileorpath}
}

# printSeqInfo
# purpose: prints SeqInfo object atributes
function printSeqInfo($obj) {
    Write-Host "Seq: ",$obj.Seq," SeqYear: ",$obj.SeqYear," Date: ",$obj.Date," Name: ",$obj.Name # ,"Path: ",$obj.Path
}

# ParserDate
# purpose: 
function ParserDate($date) {
    if ($date -match "(.*)( *)(Ho|Pr|Al|As|Co).*$") {
        $date = $Matches[1]
    }
    return $date
}

# ParseDateWith
# purpose:
function ParserDateWith($date, $withText) {
    if ($date -match "(.*)( *)$($withText).*$") {
	    $date = $Matches[1]
    }
    return $date
}

# ParseDate2
# purpose:
function ParserDate2($date) {
    if ($date -match "(.*)( *)LUG.*$") {
	    $date = $Matches[1]
    }
    return $date
}

# ParseDate3
# purpose:
function ParserDate3($date) {
    if ($date -match "(.*)( *)Esta.*$") {
	    $date = $Matches[1]
    }
    return $date
}

# ParseDate4
# purpose:
function ParserDate4($date) {
    if ($date -match "(.*)( *)Presi.*$") {
	    $date = $Matches[1]
    }
    return $date
}

# ParseDate5
# purpose:
function ParserDate5($date) {
    if ($date -match "(.*)( *)\..*$") {
	    $date = $Matches[1]
    }
    return $date
}

# ParseDate6
# purpose:
function ParserDate6($date) {
    if ($date -match "(.*)( *)Reuni.*$") {
	    $date = $Matches[1]
    }
    return $date
}

# ParseDate7
# purpose:
function ParserDate7($date) {
    if ($date -match "(.*)( *)Consu.*$") {
	    $date = $Matches[1]
    }
    return $date
}

# ParseDate8
# purpose:
function ParserDate8($date) {
    if ($date -match "(.*)( *)Dici$") {
	    $date = $Matches[1]
    }
    return $date
}

# ParseDate9
# purpose:
function ParserDate9($date) {
    if ($date -match "(.*)( *)H$") {
	    $date = $Matches[1]
    }
    return $date
}

# ReduceFromYear
# purpose: It takes a year string info and replace from: del año to: de
function ReduceFromYear($date) {
    return [string]$date.Replace("del año", "de")
}

# AllDateParse
# purpose: It takes 
# Error: TODO correct i
function AllDateParse($date) {
    # $date = ParserDate9(ParserDate8(ParserDate7(ParserDate6(ParserDate5(ParserDate4(ParserDate3(ParserDate2(ParserDate($dateActa)))))))))
    $date = ParserDate9(ParserDate8(ParserDate7(ParserDate6(ParserDate5(ParserDate4(ParserDate3(ParserDate2(ParserDate($date)))))))))
    $date = ReduceFromYear($date)
    return FinalDateTrans($date)
}

# FinalDateTrans
# purpose: To adjusts all manually all expected dates formats
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

# getStringNameFromFileName
# 
# purpose: get the name from a filename with extension
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

# Format-New-Name
#
# purpose: this function gives the actual format name for the School (Escuela)
function Format-New-Name($stringDate, $seqFormatNumber, $str) {
    $newName = $stringDate + "-consejo-de-escuela-" + $seqFormatNumber

    if (([string]$str).Length -gt 0) {
        $newName = $newName + "-" + $str
    }

    return $newName
}

# Test-Minute-File
#
# purpose: test if a file is a Minute file (Acta)
function Test-Minute-File($fileName) {
    return (($fileName -match "^(a|A)(c|C)(t|T)(a|A).*(d|D)(o|O)(c|C)(x|X)?$") -and ($fileName -notmatch "(a|A)(n|N)(e|E)(x|X)(o|O)"))
}

# Get-Date-From-Minute-File2
#
# purpose: get the date from Minute file (Word) by using COM in order to get the information of the date write down into the file
function Get-Date-From-Minute-File2($wordApp, $year, $obj) {
    
    $findText = "Fecha:(\t| )*"
    $charactersAround = 30
    $fileName = $obj.Name
    if (($obj.Name -match "^(a|A)(c|C)(t|T)(a|A).*(d|D)(o|O)(c|C)(x|X)?$") -and ($obj.Name -notmatch "(a|A)(n|N)(e|E)(x|X)(o|O)")) {

        $document = $wordApp.Documents.open($obj.FullName)
        $range = $document.content

        if ($range.Text -match "$($findText)(.{$($charactersAround)})") {
            $dateActa2 = $Matches[2]
            $dateActa2 = AllDateParse($dateActa2)
            Write-Host("FileName: " + $fileName)
            $formatDate = [datetime]::Parse($dateActa2)
            Write-Host "Date acta: "+$dateActa2
            $stringDate = $formatDate.toString("yyyy-MM-dd")
            return $stringDate
        }

        $document.Close()
    }
}

# Get-Date-From-Minute-File
#
# purpose: get the date from Minute file (Word) by using COM in order to get the information of the date write down into the file
function Get-Date-From-Minute-File ($wordApp, $year, $obj,[ref]$mla) {
    
    $findText = "Fecha:(\t| )*"
    $charactersAround = 30
    $fileName = $obj.Name
    if (($obj.Name -match "^(a|A)(c|C)(t|T)(a|A).*(d|D)(o|O)(c|C)(x|X)?$") -and ($obj.Name -notmatch "(a|A)(n|N)(e|E)(x|X)(o|O)")) {

        $document = $wordApp.Documents.open($obj.FullName)
        $range = $document.content

        if ($range.Text -match "$($findText)(.{$($charactersAround)})") {
            $dateActa2 = $Matches[2]
            $dateActa2 = AllDateParse($dateActa2)
            Write-Host("FileName: " + $fileName)
            $fileStringName = getStringNameFromFileName($fileName)
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
                }
            }
            $newName = Format-New-Name $stringDate  $formatNumber $fileStringName
            return $newName
        }
        
        $document.Close()
    }
}

# ContainsActa
#
# purpose: tests is a filename has a "[aAcCtTaA]" name inside it
function ContainsActa ($fileName) {
    return ($fileName -match "(a|A)(c|C)(t|T)(a|A)")
}

# GetYear
#
# purpose: gets if a dirnamepath contains a something like a date 19XX - 2000
function GetYear ($dirName) {
    $year = ""
    if ($dirName -match "([0-9]+)") {
        $year = $Matches[0]
    }
    return [int]$year
}

# checkPDFFileFormat
#
# purpose: checks if a PDF file were already written.
function checkPDFFileFormat($fileName) {
    
    if ($fileName -match "^[0-9]{4}-[0-9]{2}-[0-9]{2}.*\.(p|P)(d|D)(f|F)$") {
        if (-not(([String]$fileName).StartsWith("2019"))) {
            Write-Host ("Deleting file: " + $fileName)
            Remove-Item -Path $fileName
        }
    }
}

# deletePDFFilesWithFormat
#
# purpose: erase all PDF Files with a format inside a directory hierarchy
function deletePDFFilesWithFormat($dirBase) {
    Set-Location $dirBase
    
    Get-ChildItem -Directory | ForEach-Object {
        $dirName = $_.Name
        if (containsActa($dirName)) {
            Push-Location $dirName
            Get-ChildItem -File | ForEach-Object {
                checkPDFFileFormat($_.Name)
            }
            Get-ChildItem -Directory | ForEach-Object {
                $dirName2 = $_.Name
                if (containsActa($dirName2)) {
                    Push-Location $dirName2
                    Get-ChildItem -File | ForEach-Object {
                        checkPDFFileFormat($_.Name)
                    }
                    Get-ChildItem -Directory | ForEach-Object {
                        $dirName3 = $_.Name
                        if (containsActa($dirName3)) {
                            Push-Location $dirName3
                            Get-ChildItem -File | ForEach-Object {
                                checkPDFFileFormat($_.Name)
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
}

# analysisLevel1
#
function analysingLevel1([String]$dirBase, [ref]$map) {
    
    $wordApp = New-Object -ComObject Word.Application

    Set-Location $dirBase
    
    Get-ChildItem -Directory | ForEach-Object {
        $dirName = $_.Name
        $year = GetYear($dirName)

        if (-not(($map.Value).ContainsKey($year))) {
            ($map.Value).Add($year, @()) # (New-Object System.Collections.ArrayList))
        }

        if (containsActa($dirName)) {
        
            $yearInt = [int]$year   
            if ($yearInt -le 2018) {

                Push-Location $dirName
        
                Get-ChildItem -File | Sort-Object | ForEach-Object {
                    if (Test-Minute-File($_.Name)) {

                        if (($_.Name) -match "[0-9]+") {
                        
                            $date = Get-Date-From-Minute-File2 $wordApp $year $_
                            if (([String] $date).Length -gt 9) {
                                $date = ([String] $date).Substring(0,10)
                            }
                            ($map.Value[$year]) += @(newSeqInfo ([int]$Matches[0])  -1 $date $_.Name $_.FullName False)
                        }
                    }
                }
                Pop-Location
            }
        }
	    else {
	        Write-Host "No acta $dirName"
	    }
    }
    
    $wordApp.Quit()

    # return $map
}

# analysisLevel2
#
function analysingLevel2([String]$dirBase, [ref]$map) {
    
    $wordApp = New-Object -ComObject Word.Application
    # $minuteList = New-Object System.Collections.ArrayList

    Set-Location $dirBase
    
    Get-ChildItem -Directory | ForEach-Object {
        $dirName = $_.Name
        $year = GetYear($dirName)
        
        if (-not(($map.Value).ContainsKey($year))) {
            ($map.Value).Add($year, @())
        }

        if (containsActa($dirName)) {
            
            $yearInt = [int]$year

            if ($yearInt -le 2018) {

                Push-Location $dirName
            
                Get-ChildItem -Directory | ForEach-Object {

                    $dirName2 = $_.Name
                    if (containsActa($dirName2)) {
                        Push-Location $dirName2

                        Get-ChildItem -File | ForEach-Object {

                            if (Test-Minute-File($_.Name)) {

                                if (($_.Name) -match "[0-9]+") {
                                    # Write-Host $_.Name
                                    $tmpMatches=$Matches[0]
                                    if (($_.Name) -match "\(([0-9]+)\)") {
                                    
                                        $date = Get-Date-From-Minute-File2 $wordApp $year $_
                                        if (([String] $date).Length -gt 9) {
                                            $date = ([String] $date).Substring(0,10)
                                        }
                                        ($map.Value[$year]) += @(newSeqInfo ([int]$Matches[1]) -1 $date $_.Name $_.FullName True)
                                    }
                                    else {
                                        $date = Get-Date-From-Minute-File2 $wordApp $year $_
                                        if (([String] $date).Length -gt 9) {
                                            $date = ([String] $date).Substring(0,10)
                                        }
                                        ($map.Value[$year]) += @(newSeqInfo ([int]$tmpMatches) -1 "2019-10-11" $_.Name $_.FullName True)
                                    }
                                }
                                # Write-Host("File Level 1: " + $_.Name + " Year: " + $year)
                            }
			
                        }   

                        Pop-Location
                    }
                }

                Pop-Location
            }
        }
    }

    $wordApp.Quit()
    # return $map
}

# processMap
#
# purpose: $map relates a Year with several sequences a sequences inside a Year
function processMap([ref]$map) {
    foreach($key in $map.value.Keys) {
        $i = 1
        foreach ($elem in ($map.Value[$key] | Sort-Object -Property Seq)) {
            $elem.SeqYear = $i
            # Write-Host ("Year: " + $key + " seq: " + $elem.Seq + " id: " + $elem.SeqYear)
            printSeqInfo($elem)
            $i = $i + 1
        }
    }

}

# CreateDirIsNotExists
#
# purpose: if $dir doesn't exists, it creates
function CreateDirIsNotExists($dir) {
    if (-not(Test-Path $dir)) {
        New-Item -Name $dir -ItemType "directory" 
    }
}

# createDirectoryHierarchy
#
# purpose: creates a directory 
function createDirectoryHierarchy($destDir, $outputDir, [ref]$infoDirMap) {
    CreateDirIsNotExists($destDir)
    Push-Location $destDir
    CreateDirIsNotExists($outputDir)
    Push-Location $outputDir
    foreach ($key in $infoDirMap.Value.Keys) {
        CreateDirIsNotExists($key)
        Push-Location $key
        foreach ($elem in ($infoDirMap.Value[$key])) {
            $actaName = "acta-" + (($elem.SeqYear).toString("0#")) + "(" + (($elem.Seq).toString("00#")) + ")"
            CreateDirIsNotExists($actaName)  
        }
        Pop-Location
        # Write-Host "New-Item -Name Acta-",$key,'-ItemType "directory"'
        # Write-Host "New-Item -Name Acta-",$key,"\",({0:d3} -f $infoDirMap.Value[$key].Seq),"-",({0:d2} -f $infoDirMap.Value[$key].SeqYear),'-ItemType "directory"'
    }
    
    Pop-Location
    Pop-Location
}

function main() {
    $wordApp = New-Object -ComObject Word.Application
    Set-Location ~
    Set-Location "Documents\Beatriz\ACTAS CONSEJO ESCUELA"
    # $minuteNumber = 0
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

    # for ($i = 0; $i -eq 130; $i++) {
	#   if (-Not ($minuteListArray.Contains($i))) {
	    #       Write-Host ('Minute ' + $i + "doesn't exist")
	    #   }
	#}
    $wordApp.Quit()
    # Clear-Host
}

# Recompiling information
$mapDirs = @{}
$srcDir = "Documents\Beatriz\ACTAS CONSEJO ESCUELA"
Set-Location ~
Write-Host "Analysis 1"
analysingLevel1 $srcDir ([ref]$mapDirs)
Write-Host "Not yet Analysis 2"
Set-Location ~
analysingLevel2 $srcDir ([ref]$mapDirs)
Set-Location ~

processMap ([ref]$mapDirs)
createDirectoryHierarchy $env:TEMP "ACTAS CONSEJO DE ESCUELA" ([ref]$mapDirs)
Set-Location ~
