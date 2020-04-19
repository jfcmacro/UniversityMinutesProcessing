function newSeqInfo($seq,$seqYear,$date) {
    
    return [pscustomobject]@{Seq=$seq;SeqYear=$seqYear;Date=$date}
}

function printSeqInfo($obj) {
    Write-Host "Seq: ",$obj.Seq," SeqYear: ",$obj.SeqYear," Date: ",$obj.Date
}

$arreglo = @()
$arreglo += @(newSeqInfo 4 -1 "2019-11-12")
$arreglo += @(newSeqInfo 3 -1 "2019-10-09")
$arreglo += @(newSeqInfo 2 -1 "2019-10-11")
$arreglo += @(newSeqInfo 1 -1 "2019-10-10")

$arreglo = $arreglo | Sort-Object -Property Seq 
$i=$arreglo.Count
$arreglo | ForEach-Object {
    $_.SeqYear = $i
    $i = $i + 1
    printSeqInfo($_)
}
