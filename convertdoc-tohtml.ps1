param([string]$docpath,[string]$htmlpath = $docpath)

$srcfiles = Get-ChildItem $docPath -filter "*.doc"
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatFilteredHTML");
$word = new-object -comobject word.application
$word.Visible = $False

function saveas-filteredhtml
{
    $opendoc = $word.documents.open($doc.FullName);
    $opendoc.saveas([ref]"$htmlpath\$doc.fullname.html", [ref]$saveFormat);
    $opendoc.close();
}

ForEach ($doc in $srcfiles)
{
    Write-Host "Processing :" $doc.FullName
    saveas-filteredhtml
    $doc = $null
}

$word.quit();