import-module .\WordRef.psm1

StopWordProcess
$docFullPath = (dir "WordRef-Src.docx").FullName
$result = GetDocContent $docFullPath "WordRef"
# $result[$result.Count-1]
# $result
$actualValue 
foreach($str in $result){
    ($str).GetType().FullName
    $str| gm
    if(($str).Trim() -ne "")
    {
        $actualValue += $str.Trim()
    }
}
$actualValue 
"test"
$actualValue -match "WordRef"
# [string]($result) | Should be "1"
remove-module wordref