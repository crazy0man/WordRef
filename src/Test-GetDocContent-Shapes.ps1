import-module .\WordRef.psm1

StopWordProcess
$docFullPath = (dir "WordRef-RefInTextBox.docx").FullName
$result = GetDocContent -inShapes $docFullPath " "

$result = removeSpecialCharacter $result
$result

#$actualValue 
"test"
#$actualValue -match "WordRef"
# [string]($result) | Should be "1"
remove-module wordref