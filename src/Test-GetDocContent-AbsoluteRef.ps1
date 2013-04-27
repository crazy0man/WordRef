import-module .\WordRef.psm1

StopWordProcess
$docFullPath = (dir "WordRef-AbsoluteRef.docx").FullName
$result = GetDocContent $docFullPath "WordRef"

$result = removeSpecialCharacter $result

"test"

$result
remove-module wordref