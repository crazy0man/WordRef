import-module .\WordRef.psm1

StopWordProcess
$docFullPath = (dir "WordRef-Src.docx").FullName
UpdateDocContent $docFullPath "WordRef" (Get-Date)    

remove-module wordref