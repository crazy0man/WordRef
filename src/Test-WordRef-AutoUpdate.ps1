import-module .\wordref.psm1
$docAbsoluteRef = "WordRef-AbsoluteRef.docx"
(.\Word-Range-test.ps1)
dir $docAbsoluteRef | foreach { $_.fullname } | WordRef-AutoUpdate

remove-module wordref
"Complete"