import-module .\wordref.psm1

dir "WordRef-src.docx" | foreach { $_.fullname } | WordRef-AutoUpdate

remove-module wordref