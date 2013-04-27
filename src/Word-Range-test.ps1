# $doc.Fields | %{$_.Update()}
$word=new-object -ComObject "Word.Application"
$word.Visible = $false
# $word.Documents | get-member | sort name
# dir *.docm | get-member | sort name
$fileList = ( dir -path WordRef-src.docx | select name) 
# $fileList
$location = Get-Location
foreach( $fileName in $fileList){
    "Updating " + $fileName.Name + "..."
    $fullName = $location.Path + '\' + $fileName.Name
    $fullName
    $doc = $word.Documents.Open($fullName)
    # $doc | gm | sort name
    $doc.Paragraphs | foreach{
        # $_

        # "*" * 20
        $_.Range.Text
        if($_.Range.Text -match "WordRef"){
            $_.Range.Text = "WordRef:" + (Get-Date)
        }
        # $_ | gm
    }
    $doc.Save();
    $doc.Close();
    # break
}
$word.Quit()
"Complete"