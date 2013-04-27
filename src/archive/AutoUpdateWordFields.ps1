# $doc.Fields | %{$_.Update()}
$word=new-object -ComObject "Word.Application"
$word.Visible = $false
# $word.Documents | get-member | sort name
# dir *.docm | get-member | sort name
$fileList = ( dir -path *.docx | select name) 
# $fileList
$location = Get-Location
foreach( $fileName in $fileList){
    "Updating " + $fileName.Name + "..."
    $fullName = $location.Path + '\' + $fileName.Name
    $fullName
    $doc = $word.Documents.Open($fullName)
    # $doc | gm | sort name
    $doc.Fields | foreach{
        # $_ | gm
        "_" * 10
        # $linkFormat = $_.LinkFormat        
        # $linkFormat
        # if($linkFormat.SourceName -match 'docx'){
        #     $linkFormat.AutoUpdate = $true
        # }
        # $linkFormat.Update()
        "Updating field:"  + $_.Data
        $_.Update()
        # break
    }

    $doc.Save();
    $doc.Close();
    # break
}
$word.Quit()
"Complete"