# $doc.Fields | %{$_.Update()}
param($oldSourceName,
    $newSourceName)
if(($oldSourceName -eq $null) -or ($newSourceName -eq $null)){
    throw "the oldSourceName and newSourceName parameters are required."
    exit 
}
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
        "*" * 50
        $_
        "*" * 25
        $_.Next
        $_.Next.LinkFormat
        $linkFormat = $_.LinkFormat        
        $linkFormat
        if($linkFormat.SourceName -ceq $oldSourceName){
            # $linkFormat.AutoUpdate = $true
            # Issue: the update will replace the relative reference to wrong format
            # from: {filename /p}\\..\\Cross-Reference-NewSrc.docx
            # to: Document1\\..\\Cross-Reference-NewSrc.docx
            $newSourceFullName =  $linkFormat.SourcePath + "\$newSourceName"
            $linkFormat.SourceFullName = $newSourceFullName
            $linkFormat.Update()
            $_.LinkFormat
        }
    }

    $doc.Save();
    $doc.Close();
    # break
}
$word.Quit()
"Complete"