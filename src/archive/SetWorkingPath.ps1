# The default workingPath should be updated as required
param($workingPath)
$defaultWorkingPathForTFS = "D:\Decoupling"
if($workingPath -eq "default"){
    $workingPath = $defaultWorkingPathForTFS
} elseif ($workingPath -eq ".") {
    $workingPath = (Get-Location).Path
} elseif($workingPath -eq $null){
    throw "$workingPath is required."
    exit
}

if(($workingPath -eq $null)){
    throw "the workingPath parameters are required."
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

    foreach($field in $doc.Fields){
        $_ = $field
        "*" * 50
        $_
        "*" * 25
        #$_.Next
        #$_.Next.LinkFormat
        $linkFormat = $_.LinkFormat        
        $linkFormat
        if( ($linkFormat.SourcePath -ne $null) -and 
            ($linkFormat.SourcePath -ne $workingPath)){
            # $linkFormat.AutoUpdate = $true
            # Issue: the update will replace the relative reference to wrong format
            # from: {filename /p}\\..\\Cross-Reference-NewSrc.docx
            # to: Document1\\..\\Cross-Reference-NewSrc.docx
            $newSourceFullName =  $workingPath + "\" + $linkFormat.SourceName
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