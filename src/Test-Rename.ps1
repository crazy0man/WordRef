import-module .\WordRef.psm1

$renameTestFile = "WordRef-RenameTestDataToTest.docx"
(Copy "WordRef-RenameTestOrgData.docx" $renameTestFile -force)
$renameTestFilePath = (Resolve-Path (".\$renameTestFile"))
# It "Rename ref name all to src3 successful" {
#     $renameTestFilePath | WordRef-Rename -pattern "WordRef-src\d.docx" -newValue "WordRef-src3.docx" | should be $true
    
# }
# It "Rename ref name all to src3 successful precisely" {
    $expectedValue = "WordRef-src3.docx"
    $renameTestFilePath | WordRef-Rename -pattern 'WordRef-src\d?.docx' -newValue $expectedValue 
    $refList = ($renameTestFilePath | WordRef-List -inShapes)
    # $refList | foreach {
    #         $_ | should be $expectedValue        
    #     }
    $refList
# }




remove-module wordref