import-module .\WordRef.psm1

    $setPathOrigData = "WordRef-SetPathTestOrgData.docx"
    $setPathTestData = "WordRef-SetPathTestData.docx"
    Copy $setPathOrigData $setPathTestData -force
    $setPathTestDataPath = (Resolve-Path ".\$setPathTestData")
    # $setPathTestDataPath
    # It "SetPath successful" {
        # Copy $setPathOrigData $setPathTestData -force
        # $setPathTestDataPath | 
        #     WordRef-SetPath -pattern "D:\WorkStation\WordRef" -newValue "D:\WorkStation\WordRef\NewRefPath" | Should be $true
    # }
    # It "SetPath including Shapes successful" {

        $setPathTestDataPath | 
            WordRef-SetPath -pattern 'D:\WorkStation\WordRef' -newPath "D:\WorkStation\WordRef\NewRefPath" 
    # }




remove-module wordref