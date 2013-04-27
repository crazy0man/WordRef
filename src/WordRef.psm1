$script:Debug = $false
$script:fieldType = 68 # TODO: need to find out what the type's specific name and meaning
# WordRef-AutoUpdate, update the ref content to latest
function WordRef-AutoUpdate {
    param(
        [parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string] $docFullName,
        [switch] $inShapes
        )
    # Setup open word application
    begin { 
        # close all running word processes is a must
        # StopWordProcess -confirm
        StopWordProcess
        $word = SetupWordApp
        
    }
    process { 
        try{
            $doc = DocOpen $word $docFullName
            "Updating file:" + $docFullName
            $doc.Fields| foreach{                 
                $_.Update() | Out-NULL
            }
            if($inShapes){
                $doc.Shapes | foreach{
                    if($_.TextFrame.HasText){
                        $_.TextFrame.TextRange.Fields.Update() | Out-NULL
                    }
                }
            }
            # Update each field
            DocSave($doc)

        } catch{
            throw $_
        } finally{
            DocClose($doc)    
        }
    }
    end { 
        # Exit Word
        DestroyWordApp $word
    }
    
    # End close Word
}
# WordRef-Search, search out specific reference matched the pattern in the current doc
function WordRef-Search {
    return 1
}

# WordRef-List, list out all reference in the current doc
function WordRef-List {
    param(
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string] $docFullName,
        [switch] $inShapes,
        [switch] $showAsFullPath,
        [switch] $showAsPath        
        ) 
    begin{
        StopWordProcess
        $wordApp = SetupWordApp
    }

    process{
        try{
            $doc = DocOpen $wordApp $docFullName
            GetFields $doc $script:fieldType | foreach {
                    GetFieldSourceName $_ $showAsFullPath $showAsPath
                }

            if($inShapes){
                $fieldList = GetFieldsInShapes $doc
                $fieldList | foreach {
                   GetFieldSourceName $_ $showAsFullPath $showAsPath
                }
            }

        } catch{
            throw $_
        } finally{
            DocClose $doc    
        }    
    }

    end{
        DestroyWordApp $wordApp
    }
    
}

# WordRef-Remove, remove the reference which are matched with the pattern
function WordRef-Remove {
    return 1
}

# WordRef-Rename, rename the reference with the pattern
function WordRef-Rename {
    param(
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string] $docFullName,
        [parameter(Mandatory=$true)]
        [regex] $pattern,
        [parameter(Mandatory=$true)]
        [string] $newValue,
        [switch] $inShapes
        ) 
    begin{
        StopWordProcess
        $wordApp = SetupWordApp
    }

    process{
        try{
            $doc = DocOpen $wordApp $docFullName

            (GetFields $doc $script:fieldType) | foreach{
                RenameFieldSourceName $_ $pattern $newValue
            }

            if($inShapes){
                (GetFieldsInShapes $doc ) | foreach{
                    RenameFieldSourceName $_ $pattern $newValue
                }
            }
            DocSave $doc

        } catch{
            throw $_
        } finally{
            DocClose $doc    
        }
    }

    end{
        DestroyWordApp $wordApp
    }    
}

# WordRef-SetPath, set the ref path to the new value when match the path pattern
function WordRef-SetPath {
    param(
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string] $docFullName,
        [parameter(Mandatory=$true)]
        [string] $oldPathPattern,
        [parameter(Mandatory=$true)]
        [string] $newPath,
        [switch] $inShapes
    )

    begin{
        StopWordProcess
        $wordApp = SetupWordApp
    }

    process{
        try{
            $doc = DocOpen $wordApp $docFullName

            GetFields $doc $script:fieldType | foreach{
                SetFieldPath $_ $oldPathPattern $newPath
            }

            if($inShapes){
                GetFieldsInShapes $doc | foreach{
                    SetFieldPath $_ $oldPathPattern $newPath
                }
            }

            DocSave $doc
        
        }
        catch{
            throw $_
        }
        finally{
            DocClose $doc
        }    
    }

    end{
        DestroyWordApp $wordApp
    }
}

# TODO: the high priority task to implement
function WordRef-Check {
    
}

function DocOpen ($wordApp, $docFullPath) {
    try {
        if($wordApp -eq $null){
            throw "wordApp is null"
        }

        if($docFullPath -eq $null -or
            (-not (Test-Path $docFullPath))) {
            throw "docFullPath is null or not existing"
        }
        
        return $wordApp.Documents.Open($docFullPath)    
    } catch [Exception]{
        $_
        # Caused by the script to open temporay docs
        # @{should=System.Object}Exception calling "Open" with "1" argument(s): "The file appears to be corrupted."
        # At D:\WorkStation\TFS2012\Grenoble\Decoupling\WordRef\WordRef.psm1:208 char:35
        # +     return $wordApp.Documents.Open <<<< ($docFullPath)
        # + CategoryInfo          : NotSpecified: (:) [], MethodInvocationException
        # + FullyQualifiedErrorId : ComMethodTargetInvocation
    }
    
}

function GetFields($doc, $fieldType) {
    if($doc -eq $null){
        throw "The doc is null"
    }
    $doc.Fields | foreach {
            if($_.Type -eq $fieldType){
                $_
            }
        }
}

function GetFieldsInShapes($doc){
    if($doc -eq $null){
        throw "The doc is null"
    }
    $doc.Shapes | 
        Where {$_.TextFrame.HasText} | 
            foreach {
                $_.TextFrame.TextRange.Fields
            }
}

function GetFieldSourceName($field, $showAsFullPath, $showAsPath){
    if($showAsFullPath) {
        $field.LinkFormat.SourceFullName
    } elseif($showAsPath){
        $field.LinkFormat.SourcePath
    } else {
        $field.LinkFormat.SourceName
    }
}

function RenameFieldSourceName($field, $pattern,$newValue){
    # $field.LinkFormat.SourceName
    if($field.LinkFormat.SourceName -match $pattern){
        $newName = $field.LinkFormat.SourceName -replace $pattern, $newValue
        # $field.LinkFormat.SourceFullName
        # $field.LinkFormat.SourcePath
        $newFullName = $field.LinkFormat.SourcePath + "\$newName"
        # $newFullName
        $field.LinkFormat.SourceFullName = $newFullName
        $field.LinkFormat.Update()
    }    
}

function SetFieldPath($field, $oldPathPattern,$newPath){
    $field.LinkFormat.SourceFullName 
    if($field.LinkFormat.SourcePath -eq $oldPathPattern){
        if(-not $newPath.EndsWith('\')){
            $newPath += '\'
        }
        
        $field.LinkFormat.SourceFullName = $newPath + $field.LinkFormat.SourceName
        $field.LinkFormat.Update()
    }
}

function DocSave ($doc) {
    if($doc -ne $null){
        if($doc.Saved -eq $false){
            $doc.Save() | Out-NULL
        }            
    }
}

function DocClose ($doc) {
    try {
        if($doc -ne $null){
            $doc.Close() | Out-NULL
        }    
    }    
    catch [Exception]{
        $_
        # Caused by the script to open temporay docs
        #         @{should=System.Object}Exception calling "Close" with "3" argument(s): "The object invoked has disconnected from its cl
        # ients. (Exception from HRESULT: 0x80010108 (RPC_E_DISCONNECTED))"
        # At D:\WorkStation\TFS2012\Grenoble\Decoupling\WordRef\WordRef.psm1:279 char:19
        # +         $doc.Close <<<< ()
        #     + CategoryInfo          : NotSpecified: (:) [], MethodInvocationException
        #     + FullyQualifiedErrorId : DotNetMethodException
    }
}

function DocParagraphs($doc) {
    if($doc -ne $null){
        return $doc.Paragraphs
    }
}

function StopWordProcess {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param($pName="WINWORD")
    foreach($p in Get-WmiObject Win32_Process | where {$_.Name -match $pName}){
        if($PSCmdlet.ShouldProcess(
            "process $($p.Name) " +
            " (id: $($p.ProcessId))",
            "Stop Process")){
            # $p | Stop-Process
            if($p -ne $null){
                # Always throws the following exception:
                # TODO: fix the exception
                # Exception calling "Terminate" : "Not found "
                # At D:\WorkStation\TFS2012\Grenoble\Decoupling\WordRef\wordref.psm1:105 char:29
                # +                 $p.Terminate <<<< ()
                # + CategoryInfo          : NotSpecified: (:) [], MethodInvocationException
                # + FullyQualifiedErrorId : WMIMethodException

                try {
                    $p.Terminate()
                }
                catch {
                    
                }
            }                
        }
    }
}

function SetupWordApp {
    $wordApp = new-object -ComObject "Word.Application"
    # $wordApp.Visible = $script:Debug
    $wordApp.Visible = false
    return $wordApp
}

function DestroyWordApp ($wordApp) {
    if($wordApp -ne $null){
        $wordApp.Quit() | Out-NULL
    }
}

function UpdateDocContent(
    $docFullName, 
    $rangeMark,
    $newValue){
    try{
        $wordApp = SetupWordApp
        $doc = DocOpen $wordApp $docFullName    
        (DocParagraphs $doc) | foreach{
            if ($_.Range.Text -match $WordRefMark) {
                # Here can't use the $WordRefMark to construct the new value
                #TODO: fix it
                $_.Range.Text = "WordRef:" + $newValue
            }
        }
        # $doc.Paragraphs | foreach{       
        #     $_.Range.Text
        #     if($_.Range.Text -match "WordRef"){
        #         $_.Range.Text = "WordRef:" + (Get-Date)
        #     }
        #     # $_ | gm
        # }
        DocSave($doc)
    } catch{
        throw $_
    } finally{
        DocClose($doc)
        DestroyWordApp $wordApp
    }
    

}

function GetDocContent {
    param(
        [parameter(Mandatory=$true)]
        [string]$docFullName,
        [parameter(Mandatory=$true)]
        [string]$rangeMark,
        [switch]$inShapes
        )
    if($docFullName -eq "" -or 
        $rangeMark -eq ""){
        throw "Invalid docFullName or rangeMark"
    }
    $wordApp = SetupWordApp
    #$docFullName
    $doc = DocOpen $wordApp $docFullName
    #[void]$result
    (DocParagraphs $doc) | foreach{
         if($_.Range.Text -match $rangeMark){            
                #$_.Range
                $_.Range.Text
        }        
    }
    if($inShapes){
        $doc.Shapes | foreach{
            $_.TextFrame.TextRange.Text
            # if($_.TextFrame.HasText){
            #     $_.TextFrame.TextRange.Text
            # }
        }
    }
    DocClose $doc
    DestroyWordApp $wordApp   
    # $result 
}

function removeSpecialCharacter($text) {
        return $text -replace '\W*',''
}