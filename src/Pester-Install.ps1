$mm = ($ENV:PSModulePath -split ';')[0]
$pesterPath = "$mm\Pester"
$pesterRoot = "..\Pester\*" 

if( -not (Test-Path $pesterPath)){
    mkdir $pesterPath    
}

copy $pesterRoot $pesterPath -recurse -force

Import-Module -Verbose Pester


