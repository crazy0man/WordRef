$mm = ($ENV:PSModulePath -split ';')[0]
$pesterPath = "$mm\Pester"
#update the pesterRoot to correct value
$pesterRoot = "..\..\Pester\*" 

if( -not (Test-Path $pesterPath)){
    mkdir $pesterPath    
}

copy $pesterRoot $pesterPath -recurse -force

Import-Module -Verbose Pester


