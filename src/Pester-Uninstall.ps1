$mm = ($ENV:PSModulePath -split ';')[0]
$pesterPath = "$mm\Pester"

if(Test-Path $pesterPath){
    dir $pesterPath | remove-item
}