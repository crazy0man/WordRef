Import-Module ..\Pester\Pester.psm1

Invoke-Pester -EnableLegacyExpectations  -OutputXml "TestResults.xml"

Remove-Module Pester
