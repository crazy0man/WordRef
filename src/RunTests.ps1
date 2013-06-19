Import-Module Pester

Invoke-Pester -EnableLegacyExpectations  -OutputXml "TestResults.xml"

Remove-Module Pester
