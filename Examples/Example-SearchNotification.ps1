Import-Module .\PasswordSolution.psd1 -Force

Find-PasswordNotification -SearchPath $PSScriptRoot\Search\SearchLog.xml | Format-Table
Find-PasswordNotification -SearchPath "$PSScriptRoot\Search\SearchLog_2021-06.xml" -Manager | Format-Table