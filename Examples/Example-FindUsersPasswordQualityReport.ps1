Import-Module .\PasswordSolution.psd1 -Force

#Show-PasswordQuality -FilePath $PSScriptRoot\Reporting\PasswordQuality.html -Online -WeakPasswords "Test1", "Test2", "Test3" -Verbose
Show-PasswordQuality -FilePath "C:\Support\GitHub\TheDashboard\Ignore\Reports\CustomReports\PasswordQuality_$(Get-Date -f yyyy-MM-dd_HHmmss).html" -WeakPasswords "Test1", "Test2", "Test3" -SeparateDuplicateGroups -PassThru