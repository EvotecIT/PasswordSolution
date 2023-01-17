Import-Module .\PasswordSolution.psd1 -Force

Show-PasswordQuality -FilePath $PSScriptRoot\Reporting\PasswordQuality.html -Online -WeakPasswords "Test1", "Test2", "Test3"