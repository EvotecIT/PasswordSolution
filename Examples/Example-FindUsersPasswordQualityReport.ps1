Import-Module .\PasswordSolution.psd1 -Force

#Show-PasswordQuality -FilePath $PSScriptRoot\Reporting\PasswordQuality.html -Online -WeakPasswords "Test1", "Test2", "Test3" -Verbose
$showPasswordQualitySplat = @{
    #FilePath                = "C:\Support\GitHub\TheDashboard\Ignore\Reports\CustomReports\PasswordQuality_$(Get-Date -f yyyy-MM-dd_HHmmss).html"
    FilePath                = "$PSScriptRoot\Reporting\PasswordQuality_$(Get-Date -f yyyy-MM-dd_HHmmss).html"
    WeakPasswords           = "Test1", "Test2", "Test3", 'February2023!#!@ok', $Passwords | ForEach-Object { $_ }
    SeparateDuplicateGroups = $true
    PassThru                = $true
    AddWorldMap             = $true
    LogPath                 = "$PSScriptRoot\Logs\PasswordQuality_$(Get-Date -f yyyy-MM-dd_HHmmss).log"
    Online                  = $true
}

Show-PasswordQuality @showPasswordQualitySplat