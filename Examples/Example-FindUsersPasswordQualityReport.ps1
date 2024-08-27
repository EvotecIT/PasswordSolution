Clear-Host
Import-Module .\PasswordSolution.psd1 -Force

# option 1, one-liner
# Show-PasswordQuality -FilePath C:\Temp\PasswordQuality.html -Online -WeakPasswords "Test1", "Test2", "Test3" -Verbose -SeparateDuplicateGroups -AddWorldMap -PassThru

# option 2, for easier reading with splatting
$showPasswordQualitySplat = @{
    FilePath                      = "$PSScriptRoot\Reporting\PasswordQuality_$(Get-Date -f yyyy-MM-dd_HHmmss).html"
    WeakPasswords                 = "Test1", "Test2", "Test3", 'February2023!#!@ok', $Passwords | ForEach-Object { $_ }
    WeakPasswordsHashesFile       = 'C:\Support\GitHub\PwnedDatabaseDownloader\pwnedpasswords_ntlm.txt'
    WeakPasswordsHashesSortedFile = 'C:\Support\GitHub\PwnedDatabaseDownloader\pwnedpasswords_ntlm.txt'
    SeparateDuplicateGroups       = $true
    PassThru                      = $true
    AddWorldMap                   = $true
    LogPath                       = "$PSScriptRoot\Logs\PasswordQuality_$(Get-Date -f yyyy-MM-dd_HHmmss).log"
    Online                        = $true
    LogMaximum                    = 5
}

Show-PasswordQuality @showPasswordQualitySplat -Verbose