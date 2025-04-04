Import-Module .\PasswordSolution.psd1 -Force

$showPasswordQualitySplat = @{
    FilePath                = "$PSScriptRoot\Reporting\PasswordQuality_$(Get-Date -f yyyy-MM-dd_HHmmss).html"
    WeakPasswords           = "Test1", "Test2", "Test3", 'February2023!#!@ok', $Passwords | ForEach-Object { $_ }
    SeparateDuplicateGroups = $true
    PassThru                = $true
    AddWorldMap             = $true
    LogPath                 = "$PSScriptRoot\Logs\PasswordQuality_$(Get-Date -f yyyy-MM-dd_HHmmss).log"
    Online                  = $true
    LogMaximum              = 5
    Replacements            = New-PasswordConfigurationReplacement -PropertyName 'ExtensionAttribute4' -Type eq -PropertyReplacementHash @{
        'PL'      = 'Poland'
        'DE'      = 'Germany'
        'AT'      = 'Austria'
        'IT'      = 'Italy'
        'Unknown' = 'Not specified in AD'
    } #-OverwritePropertyName 'AddMe'
}

Show-PasswordQuality @showPasswordQualitySplat -Verbose