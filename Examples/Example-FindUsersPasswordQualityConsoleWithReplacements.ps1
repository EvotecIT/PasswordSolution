Import-Module .\PasswordSolution.psd1 -Force

$Replacements = @(
    New-PasswordConfigurationReplacement -PropertyName 'Country' -Type eq -PropertyReplacementHash @{
        'PL'      = 'Poland'
        'DE'      = 'Germany'
        'AT'      = 'Austria'
        'IT'      = 'Italy'
        'Unknown' = 'Not specified in AD'
    } -OverwritePropertyName 'CountryCode'
)

$Users = Find-PasswordQuality -Replacements $Replacements
$Users | Format-Table