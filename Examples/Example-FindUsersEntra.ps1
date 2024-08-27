Import-Module .\PasswordSolution.psd1 -Force


Connect-MgGraph -Scopes "User.Read.All" -NoWelcome

$Users = Find-PasswordEntra #-Verbose
#$Users | Format-Table
$Users | Out-HtmlView -ScrollX -Filtering -DataStore JavaScript {
    New-HTMLTableCondition -Name 'Enabled' -ComparisonType string -Value $True -BackgroundColor TeaGreen -FailBackgroundColor Salmon
    New-HTMLTableCondition -Name 'IsLicensed' -ComparisonType string -Value $True -BackgroundColor TeaGreen -FailBackgroundColor Salmon
    New-HTMLTableCondition -Name 'IsSynchronized' -ComparisonType string -Value $True -BackgroundColor TeaGreen -FailBackgroundColor Salmon
    New-HTMLTableCondition -Name 'PasswordExpired' -ComparisonType string -Value $false -BackgroundColor TeaGreen -FailBackgroundColor Salmon
} -DisablePaging