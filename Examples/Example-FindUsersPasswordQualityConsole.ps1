Import-Module .\PasswordSolution.psd1 -Force

$Users = Find-PasswordQuality
$Users | Format-Table

#$Users = Find-PasswordQuality -IncludeDomains 'ad.evotec.pl'
#$Users | Format-Table