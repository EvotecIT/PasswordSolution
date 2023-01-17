Import-Module .\PasswordSolution.psd1 -Force

$Users = Find-Password -OverwriteEmailProperty 'extensionAttribute7' | Where-Object { $_.Name -eq 'Test Contact'}
$Users | Format-Table Name, Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus, ManagerLastLogonDays, ManagerType, Domain, UserPrincipalName