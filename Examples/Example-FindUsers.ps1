Import-Module .\PasswordSolution.psd1 -Force

$Users = Find-Password -OverwriteEmailProperty 'extensionAttribute7' -OverwriteManagerProperty extensionAttribute1
$Users | Sort-Object -Property Manager | Format-Table Name, Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus, ManagerLastLogonDays, ManagerType, Domain, UserPrincipalName