Import-Module .\PasswordSolution.psd1 -Force

$Users = Find-Password
$Users | Format-Table Name, Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus, ManagerLastLogonDays, ManagerType, Domain, UserPrincipalName