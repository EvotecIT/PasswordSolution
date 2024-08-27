Import-Module .\PasswordSolution.psd1 -Force

$Users = Find-Password -OverwriteEmailProperty 'extensionAttribute7' -OverwriteManagerProperty extensionAttribute1 -FilterOrganizationalUnit @(
    "*OU=Accounts,OU=Administration,DC=ad,DC=evotec,DC=xyz"
)
$Users | Format-Table UserPrincipalName, Name, Domain, Type, SamAccountName, OrganizationalUnit, Manager, ManagerEmail, ManagerStatus
#$Users | Sort-Object -Property Manager | Format-Table Name, Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus, ManagerLastLogonDays, ManagerType, Domain, UserPrincipalName