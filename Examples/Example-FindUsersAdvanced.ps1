Import-Module .\PasswordSolution.psd1 -Force

$RuleProperties = @(
    'ExtensionAttribute7'
    'extensionAttribute7'
    'extensionAttribute8'
)

$Users = Find-Password -OverwriteEmailProperty 'extensionAttribute7' -RulesProperties $RuleProperties #| Where-Object { $_.Name -eq 'Test Contact' }
#$Users | Format-Table Name, Extension*, DateExpiry  #Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus, ManagerLastLogonDays, ManagerType, Domain, UserPrincipalName

$Users | Select-Object -First 5 | Format-Table Name, Extension*, DateExpiry, LastLogonDate, LastLogonDays