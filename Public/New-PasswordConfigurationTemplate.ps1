function New-PasswordConfigurationTemplate {
    [CmdletBinding()]
    param(
        [parameter(Mandatory)][ScriptBlock] $Template,
        [parameter(Mandatory)][string] $Subject,
        [parameter(Mandatory)][ValidateSet('PreExpiry', 'PostExpiry', 'Manager', 'ManagerNotCompliant', 'Security', 'Admin')] $Type
    )

    $Output = [ordered] @{
        Type     = "PasswordConfigurationTemplate$Type"
        Settings = [ordered] @{
            Template = $Template
            Subject  = $Subject
        }
    }
    $Output
}