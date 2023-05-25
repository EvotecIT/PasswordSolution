function New-PasswordConfigurationType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('User', 'Manager', 'Security', 'Admin')][string] $Type,
        [switch] $Enable,
        [int] $SendCountMaximum,
        [string] $DefaultEmail,
        [switch] $AttachCSV
    )

    $Output = [ordered] @{
        Type     = "PasswordConfigurationType$Type"
        Settings = @{
            Enable                 = $Enable.IsPresent
            SendCountMaximum       = $SendCountMaximum
            SendToDefaultEmail     = if ($DefaultEmail) { $true } else { $false }
            DefaultEmail           = $DefaultEmail
            OverwriteEmailProperty = $OverwriteEmailProperty
            AttachCSV              = $AttachCSV.IsPresent
        }
    }
    $Output
}