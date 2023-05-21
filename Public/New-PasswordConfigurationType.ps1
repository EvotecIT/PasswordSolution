function New-PasswordConfigurationType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('User', 'Manager', 'Security', 'Admin')][string] $Type,
        [switch] $Enable,
        [int] $SendCountMaximum,
        #$SendToDefaultEmail,
        [string] $DefaultEmail,
        [string] $OverwriteEmailProperty,
        [switch] $AttachCSV
    )

    $Output = [ordered] @{
        Type     = "PasswordConfigurationType$Type"
        Settings = @{
            Enable                 = $Enable.IsPresent
            SendCountMaximum       = $SendCountMaximum
            #SendToDefaultEmail = $SendToDefaultEmail.IsPresent
            DefaultEmail           = $DefaultEmail
            OverwriteEmailProperty = $OverwriteEmailProperty
            AttachCSV              = $AttachCSV.IsPresent
        }
    }
    $Output
}