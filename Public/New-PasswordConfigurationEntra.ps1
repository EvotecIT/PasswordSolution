function New-PasswordConfigurationEntra {
    [CmdletBinding()]
    param(
        [switch] $Enable
    )
    $Output = [ordered] @{
        Type     = "PasswordConfigurationEntra"
        Settings = [ordered] @{
            Enabled = $Enable.IsPresent
        }
    }
    $Output
}