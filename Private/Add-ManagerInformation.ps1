function Add-ManagerInformation {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $SummaryDictionary,
        [string] $Type,
        [string] $Key,
        [PSCustomObject] $User,
        [PSCustomObject] $Rule,
        [bool] $Enabled
    )
    if ($Enabled) {
        if ($Key) {
            if (-not $SummaryDictionary[$Key]) {
                $SummaryDictionary[$Key] = [ordered] @{
                    ManagerDefault      = [ordered] @{}
                    ManagerNotCompliant = [ordered] @{}
                    ManagerMissingEmail = [ordered] @{}
                    ManagerDisabled     = [ordered] @{}
                    ManagerMissing      = [ordered] @{}
                }
            }
            $SummaryDictionary[$Key][$Type][$User.DistinguishedName] = [ordered] @{
                Manager       = $User.ManagerDN
                User          = $User
                Rule          = $Rule
                ManagerOption = $Type
                Output        = [ordered] @{}
            }
            $Default = [ordered] @{
                DisplayName     = $User.DisplayName
                Enabled         = $User.Enabled
                SamAccountName  = $User.SamAccountName
                Domain          = $User.Domain
                DateExpiry      = $User.DateExpiry
                DaysToExpire    = $User.DaysToExpire
                PasswordLastSet = $User.PasswordLastSet
                PasswordExpired = $User.PasswordExpired
            }
            if ($Type -ne 'ManagerDefault') {
                $Extended = [ordered] @{
                    ManagerStatus = $User.ManagerStatus
                    Manager       = $User.Manager
                    ManagerEmail  = $User.ManagerEmail
                    #RuleName       = $Rule.Name
                    #Type           = $Type
                }
                $SummaryDictionary[$Key][$Type][$User.DistinguishedName]['Output'] = [PSCustomObject] ( $Extended + $Default)
            } else {
                $SummaryDictionary[$Key][$Type][$User.DistinguishedName]['Output'] = [PSCustomObject] $Default
            }
        }
    }
}