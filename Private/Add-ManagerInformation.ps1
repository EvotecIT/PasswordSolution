function Add-ManagerInformation {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $SummaryDictionary,
        [string] $Type,
        [string] $ManagerType,
        [Object] $Key,
        [PSCustomObject] $User,
        [PSCustomObject] $Rule
        # [bool] $Enabled
    )
    #if ($Enabled) {
    if ($Key) {
        if ($Key -is [string]) {
            $KeyDN = $Key
        } else {
            $KeyDN = $Key.EmailAddress
        }

        if (-not $SummaryDictionary[$KeyDN]) {
            $SummaryDictionary[$KeyDN] = [ordered] @{
                Manager             = $Key
                ManagerDefault      = [ordered] @{}
                ManagerNotCompliant = [ordered] @{}
            }
        }
        $SummaryDictionary[$KeyDN][$Type][$User.DistinguishedName] = [ordered] @{
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
                'Manager Status' = $ManagerType
                #ManagerStatus = $User.ManagerStatus
                Manager          = $User.Manager
                'Manager Email'  = $User.ManagerEmail
                #RuleName       = $Rule.Name
                #Type           = $Type
            }
            $SummaryDictionary[$KeyDN][$Type][$User.DistinguishedName]['Output'] = [PSCustomObject] ( $Extended + $Default)
        } else {
            $SummaryDictionary[$KeyDN][$Type][$User.DistinguishedName]['Output'] = [PSCustomObject] $Default
        }
    }
    # }
}