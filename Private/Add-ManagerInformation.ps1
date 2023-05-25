function Add-ManagerInformation {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $SummaryDictionary,
        [string] $Type,
        [string] $ManagerType,
        [Object] $Key,
        [PSCustomObject] $User,
        [PSCustomObject] $Rule
    )
    if ($Key) {
        if ($Key -is [string]) {
            $KeyDN = $Key
        } else {
            $KeyDN = $Key.DisplayName
        }

        if (-not $SummaryDictionary[$KeyDN]) {
            $SummaryDictionary[$KeyDN] = [ordered] @{
                Manager             = $Key
                ManagerDefault      = [ordered] @{}
                ManagerNotCompliant = [ordered] @{}
                Security            = [ordered] @{}
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
                'Status'        = $ManagerType
                'Manager'       = $User.Manager
                'Manager Email' = $User.ManagerEmail
            }
            $SummaryDictionary[$KeyDN][$Type][$User.DistinguishedName]['Output'] = [PSCustomObject] ( $Extended + $Default)
        } else {
            $SummaryDictionary[$KeyDN][$Type][$User.DistinguishedName]['Output'] = [PSCustomObject] $Default
        }
    }
}