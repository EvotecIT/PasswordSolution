function New-PasswordQualityTemplate {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory)][scriptblock] $Body,
        [string] $Subject = 'Password Quality Notification',
        [ValidateSet('OR', 'AND')][string] $Operator = 'OR',
        [string[]] $Domains,
        [switch] $HasMailbox,
        [switch] $Enabled,
        [switch] $ClearTextPassword,  #: False
        [switch] $LMHash,  #: False
        [switch] $EmptyPassword,  #: True
        [switch] $WeakPassword,  #: False
        [switch] $AESKeysMissing,  #: False
        [switch] $PreAuthNotRequired,  #: False
        [switch] $DESEncryptionOnly,  #: False
        [switch] $Kerberoastable,  #: False
        [switch] $DelegatableAdmins,  #: False
        [switch] $SmartCardUsersWithPassword,  #: False
        [switch] $DuplicatePasswordGroups,  #: False
        [int] $DuplicatePasswordGroupsCount,
        [ValidateSet('eq', 'lt', 'gt')][string] $DuplicatePasswordGroupsType = 'eq',
        [string[]] $OrganizationalUnit,
        [string[]] $MemberOf,
        [string] $Country
    )

    $Data = [ordered] @{
        Type     = 'QualityEmail'
        Settings = [ordered] @{
            Body                         = $Body
            Subject                      = $Subject
            Operator                     = $Operator
            Domains                      = if ($PSBoundParameters.ContainsKey('Domains')) { $Domains } else { $null }
            HasMailbox                   = if ($PSBoundParameters.ContainsKey('HasMailbox')) { $HasMailbox.IsPresent } else { $null }
            Enabled                      = if ($PSBoundParameters.ContainsKey('Enabled')) { $Enabled.IsPresent } else { $null }
            ClearTextPassword            = if ($PSBoundParameters.ContainsKey('ClearTextPassword')) { $ClearTextPassword.IsPresent } else { $null }
            LMHash                       = if ($PSBoundParameters.ContainsKey('LMHash')) { $LMHash.IsPresent } else { $null }
            EmptyPassword                = if ($PSBoundParameters.ContainsKey('EmptyPassword')) { $EmptyPassword.IsPresent } else { $null }
            WeakPassword                 = if ($PSBoundParameters.ContainsKey('WeakPassword')) { $WeakPassword.IsPresent } else { $null }
            AESKeysMissing               = if ($PSBoundParameters.ContainsKey('AESKeysMissing')) { $AESKeysMissing.IsPresent } else { $null }
            PreAuthNotRequired           = if ($PSBoundParameters.ContainsKey('PreAuthNotRequired')) { $PreAuthNotRequired.IsPresent } else { $null }
            DESEncryptionOnly            = if ($PSBoundParameters.ContainsKey('DESEncryptionOnly')) { $DESEncryptionOnly.IsPresent } else { $null }
            Kerberoastable               = if ($PSBoundParameters.ContainsKey('Kerberoastable')) { $Kerberoastable.IsPresent } else { $null }
            DelegatableAdmins            = if ($PSBoundParameters.ContainsKey('DelegatableAdmins')) { $DelegatableAdmins.IsPresent } else { $null }
            SmartCardUsersWithPassword   = if ($PSBoundParameters.ContainsKey('SmartCardUsersWithPassword')) { $SmartCardUsersWithPassword.IsPresent } else { $null }
            DuplicatePasswordGroups      = if ($PSBoundParameters.ContainsKey('DuplicatePasswordGroups')) { $DuplicatePasswordGroups.IsPresent } else { $null }
            DuplicatePasswordGroupsCount = if ($PSBoundParameters.ContainsKey('DuplicatePasswordGroupsCount')) { $DuplicatePasswordGroupsCount } else { $null }
            DuplicatePasswordGroupsType  = if ($PSBoundParameters.ContainsKey('DuplicatePasswordGroupsType')) { $DuplicatePasswordGroupsType } else { $null }
            OrganizationalUnit           = if ($PSBoundParameters.ContainsKey('OrganizationalUnit')) { $OrganizationalUnit } else { $null }
            MemberOf                     = if ($PSBoundParameters.ContainsKey('MemberOf')) { $MemberOf } else { $null }
            Country                      = if ($PSBoundParameters.ContainsKey('Country')) { $Country } else { $null }
        }
    }
    Remove-EmptyValue -Hashtable $Data.Settings
    if ($Data.Settings.Keys.Count -lt 4) {
        Write-Color -Text '[e]', ' At least one parameter must be specified.' -Color Yellow, DarkGray, Yellow, DarkGray, Magenta
        return
    }
    $Data
}