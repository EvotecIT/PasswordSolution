function Find-PasswordEntra {
    [CmdletBinding()]
    param(
        [Parameter(DontShow)][string] $HashtableField = 'UserPrincipalName',
        [Parameter(DontShow)][switch] $AsHashTable,
        [Parameter(DontShow)][System.Collections.IDictionary] $UsersExternalSystem
    )

    $Today = Get-Date

    # We're caching all users in their inital form to make sure it's speedy gonzales when querying for Managers
    if (-not $Cache) {
        $Cache = [ordered] @{ }
    }
    # We're caching all processed users to make sure it's easier later on to find users
    if (-not $CachedUsers) {
        $CachedUsers = [ordered] @{ }
    }

    $Properties = @(
        # 'Manager', 'DisplayName', 'GivenName', 'Surname', 'SamAccountName', 'EmailAddress',
        # 'msDS-UserPasswordExpiryTimeComputed', 'PasswordExpired', 'PasswordLastSet', 'PasswordNotRequired',
        # 'Enabled', 'PasswordNeverExpires', 'Mail', 'MemberOf', 'LastLogonDate', 'Name'
        # 'userAccountControl'
        # 'pwdLastSet', 'ObjectClass'
        # 'LastLogonDate'
        # 'Country'
        'DisplayName', 'GivenName', 'Surname', 'Mail', 'UserPrincipalName', 'Id'
        'lastPasswordChangeDateTime', 'signInActivity'
        'country', 'AccountEnabled'
        'Manager', 'passwordPolicies', 'passwordProfile',
        'OnPremisesDistinguishedName', 'OnPremisesSyncEnabled', 'OnPremisesLastSyncDateTime', 'OnPremisesSamAccountName', 'UserType'
    )

    <#
    - passwordPolicies
    Specifies password policies for the user. This value is an enumeration with one possible value being DisableStrongPassword, which allows weaker passwords than the default policy to be specified. DisablePasswordExpiration can also be specified. The two may be specified together; for example: DisablePasswordExpiration, DisableStrongPassword. For more information on the default password policies, see Microsoft Entra password policies.
    Supports $filter (ne, not, and eq on null values).

    - passwordProfile
    Specifies the password profile for the user. The profile contains the user's password. This property is required when a user is created. The password in the profile must satisfy minimum requirements as specified by the passwordPolicies property. By default, a strong password is required.
    #>

    $Properties = $Properties | Sort-Object -Unique

    <# 'signInActivity'
    LastNonInteractiveSignInDateTime LastNonInteractiveSignInRequestId    LastSignInDateTime  LastSignInRequestId
-------------------------------- ---------------------------------    ------------------  -------------------
    10.05.2022 21:50:17              66e349fd-2768-4f0c-811f-ce49219f6300 16.07.2020 11:16:38 108a99e5-b958-4071-8a11-3330c808d700
    #>
    # $Users[-2].Manager.AdditionalProperties


    $PasswordPolicies = Get-MgDomain
    if ($PasswordPolicies) {
        $PasswordPolicies = $PasswordPolicies.PasswordValidityPeriodInDays | Select-Object -First 1
    } else {
        Write-Color -Text '[-] ', "Couldn't get password policies. Unable to asses." -Color Yellow, White, Red
        return
    }
    if ($PasswordPolicies -eq '2147483647') {
        $GlobalPasswordPolicy = 'PasswordNeverExpires'
    } else {
        $GlobalPasswordPolicy = $PasswordPolicies + ' days'
    }
    Write-Color -Text "[i] ", "Global password policy is set to $GlobalPasswordPolicy" -Color Yellow, White


    Write-Color -Text "[i] ", "Preparing all users for password expirations in EntraID" -Color Yellow, White, Yellow, White
    try {
        $Users = Get-MgUser -All -ErrorAction Stop -Property $Properties -ConsistencyLevel eventual -ExpandProperty Manager | Select-Object -Property $Properties
    } catch {
        Write-Color -Text '[-] ', "Couldn't cache users. Please fix 'Find-PasswordEntra'. Error: ", "$($_.Exception.Message)" -Color Yellow, White, Red
        return
    }

    $CountUsers = 0
    foreach ($User in $Users) {
        $CountUsers++
        Write-Verbose -Message "Processing $($User.DisplayName) - $($CountUsers)/$($Users.Count)"

        $LastLogonDate = $null
        $LastLogonDays = $null
        if ($User.SignInActivity) {
            if ($User.SignInActivity -and $User.LastNonInteractiveSignInDateTime) {
                if ($User.SignInActivity.LastNonInteractiveSignInDateTime -gt $User.SignInActivity.LastSignInDateTime) {
                    $LastLogonDate = $User.SignInActivity.LastNonInteractiveSignInDateTime
                } else {
                    $LastLogonDate = $User.SignInActivity.LastSignInDateTime
                }
            } else {
                if ($User.SignInActivity.LastSignInDateTime) {
                    $LastLogonDate = $User.SignInActivity.LastSignInDateTime
                } elseif ($User.SignInActivity.LastNonInteractiveSignInDateTime) {
                    $LastLogonDate = $User.SignInActivity.LastNonInteractiveSignInDateTime
                }
            }
            if ($null -ne $LastLogonDate) {
                $LastLogonDays = ($Today - $LastLogonDate).Days
            }
        }


        $DateExpiry = $null
        $DaysToExpire = $null
        $PasswordDays = $null
        $PasswordNeverExpires = $false
        $PasswordAtNextLogon = $null
        $HasMailbox = $null



        $Country = $User.Country
        if ($Country) {
            $CountryCode = Convert-CountryToCountryCode -CountryName $User.Country
        } else {
            $CountryCode = $null
        }

        # if ($User.ObjectClass -eq 'user') {
        #     $ManagerStatus = 'Missing'
        # } else {
        #     $ManagerStatus = 'Not available'
        # }
        if ($User.Manager.AdditionalProperties) {

            $Manager = $User.Manager.AdditionalProperties
            $ManagerSamAccountName = $User.Manager.AdditionalProperties.onPremisesSamAccountName
            $ManagerDisplayName = $User.Manager.AdditionalProperties.displayName
            $ManagerEmail = $User.Manager.AdditionalProperties.mail
            $ManagerEnabled = $User.Manager.AdditionalProperties.accountEnabled
            #$ManagerLastLogon = $null
            #$ManagerLastLogonDays = $null
            $ManagerType = $User.Manager.AdditionalProperties.userType

            if ($ManagerEnabled -and $ManagerEmail) {
                if ((Test-EmailAddress -EmailAddress $ManagerEmail).IsValid -eq $true) {
                    $ManagerStatus = 'Enabled'
                } else {
                    $ManagerStatus = 'Enabled, bad email'
                }
            } elseif ($ManagerEnabled -eq $true) {
                $ManagerStatus = 'No email'
            } elseif ($ManagerEnabled -eq $false) {
                $ManagerStatus = 'Disabled'
            } else {
                $ManagerStatus = 'Missing'
            }

        } else {
            $Manager = $null
            $ManagerSamAccountName = $null
            $ManagerDisplayName = $null
            $ManagerEmail = $nullf
            $ManagerStatus = 'Missing'
            $ManagerLastLogonDays = $null
            $ManagerType = $null
        }

        if ($User.lastPasswordChangeDateTime) {
            $PasswordLastSet = $User.lastPasswordChangeDateTime
            $PasswordDays = ($Today - $PasswordLastSet).Days
        }
        if ($User.PasswordPolicies -contains 'DisablePasswordExpiration') {
            $PasswordNeverExpires = $true
        }

        if ($OverwriteEmailProperty) {
            # fix this for a user
            $EmailTemp = $User.$OverwriteEmailProperty
            if ($EmailTemp -like '*@*') {
                $EmailAddress = $EmailTemp
            } else {
                $EmailAddress = $User.Mail
            }
            # Fix this for manager as well
            if ($Cache["$($User.Manager)"]) {
                if ($Cache["$($User.Manager)"].$OverwriteEmailProperty -like '*@*') {
                    # $UserManager.Mail = $UserManager.$OverwriteEmailProperty
                    $ManagerEmail = $Cache["$($User.Manager)"].$OverwriteEmailProperty
                }
            }
        } else {
            $EmailAddress = $User.Mail
        }

        if ($UsersExternalSystem -and $UsersExternalSystem.Global -eq $true) {
            if ($UsersExternalSystem.Type -eq 'ExternalUsers') {
                $ADProperty = $UsersExternalSystem.ActiveDirectoryProperty
                $EmailProperty = $UsersExternalSystem.EmailProperty
                $ExternalUser = $UsersExternalSystem['Users'][$User.$ADProperty]
                if ($ExternalUser -and $ExternalUser.$EmailProperty -like '*@*') {
                    $EmailAddress = $ExternalUser.$EmailProperty
                } else {
                    $EmailAddress = $User.Mail
                }
            } else {
                Write-Color -Text '[-] ', "External system type not supported. Please use only type as provided using 'New-PasswordConfigurationExternalUsers'." -Color Yellow, White, Red
                return
            }
        }

        if ($AddEmptyProperties.Count -gt 0) {
            $StartUser = [ordered] @{
                UserPrincipalName    = $User.UserPrincipalName
                #SamAccountName       = $User.SamAccountName
                #Domain               = ConvertFrom-DistinguishedName -DistinguishedName $User.DistinguishedName -ToDomainCN
                RuleName             = ''
                RuleOptions          = [System.Collections.Generic.List[string]]::new()
                Enabled              = $User.AccountEnabled
                HasMailbox           = $HasMailbox
                EmailAddress         = $EmailAddress
                SystemEmailAddress   = $User.Mail
                DateExpiry           = $DateExpiry
                DaysToExpire         = $DaysToExpire
                PasswordExpired      = $User.PasswordExpired
                PasswordDays         = $PasswordDays
                PasswordAtNextLogon  = $PasswordAtNextLogon
                PasswordLastSet      = $User.lastPasswordChangeDateTime
                #PasswordNotRequired  = $User.PasswordNotRequired
                PasswordNeverExpires = $PasswordNeverExpires
                LastLogonDate        = $LastLogonDate
                LastLogonDays        = $LastLogonDays
            }
            foreach ($Property in $AddEmptyProperties) {
                $StartUser.$Property = $null
            }
            $EndUser = [ordered] @{
                Manager               = $Manager
                ManagerDisplayName    = $ManagerDisplayName
                ManagerSamAccountName = $ManagerSamAccountName
                ManagerEmail          = $ManagerEmail
                ManagerStatus         = $ManagerStatus
                ManagerLastLogonDays  = $ManagerLastLogonDays
                ManagerType           = $ManagerType
                DisplayName           = $User.DisplayName
                Name                  = $User.Name
                GivenName             = $User.GivenName
                Surname               = $User.Surname
                OrganizationalUnit    = ConvertFrom-DistinguishedName -DistinguishedName $User.OnPremisesDistinguishedName -ToOrganizationalUnit
                MemberOf              = $User.MemberOf
                DistinguishedName     = $User.OnPremisesDistinguishedName
                ManagerDN             = $User.Manager
                Country               = $Country
                CountryCode           = $CountryCode
                Type                  = 'User'
            }
            $MyUser = $StartUser + $EndUser
        } else {
            $MyUser = [ordered] @{
                UserPrincipalName     = $User.UserPrincipalName
                # SamAccountName        = $User.SamAccountName
                # Domain                = ConvertFrom-DistinguishedName -DistinguishedName $User.DistinguishedName -ToDomainCN
                RuleName              = ''
                RuleOptions           = [System.Collections.Generic.List[string]]::new()
                Enabled               = $User.AccountEnabled
                HasMailbox            = $HasMailbox
                EmailAddress          = $EmailAddress
                SystemEmailAddress    = $User.Mail
                DateExpiry            = $DateExpiry
                DaysToExpire          = $DaysToExpire
                PasswordExpired       = $User.PasswordExpired
                PasswordDays          = $PasswordDays
                PasswordAtNextLogon   = $PasswordAtNextLogon
                PasswordLastSet       = $User.lastPasswordChangeDateTime
                #PasswordNotRequired   = $User.PasswordNotRequired
                PasswordNeverExpires  = $PasswordNeverExpires
                LastLogonDate         = $LastLogonDate
                LastLogonDays         = $LastLogonDays
                Manager               = $Manager
                ManagerDisplayName    = $ManagerDisplayName
                ManagerSamAccountName = $ManagerSamAccountName
                ManagerEmail          = $ManagerEmail
                ManagerStatus         = $ManagerStatus
                ManagerLastLogonDays  = $ManagerLastLogonDays
                ManagerType           = $ManagerType
                DisplayName           = $User.DisplayName
                Name                  = $User.Name
                GivenName             = $User.GivenName
                Surname               = $User.Surname
                OrganizationalUnit    = ConvertFrom-DistinguishedName -DistinguishedName $User.OnPremisesDistinguishedName -ToOrganizationalUnit
                MemberOf              = $User.MemberOf
                DistinguishedName     = $User.OnPremisesDistinguishedName
                ManagerDN             = $User.Manager
                Country               = $Country
                CountryCode           = $CountryCode
                Type                  = 'User'
            }
        }
        foreach ($Property in $ConditionProperties) {
            $MyUser["$Property"] = $User.$Property
        }
        foreach ($E in $ExtendedProperties) {
            $MyUser[$E] = $User.$E
        }
        if ($HashtableField -eq 'NetBiosSamAccountName') {
            $HashField = $DNSNetBios[$MyUser.Domain] + '\' + $MyUser.SamAccountName
            if ($AsHashTableObject) {
                $CachedUsers["$HashField"] = $MyUser
            } else {
                $CachedUsers["$HashField"] = [PSCustomObject] $MyUser
            }
        } else {
            if ($AsHashTableObject) {
                $CachedUsers["$($User.$HashtableField)"] = $MyUser
            } else {
                $CachedUsers["$($User.$HashtableField)"] = [PSCustomObject] $MyUser
            }
        }
    }
    if ($AsHashTable) {
        $CachedUsers
    } else {
        $CachedUsers.Values
    }
}