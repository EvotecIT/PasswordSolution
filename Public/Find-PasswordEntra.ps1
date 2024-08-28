function Find-PasswordEntra {
    [CmdletBinding()]
    param(
        [Parameter(DontShow)][string] $HashtableField = 'UserPrincipalName',
        [Parameter(DontShow)][switch] $AsHashTable,
        [string] $OverwriteEmailProperty,
        [Parameter(DontShow)][string[]] $AddEmptyProperties = @(),
        [Parameter(DontShow)][string[]] $RulesProperties,
        [string] $OverwriteManagerProperty,
        [System.Collections.IDictionary] $Cache = [ordered] @{},
        [System.Collections.IDictionary] $CacheManager = [ordered] @{},
        [Parameter(DontShow)][System.Collections.IDictionary] $UsersExternalSystem,
        [Parameter(DontShow)][System.Collections.IDictionary] $ExternalSystemReplacements = [ordered] @{
            Managers = [System.Collections.Generic.List[PSCustomObject]]::new()
            Users    = [System.Collections.Generic.List[PSCustomObject]]::new()
        },
        [string[]] $FilterOrganizationalUnit
    )

    $ExternalSystemManagers = [ordered]@{}
    if ($UsersExternalSystem.Name) {
        Write-Color -Text '[i] ', "Using external system ", $UsersExternalSystem.Name, " for EMAIL replacement functionality" -Color Yellow, White, Yellow, White
        Write-Color -Text '[i] ', "There are ", $UsersExternalSystem.Users.Count, " users in the external system" -Color Yellow, White, Yellow, White
    }
    if (-not $ExternalSystemReplacements.Users) {
        $ExternalSystemReplacements.Users = [System.Collections.Generic.List[PSCustomObject]]::new()
    }
    if (-not $ExternalSystemReplacements.Managers) {
        $ExternalSystemReplacements.Managers = [System.Collections.Generic.List[PSCustomObject]]::new()
    }


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
        'DisplayName', 'GivenName', 'Surname', 'Mail', 'UserPrincipalName', 'Id'
        'lastPasswordChangeDateTime', 'signInActivity'
        'country', 'AccountEnabled'
        'Manager', 'passwordPolicies', 'passwordProfile',
        'OnPremisesDistinguishedName', 'OnPremisesSyncEnabled', 'OnPremisesLastSyncDateTime', 'OnPremisesSamAccountName', 'UserType'
        'assignedLicenses'
        if ($UsersExternalSystem -and $UsersExternalSystem.Type -eq 'ExternalUsers') {
            $UsersExternalSystem.ActiveDirectoryProperty
        }
        if ($OverwriteEmailProperty) {
            $OverwriteEmailProperty
        }
        if ($OverwriteManagerProperty) {
            $OverwriteManagerProperty
        }
        foreach ($Rule in $RulesProperties) {
            $Rule
        }
    )

    $Properties = $Properties | Sort-Object -Unique
    # lets build extended properties that need
    [Array] $ExtendedProperties = foreach ($Rule in $RulesProperties) {
        $Rule
    }
    [Array] $ExtendedProperties = $ExtendedProperties | Sort-Object -Unique

    <# 'signInActivity'
    LastNonInteractiveSignInDateTime LastNonInteractiveSignInRequestId    LastSignInDateTime  LastSignInRequestId
-------------------------------- ---------------------------------    ------------------  -------------------
    10.05.2022 21:50:17              66e349fd-2768-4f0c-811f-ce49219f6300 16.07.2020 11:16:38 108a99e5-b958-4071-8a11-3330c808d700
    #>
    # $Users[-2].Manager.AdditionalProperties


    try {
        $PasswordPolicies = Get-MgDomain -ErrorAction Stop
    } catch {
        Write-Color -Text '[-] ', "Couldn't get password policies. Unable to asses. Error: ", $_.Exception.Message -Color Yellow, White, Red
        return
    }
    if ($PasswordPolicies) {
        $PasswordPolicies = $PasswordPolicies.PasswordValidityPeriodInDays | Select-Object -First 1
    } else {
        Write-Color -Text '[-] ', "Couldn't get password policies. Unable to asses." -Color Yellow, White, Red
        return
    }
    if ($PasswordPolicies -eq '2147483647') {
        $GlobalPasswordPolicy = 'PasswordNeverExpires'
        $GlobalPasswordPolicyDays = $null
    } else {
        $GlobalPasswordPolicy = "$PasswordPolicies days"
        $GlobalPasswordPolicyDays = $PasswordPolicies
    }
    Write-Color -Text "[i] ", "Global password policy is set to $GlobalPasswordPolicy" -Color Yellow, White
    Write-Color -Text "[i] ", "Preparing all users for password expirations in EntraID" -Color Yellow, White, Yellow, White
    try {
        # Get only members, not guests or other types -Filter "userType eq 'member'"
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
        #$HasMailbox = $null



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

        $IsSynchronized = $null -ne $User.OnPremisesDistinguishedName
        $IsLicensed = $User.AssignedLicenses.Count -gt 0

        if ($User.lastPasswordChangeDateTime) {
            $PasswordLastSet = $User.lastPasswordChangeDateTime
            $PasswordDays = ($Today - $PasswordLastSet).Days
        }


        <#
        This value is an enumeration with one possible value being DisableStrongPassword,
        which allows weaker passwords than the default policy to be specified.
        DisablePasswordExpiration can also be specified.
        The two may be specified together; for example:
        DisablePasswordExpiration, DisableStrongPassword.
        For more information on the default password policies, see Microsoft Entra password policies.
        Supports $filter (ne, not, and eq on null values).
        #>
        if ($null -eq $User.PasswordPolicies -or $User.PasswordPolicies -eq 'None') {
            If ($GlobalPasswordPolicy -contains 'PasswordNeverExpires') {
                $PasswordNeverExpires = $true
                $DaysToExpire = $null
                $DateExpiry = $null
            } else {
                $PasswordNeverExpires = $false
                try {
                    # Get the date when password expires based on PasswordLastSet and GlobalPasswordPolicyDays
                    $DateExpiry = $PasswordLastSet.AddDays($GlobalPasswordPolicyDays)
                    $DaysToExpire = ($DateExpiry - $Today).Days
                } catch {
                    $DaysToExpire = $null
                    $DateExpiry = $null
                }
            }
        } elseif ($User.PasswordPolicies -contains 'DisablePasswordExpiration') {
            $PasswordNeverExpires = $true
            $DaysToExpire = $null
            $DateExpiry = $null
        } else {
            Write-Color -Text '[-] ', "Password policy ($($User.PasswordPolicies)) not supported. We need to investigate what changed" -Color Yellow, White, Red
            return
        }

        if ($PasswordNeverExpires) {
            $PasswordExpired = $false
        } else {
            if ($PasswordDays -gt $GlobalPasswordPolicyDays) {
                $PasswordExpired = $true
            } else {
                $PasswordExpired = $false
            }
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
                UserPrincipalName                    = $User.UserPrincipalName
                SamAccountName                       = $User.OnPremisesSamAccountName
                Domain                               = ConvertFrom-DistinguishedName -DistinguishedName $User.OnPremisesDistinguishedName -ToDomainCN

                RuleName                             = ''
                RuleOptions                          = [System.Collections.Generic.List[string]]::new()
                Enabled                              = $User.AccountEnabled
                IsLicensed                           = $IsLicensed
                EmailAddress                         = $EmailAddress
                SystemEmailAddress                   = $User.Mail

                UserType                             = $User.UserType
                IsSynchronized                       = $IsSynchronized
                PasswordPolicies                     = if ($User.PasswordPolicies) { $User.PasswordPolicies } else { 'Not set' }
                ForceChangePasswordNextSignIn        = $User.PasswordProfile.ForceChangePasswordNextSignIn
                ForceChangePasswordNextSignInWithMfa = $User.PasswordProfile.ForceChangePasswordNextSignInWithMfa
                DateExpiry                           = $DateExpiry
                DaysToExpire                         = $DaysToExpire
                PasswordExpired                      = $PasswordExpired
                PasswordDays                         = $PasswordDays
                PasswordAtNextLogon                  = $PasswordAtNextLogon
                PasswordLastSet                      = $User.lastPasswordChangeDateTime
                PasswordNeverExpires                 = $PasswordNeverExpires
                LastLogonDate                        = $LastLogonDate
                LastLogonDays                        = $LastLogonDays
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
                UserPrincipalName                    = $User.UserPrincipalName
                SamAccountName                       = $User.OnPremisesSamAccountName
                Domain                               = ConvertFrom-DistinguishedName -DistinguishedName $User.OnPremisesDistinguishedName -ToDomainCN
                RuleName                             = ''
                RuleOptions                          = [System.Collections.Generic.List[string]]::new()
                Enabled                              = $User.AccountEnabled
                IsLicensed                           = $IsLicensed
                EmailAddress                         = $EmailAddress
                SystemEmailAddress                   = $User.Mail

                UserType                             = $User.UserType
                IsSynchronized                       = $IsSynchronized
                PasswordPolicies                     = if ($User.PasswordPolicies) { $User.PasswordPolicies } else { 'Not set' }
                ForceChangePasswordNextSignIn        = $User.PasswordProfile.ForceChangePasswordNextSignIn
                ForceChangePasswordNextSignInWithMfa = $User.PasswordProfile.ForceChangePasswordNextSignInWithMfa
                DateExpiry                           = $DateExpiry
                DaysToExpire                         = $DaysToExpire
                PasswordExpired                      = $PasswordExpired
                PasswordDays                         = $PasswordDays
                PasswordAtNextLogon                  = $PasswordAtNextLogon
                PasswordLastSet                      = $User.lastPasswordChangeDateTime
                PasswordNeverExpires                 = $PasswordNeverExpires
                LastLogonDate                        = $LastLogonDate
                LastLogonDays                        = $LastLogonDays
                Manager                              = $Manager
                ManagerDisplayName                   = $ManagerDisplayName
                ManagerSamAccountName                = $ManagerSamAccountName
                ManagerEmail                         = $ManagerEmail
                ManagerStatus                        = $ManagerStatus
                ManagerLastLogonDays                 = $ManagerLastLogonDays
                ManagerType                          = $ManagerType
                DisplayName                          = $User.DisplayName
                Name                                 = $User.Name
                GivenName                            = $User.GivenName
                Surname                              = $User.Surname
                OrganizationalUnit                   = ConvertFrom-DistinguishedName -DistinguishedName $User.OnPremisesDistinguishedName -ToOrganizationalUnit
                MemberOf                             = $User.MemberOf
                DistinguishedName                    = $User.OnPremisesDistinguishedName
                ManagerDN                            = $User.Manager
                Country                              = $Country
                CountryCode                          = $CountryCode
                Type                                 = 'User'
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