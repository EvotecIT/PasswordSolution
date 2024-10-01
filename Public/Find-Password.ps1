function Find-Password {
    <#
    .SYNOPSIS
    Scan Active Directory forest for all users and their password expiration date

    .DESCRIPTION
    Scan Active Directory forest for all users and their password expiration date

    .PARAMETER Forest
    Target different Forest, by default current forest is used

    .PARAMETER ExcludeDomains
    Exclude domain from search, by default whole forest is scanned

    .PARAMETER IncludeDomains
    Include only specific domains, by default whole forest is scanned

    .PARAMETER ExtendedForestInformation
    Ability to provide Forest Information from another command to speed up processing

    .PARAMETER OverwriteEmailProperty
    Overwrite EmailAddress property with different property name

    .PARAMETER OverwriteManagerProperty
    Overwrite Manager property with different property name.
    Can use DistinguishedName or SamAccountName

    .PARAMETER RulesProperties
    Add additional properties to be returned from rules

    .PARAMETER UsersExternalSystem

    .PARAMETER ExternalSystemReplacements

    .PARAMETER FilterOrganizationalUnit

    .PARAMETER SearchBase

    .PARAMETER AsHashTable

    .PARAMETER AsHashTableObject

    .PARAMETER AddEmptyProperties

    .PARAMETER ReturnObjectsType

    .PARAMETER Cache

    .PARAMETER HashtableField

    .PARAMETER CacheManager

    .EXAMPLE
    Find-Password | ft

    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [alias('ForestName')][string] $Forest,
        [string[]] $ExcludeDomains,
        [alias('Domain', 'Domains')][string[]] $IncludeDomains,
        [System.Collections.IDictionary] $ExtendedForestInformation,
        [string] $OverwriteEmailProperty,
        [Parameter(DontShow)][switch] $AsHashTable,
        [Parameter(DontShow)][string] $HashtableField = 'DistinguishedName',
        [ValidateSet('Users', 'Contacts')][string[]] $ReturnObjectsType = @('Users', 'Contacts'),
        [Parameter(DontShow)][switch] $AsHashTableObject,
        [Parameter(DontShow)][string[]] $AddEmptyProperties = @(),
        [Parameter(DontShow)][string[]] $RulesProperties,
        [string] $OverwriteManagerProperty,
        [Parameter(DontShow)][System.Collections.IDictionary] $UsersExternalSystem,
        [Parameter(DontShow)][System.Collections.IDictionary] $ExternalSystemReplacements = [ordered] @{
            Managers = [System.Collections.Generic.List[PSCustomObject]]::new()
            Users    = [System.Collections.Generic.List[PSCustomObject]]::new()
        },
        [string[]] $FilterOrganizationalUnit,
        [string[]] $SearchBase,
        [System.Collections.IDictionary] $Cache = [ordered] @{},
        [System.Collections.IDictionary] $CacheManager = [ordered] @{}
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

    $GuidForExchange = Convert-ADSchemaToGuid -SchemaName 'msExchMailboxGuid'
    if ($GuidForExchange) {
        $ExchangeProperty = 'msExchMailboxGuid'
    }

    $Properties = @(
        'Manager', 'DisplayName', 'GivenName', 'Surname', 'SamAccountName', 'EmailAddress',
        'msDS-UserPasswordExpiryTimeComputed', 'PasswordExpired', 'PasswordLastSet', 'PasswordNotRequired',
        'Enabled', 'PasswordNeverExpires', 'Mail', 'MemberOf', 'LastLogonDate', 'Name'
        'userAccountControl'
        'pwdLastSet', 'ObjectClass'
        'LastLogonDate'
        'Country'
        if ($UsersExternalSystem -and $UsersExternalSystem.Type -eq 'ExternalUsers') {
            $UsersExternalSystem.ActiveDirectoryProperty
        }
        if ($ExchangeProperty) {
            $ExchangeProperty
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

    $PropertiesContacts = @(
        'SamAccountName', 'CanonicalName', 'WhenChanged', 'WhenChanged', 'DisplayName', 'DistinguishedName', 'Name', 'Mail', 'TargetAddress', 'ObjectClass'
    )

    # We're caching all users in their inital form to make sure it's speedy gonzales when querying for Managers
    if (-not $Cache) {
        $Cache = [ordered] @{ }
    }
    # We're caching all processed users to make sure it's easier later on to find users
    if (-not $CachedUsers) {
        $CachedUsers = [ordered] @{ }
    }
    Write-Color -Text '[i] ', "Discovering forest information" -Color Yellow, White
    $ForestInformation = Get-WinADForestDetails -PreferWritable -Extended -Forest $Forest -ExcludeDomains $ExcludeDomains -IncludeDomains $IncludeDomains -ExtendedForestInformation $ExtendedForestInformation

    $SearchBaseCache = [ordered]@{}
    if ($SearchBase) {
        foreach ($S in $SearchBase) {
            $ConvertedS = ConvertFrom-DistinguishedName -DistinguishedName $S -ToDomainCN
            if (-not $SearchBaseCache[$ConvertedS]) {
                $SearchBaseCache[$ConvertedS] = [System.Collections.Generic.List[string]]::new()
            }
            $SearchBaseCache[$ConvertedS].Add($S)
        }
    }


    # lets get domain name / netbios hashtable for easy use
    $DNSNetBios = @{ }
    foreach ($NETBIOS in $ForestInformation.DomainsExtendedNetBIOS.Keys) {
        $DNSNetBios[$ForestInformation.DomainsExtendedNetBIOS[$NETBIOS].DnsRoot] = $NETBIOS
    }

    [Array] $Users = foreach ($Domain in $ForestInformation.Domains) {
        Write-Color -Text "[i] ", "Discovering DC for domain ", "$($Domain)", " in forest ", $ForestInformation.Name -Color Yellow, White, Yellow, White
        $Server = $ForestInformation['QueryServers'][$Domain]['HostName'][0]

        if ($SearchBase) {
            foreach ($SB in $SearchBaseCache[$Domain]) {
                Write-Color -Text "[i] ", "Getting users from ", "$($Domain)", " using ", $Server, " and SearchBase ", $SB -Color Yellow, White, Yellow, White, Yellow, White, Yellow
                try {
                    Get-ADUser -Server $Server -Filter '*' -SearchBase $SB -Properties $Properties -ErrorAction Stop
                } catch {
                    $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
                    Write-Color '[e] Error: ', $ErrorMessage -Color White, Red
                }
            }
        } else {
            Write-Color -Text "[i] ", "Getting users from ", "$($Domain)", " using ", $Server -Color Yellow, White, Yellow, White
            try {
                Get-ADUser -Server $Server -Filter '*' -Properties $Properties -ErrorAction Stop
            } catch {
                $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
                Write-Color '[e] Error: ', $ErrorMessage -Color White, Red
            }
        }
    }
    Write-Color -Text "[i] ", "Caching users for easy access" -Color Yellow, White
    foreach ($User in $Users) {
        $Cache[$User.DistinguishedName] = $User
        # SAmAccountName will overwrite itself when we have multiple domains and there are duplicates
        # but sicne we use only on in case manager is used in special fields such as extensionAttribute, it shouldn't affect much
        $Cache[$User.SamAccountName] = $User
    }

    if ($ReturnObjectsType -contains 'Contacts') {
        [Array] $Contacts = foreach ($Domain in $ForestInformation.Domains) {
            Write-Color -Text "[i] ", "Discovering DC for domain ", "$($Domain)", " in forest ", $ForestInformation.Name -Color Yellow, White, Yellow, White
            $Server = $ForestInformation['QueryServers'][$Domain]['HostName'][0]

            if ($SearchBase) {
                foreach ($SB in $SearchBaseCache[$Domain]) {
                    Write-Color -Text "[i] ", "Getting contacts from ", "$($Domain)", " using ", $Server, " and SearchBase ", $SB -Color Yellow, White, Yellow, White, Yellow, White, Yellow
                    try {
                        Get-ADObject -LDAPFilter "objectClass=Contact" -Server $Server -SearchBase $SB -Properties $PropertiesContacts -ErrorAction Stop
                    } catch {
                        $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
                        Write-Color '[e] Error: ', $ErrorMessage -Color White, Red
                    }
                }
            } else {
                Write-Color -Text "[i] ", "Getting contacts from ", "$($Domain)", " using ", $Server -Color Yellow, White, Yellow, White
                try {
                    Get-ADObject -LDAPFilter "objectClass=Contact" -Server $Server -Properties $PropertiesContacts -ErrorAction Stop
                } catch {
                    $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
                    Write-Color '[e] Error: ', $ErrorMessage -Color White, Red
                }
            }
        }
        foreach ($Contact in $Contacts) {
            $Cache[$Contact.DistinguishedName] = $Contact
        }
    }

    Write-Color -Text "[i] ", "Preparing users ", $Users.Count, " for password expirations in forest ", $Forest.Name -Color Yellow, White, Yellow, White, Yellow, White
    foreach ($OU in $FilterOrganizationalUnit) {
        Write-Color -Text "[i] ", "Filtering users by Organizational Unit ", $OU -Color Yellow, White, Yellow, White
    }
    $CountUsers = 0
    foreach ($User in $Users) {
        $CountUsers++
        Write-Verbose -Message "Processing $($User.DisplayName) / $($User.DistinguishedName) - $($CountUsers)/$($Users.Count)"
        $SkipUser = $false
        $DateExpiry = $null
        $DaysToExpire = $null
        $PasswordDays = $null
        $PasswordNeverExpires = $null
        $PasswordAtNextLogon = $null
        $HasMailbox = $null

        $OUPath = ConvertFrom-DistinguishedName -DistinguishedName $User.DistinguishedName -ToOrganizationalUnit
        # Allow filtering to prevent huge time processing for huge domains when only some users are needed
        # from specific Organizational Units
        foreach ($OU in $FilterOrganizationalUnit) {
            if ($null -ne $OUPath -and $OUPath -like "$OU") {
                $SkipUser = $false
                break
            } else {
                $SkipUser = $true
            }
        }
        if ($SkipUser) {
            continue
        }

        # This is a special case for users that have a manager in a special field such as extensionAttributes
        # This is useful for service accounts or other accounts that don't have a manager in AD
        if ($OverwriteManagerProperty) {
            # fix this for a user
            $ManagerTemp = $User.$OverwriteManagerProperty
            if ($ManagerTemp) {
                $ManagerSpecial = $Cache[$ManagerTemp]
            } else {
                $ManagerSpecial = $null
            }
        } else {
            $ManagerSpecial = $null
        }

        if ($ManagerSpecial) {
            # We have manager in different field such as extensionAttribute
            $ManagerDN = $ManagerSpecial.DistinguishedName
            $Manager = $ManagerSpecial.DisplayName
            $ManagerSamAccountName = $ManagerSpecial.SamAccountName
            $ManagerDisplayName = $ManagerSpecial.DisplayName
            $ManagerEmail = $ManagerSpecial.Mail

            # we check if we have external system and if we have global email replacement for managers in place
            # we check only if SamAccountName is there (contacts don't have it)
            if ($ManagerSamAccountName -and $UsersExternalSystem -and $UsersExternalSystem.Global -eq $true) {
                $ADProperty = $UsersExternalSystem.ActiveDirectoryProperty
                if ($ADProperty -eq 'SamAccountName') {
                    # we need to find manager by SamAccountName, and we need to find it in external system
                    # any other property is not supported
                    $EmailProperty = $UsersExternalSystem.EmailProperty
                    $ExternalUser = $UsersExternalSystem['Users'][$ManagerSamAccountName]
                    if ($ExternalUser -and $ExternalUser.$EmailProperty -like '*@*' -and $ExternalUser.$EmailProperty -ne $ManagerEmail) {
                        $ReplacedManagerEmail = $ManagerEmail
                        $ManagerEmail = $ExternalUser.$EmailProperty

                        if (-not $ExternalSystemManagers[$ManagerSamAccountName]) {
                            $ExternalSystemManagers[$ManagerSamAccountName] = $ManagerSamAccountName
                            $ExternalSystemReplacements.Managers.Add(
                                [PSCustomObject]@{
                                    ManagerSamAccountName = $ManagerSamAccountName
                                    ExternalEmail         = $ManagerEmail
                                    ADEmailAddress        = $ReplacedManagerEmail
                                    ExternalSystem        = $UsersExternalSystem.Name
                                }
                            )
                        }
                        #Write-Color -Text '[i] ', "Overwriting manager email address for ", $Manager, " with ", $ManagerEmail, " (old email: $ReplacedManagerEmail)", " from ", $UsersExternalSystem.Name -Color Yellow, White, Yellow, White, Green, Red, White, Yellow
                    }
                }
            }

            $ManagerEnabled = $ManagerSpecial.Enabled
            $ManagerLastLogon = $ManagerSpecial.LastLogonDate
            if ($ManagerLastLogon) {
                $ManagerLastLogonDays = $( - $($ManagerLastLogon - $Today).Days)
            } else {
                $ManagerLastLogonDays = $null
            }
            $ManagerType = $ManagerSpecial.ObjectClass
        } elseif ($User.Manager) {
            $ManagerDN = $Cache[$User.Manager].DistinguishedName
            $Manager = $Cache[$User.Manager].DisplayName
            $ManagerSamAccountName = $Cache[$User.Manager].SamAccountName
            $ManagerDisplayName = $Cache[$User.Manager].DisplayName
            $ManagerEmail = $Cache[$User.Manager].Mail

            # we check if we have external system and if we have global email replacement for managers in place
            # we check only if SamAccountName is there (contacts don't have it)
            if ($ManagerSamAccountName -and $UsersExternalSystem -and $UsersExternalSystem.Global -eq $true) {
                $ADProperty = $UsersExternalSystem.ActiveDirectoryProperty
                if ($ADProperty -eq 'SamAccountName') {
                    # we need to find manager by SamAccountName, and we need to find it in external system
                    # any other property is not supported
                    $EmailProperty = $UsersExternalSystem.EmailProperty
                    $ExternalUser = $UsersExternalSystem['Users'][$ManagerSamAccountName]
                    if ($ExternalUser -and $ExternalUser.$EmailProperty -like '*@*' -and $ExternalUser.$EmailProperty -ne $ManagerEmail) {
                        $ReplacedManagerEmail = $ManagerEmail
                        $ManagerEmail = $ExternalUser.$EmailProperty
                        if (-not $ExternalSystemManagers[$ManagerSamAccountName]) {
                            $ExternalSystemManagers[$ManagerSamAccountName] = $ManagerSamAccountName
                            $ExternalSystemReplacements.Managers.Add(
                                [PSCustomObject]@{
                                    ManagerSamAccountName = $ManagerSamAccountName
                                    ExternalEmail         = $ManagerEmail
                                    ADEmailAddress        = $ReplacedManagerEmail
                                    ExternalSystem        = $UsersExternalSystem.Name
                                }
                            )
                        }
                        #Write-Color -Text '[i] ', "Overwriting manager email address for ", $Manager, " with ", $ManagerEmail, " (old email: $ReplacedManagerEmail)", " from ", $UsersExternalSystem.Name -Color Yellow, White, Yellow, White, Green, Red, White, Yellow
                    }
                }
            }
            $ManagerEnabled = $Cache[$User.Manager].Enabled
            $ManagerLastLogon = $Cache[$User.Manager].LastLogonDate
            if ($ManagerLastLogon) {
                $ManagerLastLogonDays = $( - $($ManagerLastLogon - $Today).Days)
            } else {
                $ManagerLastLogonDays = $null
            }
            $ManagerType = $Cache[$User.Manager].ObjectClass
        } else {
            if ($User.ObjectClass -eq 'user') {
                $ManagerStatus = 'Missing'
            } else {
                $ManagerStatus = 'Not available'
            }
            $ManagerDN = $null
            $Manager = $null
            $ManagerSamAccountName = $null
            $ManagerDisplayName = $null
            $ManagerEmail = $null
            $ManagerEnabled = $null
            $ManagerLastLogon = $null
            $ManagerLastLogonDays = $null
            $ManagerType = $null
        }

        if ($ManagerDN -and -not $CacheManager[$ManagerDN]) {
            $CacheManager[$ManagerDN] = [PSCustomObject] @{
                DistinguishedName = $ManagerDN
                Domain            = ConvertFrom-DistinguishedName -DistinguishedName $ManagerDN -ToDomainCN
                DisplayName       = $ManagerDisplayName
                SamAccountName    = $ManagerSamAccountName
                EmailAddress      = $ManagerEmail
                Enabled           = $ManagerEnabled
                LastLogonDate     = $ManagerLastLogon
                LastLogonDays     = $ManagerLastLogonDays
                Type              = $ManagerType
            }
        }

        if ($OverwriteEmailProperty) {
            # fix this for a user
            $EmailTemp = $User.$OverwriteEmailProperty
            if ($EmailTemp -like '*@*') {
                $EmailAddress = $EmailTemp
            } else {
                $EmailAddress = $User.EmailAddress
            }
            # Fix this for manager as well
            if ($Cache["$($User.Manager)"]) {
                if ($Cache["$($User.Manager)"].$OverwriteEmailProperty -like '*@*') {
                    # $UserManager.Mail = $UserManager.$OverwriteEmailProperty
                    $ManagerEmail = $Cache["$($User.Manager)"].$OverwriteEmailProperty
                }
            }
        } else {
            $EmailAddress = $User.EmailAddress
        }

        if ($UsersExternalSystem -and $UsersExternalSystem.Global -eq $true) {
            if ($UsersExternalSystem.Type -eq 'ExternalUsers') {
                $ADProperty = $UsersExternalSystem.ActiveDirectoryProperty
                $EmailProperty = $UsersExternalSystem.EmailProperty
                $ExternalUser = $UsersExternalSystem['Users'][$User.$ADProperty]
                # $EmailAddress = $User.EmailAddress
                $EmailFrom = 'AD'
                if ($ExternalUser -and $ExternalUser.$EmailProperty -like '*@*' -and $EmailAddress -ne $ExternalUser.$EmailProperty) {
                    $EmailFrom = 'ILM'
                    $EmailAddress = $ExternalUser.$EmailProperty
                    $ExternalSystemReplacements.Users.Add(
                        [PSCustomObject]@{
                            UserSamAccountName = $User.SamAccountName
                            ExternalEmail      = $EmailAddress
                            ADEmailAddress     = $User.EmailAddress
                            ExternalSystem     = $UsersExternalSystem.Name
                        }
                    )
                }
            } else {
                Write-Color -Text '[-] ', "External system type not supported. Please use only type as provided using 'New-PasswordConfigurationExternalUsers'." -Color Yellow, White, Red
                return
            }
        } else {
            $EmailFrom = 'AD'
        }

        if ($User.PasswordLastSet) {
            $PasswordDays = (New-TimeSpan -Start ($User.PasswordLastSet) -End ($Today)).Days
        } else {
            $PasswordDays = $null
        }

        # Since we fixed manager above, we now check for status
        if ($User.Manager) {
            if ($ManagerEnabled -and $ManagerEmail) {
                if ((Test-EmailAddress -EmailAddress $ManagerEmail).IsValid -eq $true) {
                    $ManagerStatus = 'Enabled'
                } else {
                    $ManagerStatus = 'Enabled, bad email'
                }
            } elseif ($ManagerEnabled) {
                $ManagerStatus = 'No email'
            } elseif ($Cache[$User.Manager].ObjectClass -eq 'Contact') {
                $ManagerStatus = 'Enabled' # we need to treat it as always enabled
            } else {
                $ManagerStatus = 'Disabled'
            }
        }

        if ($User."msDS-UserPasswordExpiryTimeComputed" -ne 9223372036854775807) {
            # This is standard situation where users password is expiring as needed
            try {
                $DateExpiry = ([datetime]::FromFileTime($User."msDS-UserPasswordExpiryTimeComputed"))
            } catch {
                $DateExpiry = $User."msDS-UserPasswordExpiryTimeComputed"
            }
            try {
                $DaysToExpire = (New-TimeSpan -Start ($Today) -End ([datetime]::FromFileTime($User."msDS-UserPasswordExpiryTimeComputed"))).Days
            } catch {
                $DaysToExpire = $null
            }
            $PasswordNeverExpires = $User.PasswordNeverExpires
        } else {
            # This is non-standard situation. This basically means most likely Fine Grained Group Policy is in action where it makes PasswordNeverExpires $true
            # Since FGP policies are a bit special they do not tick the PasswordNeverExpires box, but at the same time value for "msDS-UserPasswordExpiryTimeComputed" is set to 9223372036854775807
            $PasswordNeverExpires = $true
        }

        if ($User.pwdLastSet -eq 0 -and $DateExpiry.Year -eq 1601) {
            $PasswordAtNextLogon = $true
        } else {
            $PasswordAtNextLogon = $false
        }

        if ($PasswordNeverExpires -or $null -eq $User.PasswordLastSet) {
            # If password last set is null or password never expires is set to true, then date of expiry and days to expire is not applicable
            $DateExpiry = $null
            $DaysToExpire = $null
        }

        $UserAccountControl = Convert-UserAccountControl -UserAccountControl $User.UserAccountControl
        if ($UserAccountControl -contains 'INTERDOMAIN_TRUST_ACCOUNT') {
            continue
        }
        if ($ExchangeProperty) {
            if ($User.'msExchMailboxGuid') {
                $HasMailbox = 'Yes'
            } else {
                $HasMailbox = 'No'
            }
        } else {
            $HasMailbox = 'Unknown'
        }
        if ($User.LastLogonDate) {
            $LastLogonDays = $( - $($User.LastLogonDate - $Today).Days)
        } else {
            $LastLogonDays = $null
        }

        if ($User.Country) {
            $Country = Convert-CountryCodeToCountry -CountryCode $User.Country
            $CountryCode = $User.Country
        } else {
            $Country = 'Unknown'
            $CountryCode = 'Unknown'
        }


        if ($AddEmptyProperties.Count -gt 0) {
            $StartUser = [ordered] @{
                UserPrincipalName    = $User.UserPrincipalName
                SamAccountName       = $User.SamAccountName
                Domain               = ConvertFrom-DistinguishedName -DistinguishedName $User.DistinguishedName -ToDomainCN
                RuleName             = ''
                RuleOptions          = [System.Collections.Generic.List[string]]::new()
                Enabled              = $User.Enabled
                HasMailbox           = $HasMailbox
                EmailAddress         = $EmailAddress
                SystemEmailAddress   = $User.EmailAddress
                DateExpiry           = $DateExpiry
                DaysToExpire         = $DaysToExpire
                PasswordExpired      = $User.PasswordExpired
                PasswordDays         = $PasswordDays
                PasswordAtNextLogon  = $PasswordAtNextLogon
                PasswordLastSet      = $User.PasswordLastSet
                PasswordNotRequired  = $User.PasswordNotRequired
                PasswordNeverExpires = $PasswordNeverExpires
                LastLogonDate        = $User.LastLogonDate
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
                OrganizationalUnit    = $OUPath
                MemberOf              = $User.MemberOf
                DistinguishedName     = $User.DistinguishedName
                ManagerDN             = $User.Manager
                Country               = $Country
                CountryCode           = $CountryCode
                Type                  = 'User'
                EmailFrom             = $EmailFrom
            }
            $MyUser = $StartUser + $EndUser
        } else {
            $MyUser = [ordered] @{
                UserPrincipalName     = $User.UserPrincipalName
                SamAccountName        = $User.SamAccountName
                Domain                = ConvertFrom-DistinguishedName -DistinguishedName $User.DistinguishedName -ToDomainCN
                RuleName              = ''
                RuleOptions           = [System.Collections.Generic.List[string]]::new()
                Enabled               = $User.Enabled
                HasMailbox            = $HasMailbox
                EmailAddress          = $EmailAddress
                SystemEmailAddress    = $User.EmailAddress
                DateExpiry            = $DateExpiry
                DaysToExpire          = $DaysToExpire
                PasswordExpired       = $User.PasswordExpired
                PasswordDays          = $PasswordDays
                PasswordAtNextLogon   = $PasswordAtNextLogon
                PasswordLastSet       = $User.PasswordLastSet
                PasswordNotRequired   = $User.PasswordNotRequired
                PasswordNeverExpires  = $PasswordNeverExpires
                LastLogonDate         = $User.LastLogonDate
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
                OrganizationalUnit    = ConvertFrom-DistinguishedName -DistinguishedName $User.DistinguishedName -ToOrganizationalUnit
                MemberOf              = $User.MemberOf
                DistinguishedName     = $User.DistinguishedName
                ManagerDN             = $User.Manager
                Country               = $Country
                CountryCode           = $CountryCode
                Type                  = 'User'
                EmailFrom             = $EmailFrom
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
    if ($ReturnObjectsType -contains 'Contacts') {
        $CountContacts = 0
        foreach ($Contact in $Contacts) {
            $CountContacts++

            $OUPath = ConvertFrom-DistinguishedName -DistinguishedName $Contact.DistinguishedName -ToOrganizationalUnit
            # Allow filtering to prevent huge time processing for huge domains when only some users are needed
            foreach ($OU in $FilterOrganizationalUnit) {
                if ($null -eq $OUPath) {
                    $SkipUser = $true
                    break
                } elseif ($OUPath -notlike "$OU") {
                    $SkipUser = $true
                    break
                }
            }
            if ($SkipUser) {
                continue
            }

            Write-Verbose -Message "Processing $($Contact.DisplayName) - $($CountContacts)/$($Contacts.Count)"
            # create dummy objects for manager contacts
            $MyUser = [ordered] @{
                UserPrincipalName     = $null
                SamAccountName        = $null
                Domain                = ConvertFrom-DistinguishedName -DistinguishedName $Contact.DistinguishedName -ToDomainCN
                RuleName              = ''
                RuleOptions           = [System.Collections.Generic.List[string]]::new()
                Enabled               = $true
                HasMailbox            = $null
                EmailAddress          = $Contact.Mail
                SystemEmailAddress    = $Contact.Mail
                DateExpiry            = $null
                DaysToExpire          = $null
                PasswordExpired       = $null
                PasswordDays          = $null
                PasswordAtNextLogon   = $null
                PasswordLastSet       = $null
                PasswordNotRequired   = $null
                PasswordNeverExpires  = $null
                LastLogonDate         = $null
                LastLogonDays         = $null
                Manager               = $null
                ManagerDisplayName    = $null
                ManagerSamAccountName = $null
                ManagerEmail          = $null
                ManagerStatus         = $null
                ManagerLastLogonDays  = $null
                ManagerType           = $null
                DisplayName           = $Contact.DisplayName
                Name                  = $Contact.Name
                GivenName             = $null
                Surname               = $null
                OrganizationalUnit    = $OUPath
                MemberOf              = $Contact.MemberOf
                DistinguishedName     = $Contact.DistinguishedName
                ManagerDN             = $null
                Country               = $null
                CountryCode           = $null
                Type                  = 'Contact'
                EmailFrom             = $EmailFrom
            }
            # this allows to extend the object with custom properties requested by user
            # especially custom extensions for use within rules
            foreach ($E in $ExtendedProperties) {
                $MyUser[$E] = $User.$E
            }
            if ($HashtableField -eq 'NetBiosSamAccountName') {
                # Contacts do not have NetBiosSamAccountName
                continue
            } else {
                if ($AsHashTableObject) {
                    $CachedUsers["$($Contact.$HashtableField)"] = $MyUser
                } else {
                    $CachedUsers["$($Contact.$HashtableField)"] = [PSCustomObject] $MyUser
                }
            }
        }
    }
    if ($AsHashTable) {
        $CachedUsers
    } else {
        $CachedUsers.Values
    }
}