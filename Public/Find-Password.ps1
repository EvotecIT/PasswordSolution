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
        [Parameter(DontShow)][string[]] $AddEmptyProperties = @()
    )
    $Today = Get-Date

    $Properties = @(
        'Manager', 'DisplayName', 'GivenName', 'Surname', 'SamAccountName', 'EmailAddress', 'msDS-UserPasswordExpiryTimeComputed', 'PasswordExpired', 'PasswordLastSet', 'PasswordNotRequired', 'Enabled', 'PasswordNeverExpires', 'Mail', 'MemberOf', 'LastLogonDate', 'Name'
        'userAccountControl'
        'msExchMailboxGuid'
        'pwdLastSet', 'ObjectClass'
        if ($OverwriteEmailProperty) {
            $OverwriteEmailProperty
        }
    )
    $PropertiesContacts = @(
        'SamAccountName', 'CanonicalName', 'WhenChanged', 'WhenChanged', 'DisplayName', 'DistinguishedName', 'Name', 'Mail', 'TargetAddress', 'ObjectClass'
    )

    # We're caching all users to make sure it's speedy gonzales when querying for Managers
    if (-not $CachedUsers) {
        $CachedUsers = [ordered] @{ }
    }
    if (-not $Cache) {
        $Cache = [ordered] @{ }
    }
    Write-Color -Text "[i] Discovering forest information" -Color White, Yellow, White, Yellow, White, Yellow, White
    $ForestInformation = Get-WinADForestDetails -Extended -Forest $Forest -ExcludeDomains $ExcludeDomains -IncludeDomains $IncludeDomains -ExtendedForestInformation $ExtendedForestInformation

    # lets get domain name / netbios hashtable for easy use
    $DNSNetBios = @{ }
    foreach ($NETBIOS in $ForestInformation.DomainsExtendedNetBIOS.Keys) {
        $DNSNetBios[$ForestInformation.DomainsExtendedNetBIOS[$NETBIOS].DnsRoot] = $NETBIOS
    }

    [Array] $Users = foreach ($Domain in $ForestInformation.Domains) {
        Write-Color -Text "[i] Discovering DC for domain ", "$($Domain)", " in forest ", $ForestInformation.Name -Color White, Yellow, White, Yellow, White, Yellow, White
        $Server = $ForestInformation['QueryServers'][$Domain]['HostName'][0]

        Write-Color -Text "[i] Getting users from ", "$($Domain)", " using ", $Server -Color White, Yellow, White, Yellow, White, Yellow, White
        try {
            Get-ADUser -Server $Server -Filter '*' -Properties $Properties -ErrorAction Stop
        } catch {
            $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
            Write-Color '[e] Error: ', $ErrorMessage -Color White, Red
        }
    }
    foreach ($User in $Users) {
        $Cache[$User.DistinguishedName] = $User
    }

    if ($ReturnObjectsType -contains 'Contacts') {
        [Array] $Contacts = foreach ($Domain in $ForestInformation.Domains) {
            Write-Color -Text "[i] Discovering DC for domain ", "$($Domain)", " in forest ", $ForestInformation.Name -Color White, Yellow, White, Yellow, White, Yellow, White
            $Server = $ForestInformation['QueryServers'][$Domain]['HostName'][0]

            Write-Color -Text "[i] Getting contacts from ", "$($Domain)", " using ", $Server -Color White, Yellow, White, Yellow, White, Yellow, White
            try {
                Get-ADObject -LDAPFilter "objectClass=Contact" -Server $Server -Properties $PropertiesContacts -ErrorAction Stop
            } catch {
                $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
                Write-Color '[e] Error: ', $ErrorMessage -Color White, Red
            }
        }
        foreach ($Contact in $Contacts) {
            $Cache[$Contact.DistinguishedName] = $Contact
        }
    }

    Write-Color -Text "[i] Preparing all users for password expirations in forest ", $Forest.Name -Color White, Yellow, White, Yellow, White, Yellow, White
    $CountUsers = 0
    foreach ($User in $Users) {
        $CountUsers++
        Write-Verbose -Message "Processing $($User.DisplayName) - $($CountUsers)/$($Users.Count)"
        $DateExpiry = $null
        $DaysToExpire = $null
        $PasswordDays = $null
        $PasswordNeverExpires = $null
        $PasswordAtNextLogon = $null
        $HasMailbox = $null
        #$UserManager = $Cache["$($User.Manager)"]
        if ($User.Manager) {
            $Manager = $Cache[$User.Manager].DisplayName
            $ManagerSamAccountName = $Cache[$User.Manager].SamAccountName
            $ManagerDisplayName = $Cache[$User.Manager].DisplayName
            $ManagerEmail = $Cache[$User.Manager].Mail
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
            $Manager = $null
            $ManagerSamAccountName = $null
            $ManagerDisplayName = $null
            $ManagerEmail = $null
            $ManagerEnabled = $null
            $ManagerLastLogon = $null
            $ManagerLastLogonDays = $null
            $ManagerType = $null
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
        if ($User.'msExchMailboxGuid') {
            $HasMailbox = $true
        } else {
            $HasMailbox = $false
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
                DateExpiry           = $DateExpiry
                DaysToExpire         = $DaysToExpire
                PasswordExpired      = $User.PasswordExpired
                PasswordDays         = $PasswordDays
                PasswordAtNextLogon  = $PasswordAtNextLogon
                PasswordLastSet      = $User.PasswordLastSet
                PasswordNotRequired  = $User.PasswordNotRequired
                PasswordNeverExpires = $PasswordNeverExpires
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
                OrganizationalUnit    = ConvertFrom-DistinguishedName -DistinguishedName $User.DistinguishedName -ToOrganizationalUnit
                MemberOf              = $User.MemberOf
                DistinguishedName     = $User.DistinguishedName
                ManagerDN             = $User.Manager
                Type                  = 'User'
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
                DateExpiry            = $DateExpiry
                DaysToExpire          = $DaysToExpire
                PasswordExpired       = $User.PasswordExpired
                PasswordDays          = $PasswordDays
                PasswordAtNextLogon   = $PasswordAtNextLogon
                PasswordLastSet       = $User.PasswordLastSet
                PasswordNotRequired   = $User.PasswordNotRequired
                PasswordNeverExpires  = $PasswordNeverExpires
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
                Type                  = 'User'
            }
        }
        foreach ($Property in $ConditionProperties) {
            $MyUser["$Property"] = $User.$Property
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
                DateExpiry            = $null
                DaysToExpire          = $null
                PasswordExpired       = $null
                PasswordDays          = $null
                PasswordAtNextLogon   = $null
                PasswordLastSet       = $null
                PasswordNotRequired   = $null
                PasswordNeverExpires  = $null
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
                OrganizationalUnit    = ConvertFrom-DistinguishedName -DistinguishedName $Contact.DistinguishedName -ToOrganizationalUnit
                MemberOf              = $Contact.MemberOf
                DistinguishedName     = $Contact.DistinguishedName
                ManagerDN             = $null
                Type                  = 'Contact'
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