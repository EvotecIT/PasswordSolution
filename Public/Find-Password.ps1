function Find-Password {
    [CmdletBinding()]
    param(
        [alias('ForestName')][string] $Forest,
        [string[]] $ExcludeDomains,
        [alias('Domain', 'Domains')][string[]] $IncludeDomains,
        [System.Collections.IDictionary] $ExtendedForestInformation,
        [string] $OverwriteEmailProperty,
        #[Array] $ConditionProperties,
        #[System.Collections.IDictionary] $CachedUsers,
        #[System.Collections.IDictionary] $CachedUsersPrepared,
        #[System.Collections.IDictionary] $CachedManagers
        [switch] $AsHashTable
    )
    $Today = Get-Date

    $Properties = @(
        'Manager', 'DisplayName', 'GivenName', 'Surname', 'SamAccountName', 'EmailAddress', 'msDS-UserPasswordExpiryTimeComputed', 'PasswordExpired', 'PasswordLastSet', 'PasswordNotRequired', 'Enabled', 'PasswordNeverExpires', 'Mail', 'MemberOf', 'LastLogonDate'
        'userAccountControl'
        if ($OverwriteEmailProperty) {
            $OverwriteEmailProperty
        }
        if ($ConditionProperties) {
            $ConditionProperties
        }
    )
    # We're caching all users to make sure it's speedy gonzales when querying for Managers
    if (-not $CachedUsers) {
        $CachedUsers = [ordered] @{ }
    }
    if (-not $Cache) {
        $Cache = [ordered] @{ }
    }
    Write-Color -Text "[i] Discovering forest information" -Color White, Yellow, White, Yellow, White, Yellow, White
    $ForestInformation = Get-WinADForestDetails -Forest $Forest -ExcludeDomains $ExcludeDomains -IncludeDomains $IncludeDomains -ExtendedForestInformation $ExtendedForestInformation

    [Array] $Users = foreach ($Domain in $ForestInformation.Domains) {
        Write-Color -Text "[i] Discovering DC for domain ", "$($Domain)", " in forest ", $ForestInformation.Name -Color White, Yellow, White, Yellow, White, Yellow, White
        $Server = $ForestInformation['QueryServers'][$Domain]['HostName'][0]

        Write-Color -Text "[i] Getting users from ", "$($Domain)", " using ", $Server -Color White, Yellow, White, Yellow, White, Yellow, White
        # We query all users instead of using filter. Since we need manager field and manager data this way it should be faster (query once - get it all)
        #$DomainUsers = Get-ADUser -Server $Server -Filter '*' -Properties $Properties -ErrorAction Stop
        #foreach ($_ in $DomainUsers) {
        #$CachedUsers["$($_.DistinguishedName)"] = $_
        # We reuse filtering, account is enabled, password is required and password is not set to change on next logon
        #if ($_.Enabled -eq $true -and $_.PasswordNotRequired -ne $true -and $null -ne $_.PasswordLastSet) {
        #    $_
        #}
        #}
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
    Write-Color -Text "[i] Preparing all users for password expirations in forest ", $Forest.Name -Color White, Yellow, White, Yellow, White, Yellow, White
    $ProcessedUsers = foreach ($User in $Users) {
        #$UserManager = $Cache["$($User.Manager)"]
        if ($User.Manager) {
            $Manager = $Cache[$User.Manager].DisplayName
            $ManagerSamAccountName = $Cache[$User.Manager].SamAccountName
            $ManagerEmail = $Cache[$User.Manager].Mail
            $ManagerEnabled = $Cache[$User.Manager].Enabled
            $ManagerLastLogon = $Cache[$User.Manager].LastLogonDate
            if ($ManagerLastLogon) {
                $ManagerLastLogonDays = $( - $($ManagerLastLogon - $Today).Days)
            } else {
                $ManagerLastLogonDays = $null
            }
            $ManagerStatus = if ($ManagerEnabled) { 'Enabled' } else { 'Disabled' }
        } else {
            if ($User.ObjectClass -eq 'user') {
                $ManagerStatus = 'Missing'
            } else {
                $ManagerStatus = 'Not available'
            }
            $Manager = $null
            $ManagerSamAccountName = $null
            $ManagerEmail = $null
            $ManagerEnabled = $null
            $ManagerLastLogon = $null
            $ManagerLastLogonDays = $null
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
        if ($PasswordNeverExpires -or $null -eq $User.PasswordLastSet) {
            $DateExpiry = $null
            $DaysToExpire = $null
        }

        $UserAccountControl = Convert-UserAccountControl -UserAccountControl $User.UserAccountControl
        if ($UserAccountControl -contains 'INTERDOMAIN_TRUST_ACCOUNT') {
            continue
        }

        $MyUser = [ordered] @{
            UserPrincipalName     = $User.UserPrincipalName
            SamAccountName        = $User.SamAccountName
            Domain                = ConvertFrom-DistinguishedName -DistinguishedName $User.DistinguishedName -ToDomainCN
            Enabled               = $User.Enabled
            EmailAddress          = $EmailAddress
            DateExpiry            = $DateExpiry
            DaysToExpire          = $DaysToExpire
            PasswordExpired       = $User.PasswordExpired
            PasswordDays          = $PasswordDays
            PasswordLastSet       = $User.PasswordLastSet
            PasswordNotRequired   = $User.PasswordNotRequired
            PasswordNeverExpires  = $PasswordNeverExpires
            #PasswordAtNextLogon  = $null -eq $User.PasswordLastSet
            Manager               = $Manager
            ManagerSamAccountName = $ManagerSamAccountName
            ManagerEmail          = $ManagerEmail
            ManagerStatus         = $ManagerStatus
            ManagerLastLogonDays  = $ManagerLastLogonDays
            DisplayName           = $User.DisplayName
            GivenName             = $User.GivenName
            Surname               = $User.Surname
            OrganizationalUnit    = ConvertFrom-DistinguishedName -DistinguishedName $User.DistinguishedName -ToOrganizationalUnit
            MemberOf              = $User.MemberOf
            DistinguishedName     = $User.DistinguishedName
            ManagerDN             = $User.Manager
        }
        foreach ($Property in $ConditionProperties) {
            $MyUser["$Property"] = $User.$Property
        }
        #[PSCustomObject] $MyUser

        $CachedUsers["$($User.DistinguishedName)"] = [PSCustomObject] $MyUser
    }
    if ($AsHashTable) {
        $CachedUsers
    } else {
        $CachedUsers.Values
    }
}