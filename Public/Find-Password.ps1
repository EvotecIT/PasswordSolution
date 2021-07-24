function Find-Password {
    [CmdletBinding()]
    param(
        [alias('ForestName')][string] $Forest,
        [string[]] $ExcludeDomains,
        [alias('Domain', 'Domains')][string[]] $IncludeDomains,
        [System.Collections.IDictionary] $ExtendedForestInformation,
        [string] $OverwriteEmailProperty,
        [switch] $AsHashTable
    )
    $Today = Get-Date

    $Properties = @(
        'Manager', 'DisplayName', 'GivenName', 'Surname', 'SamAccountName', 'EmailAddress', 'msDS-UserPasswordExpiryTimeComputed', 'PasswordExpired', 'PasswordLastSet', 'PasswordNotRequired', 'Enabled', 'PasswordNeverExpires', 'Mail', 'MemberOf', 'LastLogonDate', 'Name'
        'userAccountControl'
        'msExchMailboxGuid'
        'pwdLastSet'
        if ($OverwriteEmailProperty) {
            $OverwriteEmailProperty
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
    foreach ($User in $Users) {
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
            $ManagerEmail = $Cache[$User.Manager].Mail
            $ManagerEnabled = $Cache[$User.Manager].Enabled
            $ManagerLastLogon = $Cache[$User.Manager].LastLogonDate
            if ($ManagerLastLogon) {
                $ManagerLastLogonDays = $( - $($ManagerLastLogon - $Today).Days)
            } else {
                $ManagerLastLogonDays = $null
            }
            if ($ManagerEnabled -and $ManagerEmail) {
                if ((Test-EmailAddress -EmailAddress $ManagerEmail).IsValid -eq $true) {
                    $ManagerStatus = 'Enabled'
                } else {
                    $ManagerStatus = 'Enabled, bad email'
                }
            } elseif ($ManagerEnabled) {
                $ManagerStatus = 'No email'
            } else {
                $ManagerStatus = 'Disabled'
            }
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
        } else {
            $PasswordDays = $null
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

        $MyUser = [ordered] @{
            UserPrincipalName     = $User.UserPrincipalName
            SamAccountName        = $User.SamAccountName
            Domain                = ConvertFrom-DistinguishedName -DistinguishedName $User.DistinguishedName -ToDomainCN
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
            ManagerSamAccountName = $ManagerSamAccountName
            ManagerEmail          = $ManagerEmail
            ManagerStatus         = $ManagerStatus
            ManagerLastLogonDays  = $ManagerLastLogonDays
            DisplayName           = $User.DisplayName
            Name                  = $User.Name
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
        $CachedUsers["$($User.DistinguishedName)"] = [PSCustomObject] $MyUser
    }
    if ($AsHashTable) {
        $CachedUsers
    } else {
        $CachedUsers.Values
    }
}