function Find-Password {
    [CmdletBinding()]
    param(
        [alias('ForestName')][string] $Forest,
        [string[]] $ExcludeDomains,
        [alias('Domain', 'Domains')][string[]] $IncludeDomains,
        [System.Collections.IDictionary] $ExtendedForestInformation,
        [string] $AdditionalProperties,
        [Array] $ConditionProperties,
        [System.Collections.IDictionary] $WriteParameters,
        [System.Collections.IDictionary] $CachedUsers,
        [System.Collections.IDictionary] $CachedUsersPrepared,
        [System.Collections.IDictionary] $CachedManagers
    )
    if ($null -eq $WriteParameters) {
        $WriteParameters = @{
            ShowTime   = $true
            LogFile    = ""
            TimeFormat = "yyyy-MM-dd HH:mm:ss"
        }
    }


    $Properties = @(
        'Manager', 'DisplayName', 'GivenName', 'Surname', 'SamAccountName', 'EmailAddress', 'msDS-UserPasswordExpiryTimeComputed', 'PasswordExpired', 'PasswordLastSet', 'PasswordNotRequired', 'Enabled', 'PasswordNeverExpires', 'Mail', 'MemberOf'
        if ($AdditionalProperties) {
            $AdditionalProperties
        }
        if ($ConditionProperties) {
            $ConditionProperties
        }
    )
    # We're caching all users to make sure it's speedy gonzales when querying for Managers
    if (-not $CachedUsers) {
        $CachedUsers = [ordered] @{ }
    }
    if (-not $CachedUsersPrepared) {
        $CachedUsersPrepared = [ordered] @{ }
    }
    if (-not $CachedManagers) {
        $CachedManagers = [ordered] @{}
    }
    Write-Color @WriteParameters -Text "[i] Discovering forest information" -Color White, Yellow, White, Yellow, White, Yellow, White
    $ForestInformation = Get-WinADForestDetails -Forest $Forest -ExcludeDomains $ExcludeDomains -IncludeDomains $IncludeDomains -ExtendedForestInformation $ExtendedForestInformation

    [Array] $Users = foreach ($Domain in $ForestInformation.Domains) {
        Write-Color @WriteParameters -Text "[i] Discovering DC for domain ", "$($Domain)", " in forest ", $ForestInformation.Name -Color White, Yellow, White, Yellow, White, Yellow, White
        $Server = $ForestInformation['QueryServers'][$Domain]['HostName'][0]

        Write-Color @WriteParameters -Text "[i] Getting users from ", "$($Domain)", " using ", $Server -Color White, Yellow, White, Yellow, White, Yellow, White
        # We query all users instead of using filter. Since we need manager field and manager data this way it should be faster (query once - get it all)
        $DomainUsers = Get-ADUser -Server $Server -Filter '*' -Properties $Properties -ErrorAction Stop
        foreach ($_ in $DomainUsers) {
            Add-Member -InputObject $_ -Value $Domain -Name 'Domain' -Force -Type NoteProperty
            $CachedUsers["$($_.DistinguishedName)"] = $_
            # We reuse filtering, account is enabled, password is required and password is not set to change on next logon
            if ($_.Enabled -eq $true -and $_.PasswordNotRequired -ne $true -and $null -ne $_.PasswordLastSet) {
                #if ($_.Enabled -eq $true -and $_.PasswordNeverExpires -eq $false -and $null -ne $_.PasswordLastSet -and $_.PasswordNotRequired -ne $true) {
                $_
            }
        }
        try {

        } catch {
            $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
            Write-Color @WriteParameters '[e] Error: ', $ErrorMessage -Color White, Red
        }
    }
    Write-Color @WriteParameters -Text "[i] Preparing all users for password expirations in forest ", $Forest.Name -Color White, Yellow, White, Yellow, White, Yellow, White
    $ProcessedUsers = foreach ($User in $Users) {
        $UserManager = $CachedUsers["$($User.Manager)"]
        if ($User.Manager) {
            $Manager = $CachedUsers[$User.Manager].DisplayName
            $ManagerSamAccountName = $CachedUsers[$User.Manager].SamAccountName
            $ManagerEmail = $CachedUsers[$User.Manager].Mail
            $ManagerEnabled = $CachedUsers[$User.Manager].Enabled
            $ManagerLastLogon = $CachedUsers[$User.Manager].LastLogonDate
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


        if ($AdditionalProperties) {
            # fix this for a user
            $EmailTemp = $User.$AdditionalProperties
            if ($EmailTemp -like '*@*') {
                $EmailAddress = $EmailTemp
            } else {
                $EmailAddress = $User.EmailAddress
            }
            # Fix this for manager as well
            if ($UserManager) {
                if ($UserManager.$AdditionalProperties -like '*@*') {
                    $UserManager.Mail = $UserManager.$AdditionalProperties
                }
            }
        } else {
            $EmailAddress = $User.EmailAddress
        }

        if ($User."msDS-UserPasswordExpiryTimeComputed" -ne 9223372036854775807) {
            # This is standard situation where users password is expiring as needed
            try {
                $DateExpiry = ([datetime]::FromFileTime($User."msDS-UserPasswordExpiryTimeComputed"))
            } catch {
                $DateExpiry = $User."msDS-UserPasswordExpiryTimeComputed"
            }
            try {
                $DaysToExpire = (New-TimeSpan -Start (Get-Date) -End ([datetime]::FromFileTime($User."msDS-UserPasswordExpiryTimeComputed"))).Days
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

        $MyUser = [ordered] @{
            UserPrincipalName     = $User.UserPrincipalName
            SamAccountName        = $User.SamAccountName
            Domain                = $User.Domain
            EmailAddress          = $EmailAddress
            DateExpiry            = $DateExpiry
            DaysToExpire          = $DaysToExpire
            PasswordExpired       = $User.PasswordExpired
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
        }
        foreach ($Property in $ConditionProperties) {
            $MyUser["$Property"] = $User.$Property
        }
        [PSCustomObject] $MyUser
        $CachedUsersPrepared["$($User.DistinguishedName)"] = $MyUser
    }
    foreach ($User in $CachedUsersPrepared.Keys) {
        $ManagerDN = $CachedUsersPrepared[$User]['ManagerDN']
        if ($ManagerDN) {
            $Manager = $CachedUsers[$ManagerDN]

            $MyUser = [ordered] @{
                UserPrincipalName = $Manager.UserPrincipalName
                Domain            = $Manager.Domain
                SamAccountName    = $Manager.SamAccountName
                DisplayName       = $Manager.DisplayName
                GivenName         = $Manager.GivenName
                Surname           = $Manager.Surname
                DistinguishedName = $ManagerDN
            }
            foreach ($Property in $ConditionProperties) {
                $MyUser["$Property"] = $User.$Property
            }
            $CachedManagers[$ManagerDN] = $MyUser
        }
    }


    $ProcessedUsers
}

#$Test = Find-PasswordExpiryCheck -AdditionalProperties 'extensionAttribute13'
#$Test | Format-Table -AutoSize *
