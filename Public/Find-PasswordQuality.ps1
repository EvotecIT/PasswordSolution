function Find-PasswordQuality {
    <#
    .SYNOPSIS
    Scan Active Directory forest for asses password quality of users

    .DESCRIPTION
    Scan Active Directory forest for asses password quality of users including weak passwords, duplicate groups and more.

    .PARAMETER WeakPasswords
    List of weak passwords to check against

    .PARAMETER WeakPasswordsFilePath
    Path to a file that contains weak passwords, one password per line.

    .PARAMETER WeakPasswordsHashesFile
    Path to a file that contains NT hashes of weak passwords, one hash in HEX format per line. For performance reasons, the -WeakPasswordHashesSortedFile parameter should be used instead.

    .PARAMETER WeakPasswordsHashesSortedFile
    Path to a file that contains NT hashes of weak passwords, one hash in HEX format per line. The hashes must be sorted alphabetically, because a binary search is performed. This parameter is typically used with a list of leaked password hashes from HaveIBeenPwned.

    .PARAMETER IncludeStatistics
    Include statistics in output

    .PARAMETER Forest
    Target different Forest, by default current forest is used

    .PARAMETER ExcludeDomains
    Exclude domain from search, by default whole forest is scanned

    .PARAMETER IncludeDomains
    Include only specific domains, by default whole forest is scanned

    .PARAMETER ExtendedForestInformation
    Ability to provide Forest Information from another command to speed up processing

    .EXAMPLE
    Find-PasswordQuality -WeakPasswords "Test1", "Test2", "Test3"

    .EXAMPLE
    Find-PasswordQuality -WeakPasswords "Test1", "Test2", "Test3" -IncludeStatistics

    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [string[]] $WeakPasswords,
        [string] $WeakPasswordsFilePath,
        [string] $WeakPasswordsHashesFile,
        [string] $WeakPasswordsHashesSortedFile,
        [switch] $IncludeStatistics,

        [alias('ForestName')][string] $Forest,
        [string[]] $ExcludeDomains,
        [alias('Domain', 'Domains')][string[]] $IncludeDomains,
        [System.Collections.IDictionary] $ExtendedForestInformation
    )

    $PropertiesToAdd = @(
        'ClearTextPassword'
        'LMHash'
        'EmptyPassword'
        'WeakPassword'
        #'DefaultComputerPassword'
        #'PasswordNotRequired'
        #'PasswordNeverExpires'
        'AESKeysMissing'
        'PreAuthNotRequired'
        'DESEncryptionOnly'
        'Kerberoastable'
        'DelegatableAdmins'
        'SmartCardUsersWithPassword'
        'DuplicatePasswordGroups'
    )
    if ($WeakPasswordsHashesFile) {
        if (Test-Path -LiteralPath $WeakPasswordsHashesFile) {
            Write-Color -Text "[i] ", "Weak password hashes available to read  from ", $WeakPasswordsHashesFile -Color Yellow, Gray, White, Yellow, White, Yellow, White
            $WeakPasswordHashesStats = Get-FileInformation -File $WeakPasswordsHashesFile
        } else {
            Write-Color -Text "[e] ", "Weak password hashes file not found at ", $WeakPasswordsHashesFile -Color Red, Yellow, White, Yellow, Red
            return
        }
    }
    if ($WeakPasswordsHashesSortedFile) {
        if (Test-Path -LiteralPath $WeakPasswordsHashesSortedFile) {
            Write-Color -Text "[i] ", "Weak passwords hashes (sorted) available to read from ", $WeakPasswordsHashesSortedFile -Color Yellow, Gray, White, Yellow, White, Yellow, White
            $WeakPasswordHashesSortedStats = Get-FileInformation -File $WeakPasswordsHashesSortedFile
        } else {
            Write-Color -Text "[e] ", "Weak passwords hashes (sorted) file not found at ", $WeakPasswordsHashesSortedFile -Color Red, Yellow, White, Yellow, Red
            return
        }
    }
    if ($WeakPasswordsFilePath) {
        if (Test-Path -LiteralPath $WeakPasswordsFilePath) {
            Write-Color -Text "[i] ", "Weak passwords available to read from ", $WeakPasswordsFilePath -Color Yellow, Gray, White, Yellow, White, Yellow, White
            $WeakPasswordsStats = Get-FileInformation -File $WeakPasswordsFilePath
        } else {
            Write-Color -Text "[e] ", "Weak passwords file not found at ", $WeakPasswordsFilePath -Color Red, Yellow, White, Yellow, Red
            return
        }
    }
    $ModuleExists = Get-Command -Module DSInternals -ErrorAction SilentlyContinue
    if (-not $ModuleExists) {
        Write-Color -Text "[e] ", "DSInternals module is not installed. Please install it using Install-Module DSInternals -Verbose" -Color Yellow, Red
        return
    }
    $AllUsers = Find-Password -AsHashTable -HashtableField NetBiosSamAccountName -ReturnObjectsType Users -AsHashTableObject -AddEmptyProperties $PropertiesToAdd -Forest $Forest -IncludeDomains $IncludeDomains -ExcludeDomains $ExcludeDomains

    Write-Color -Text "[i] ", "Discovering forest information" -Color Yellow, Gray, White, Yellow, White, Yellow, White
    $ForestInformation = Get-WinADForestDetails -PreferWritable -Forest $Forest -ExcludeDomains $ExcludeDomains -IncludeDomains $IncludeDomains -ExtendedForestInformation $ExtendedForestInformation

    $PasswordsInHash = [ordered] @{}
    $PasswordQuality = foreach ($Domain in $ForestInformation.Domains) {
        Write-Color -Text "[i] ", "Discovering DC for domain ", "$($Domain)", " in forest ", $ForestInformation.Name -Color Yellow, Gray, White, Yellow, White, Yellow, White
        $Server = $ForestInformation['QueryServers'][$Domain]['HostName'][0]

        Write-Color -Text "[i] ", "Getting replication data from ", "$($Domain)", " using ", $Server -Color Yellow, Gray, White, Yellow, White, Yellow, White

        $testPasswordQualitySplat = @{
            WeakPasswords                = $WeakPasswords
            WeakPasswordsFile            = $WeakPasswordsFilePath
            WeakPasswordHashesFile       = $WeakPasswordsHashesFile
            WeakPasswordHashesSortedFile = $WeakPasswordsHashesSortedFile
            IncludeDisabledAccounts      = $true
        }
        Remove-EmptyValue -Hashtable $testPasswordQualitySplat

        try {
            Get-ADReplAccount -All -Server $Server -ErrorAction Stop
        } catch {
            Write-Color -Text "[e] ", "Unable to get replication data from ", "$($Domain)", " using ", $Server, ". Error: ", $_.Exception.Message -Color Red, Yellow, White, Yellow, Red, Red
        }
    }
    Write-Color -Text "[i] Testing password quality" -Color Yellow, Gray, White, Yellow, White, Yellow, White
    $Quality = $PasswordQuality | Test-PasswordQuality @testPasswordQualitySplat

    Write-Color -Text "[i] Processing results, merging data from DSInternals" -Color Yellow, Gray, White, Yellow, White, Yellow, White
    foreach ($Property in $Quality.PSObject.Properties.Name) {
        $PasswordsInHash[$Property] = $Quality.$Property
    }

    $PasswordGroupsUsers = [ordered] @{}
    $Count = 0
    foreach ($Group in $PasswordsInHash.DuplicatePasswordGroups) {
        $Count++
        foreach ($User in $Group) {
            $PasswordGroupsUsers[$User] = "Group $Count"
        }
    }

    $QualityStatistics = [ordered] @{
        AESKeysMissing                         = $PasswordsInHash.AESKeysMissing.Count
        AESKeysMissingEnabledOnly              = 0
        AESKeysMissingDisabledOnly             = 0
        DESEncryptionOnly                      = $PasswordsInHash.DESEncryptionOnly.Count
        DESEncryptionOnlyEnabledOnly           = 0
        DESEncryptionOnlyDisabledOnly          = 0
        DelegatableAdmins                      = $PasswordsInHash.DelegatableAdmins.Count
        DelegatableAdminsEnabledOnly           = 0
        DelegatableAdminsDisabledOnly          = 0
        DuplicatePasswordGroups                = $PasswordsInHash.DuplicatePasswordGroups.Count
        DuplicatePasswordUsers                 = $PasswordGroupsUsers.Keys.Count
        DuplicatePasswordUsersEnabledOnly      = 0
        DuplicatePasswordUsersDisabledOnly     = 0
        ClearTextPassword                      = $PasswordsInHash.ClearTextPassword.Count
        ClearTextPasswordEnabledOnly           = 0
        ClearTextPasswordDisabledOnly          = 0
        LMHash                                 = $PasswordsInHash.LMHash.Count
        LMHashEnabledOnly                      = 0
        LMHashDisabledOnly                     = 0
        EmptyPassword                          = $PasswordsInHash.EmptyPassword.Count
        EmptyPasswordEnabledOnly               = 0
        EmptyPasswordDisabledOnly              = 0
        WeakPassword                           = $PasswordsInHash.WeakPassword.Count
        WeakPasswordEnabledOnly                = 0
        WeakPasswordDisabledOnly               = 0
        #DefaultComputerPassword                = $PasswordsInHash.DefaultComputerPassword.Count
        #DefaultComputerPasswordEnabledOnly     = 0
        #DefaultComputerPasswordDisabledOnly    = 0
        PasswordNotRequired                    = 0 # $PasswordsInHash.PasswordNotRequired.Count
        PasswordNotRequiredEnabledOnly         = 0
        PasswordNotRequiredDisabledOnly        = 0
        PasswordNeverExpires                   = 0 #$PasswordsInHash.PasswordNeverExpires.Count
        PasswordNeverExpiresEnabledOnly        = 0
        PasswordNeverExpiresDisabledOnly       = 0
        PreAuthNotRequired                     = $PasswordsInHash.PreAuthNotRequired.Count
        PreAuthNotRequiredEnabledOnly          = 0
        PreAuthNotRequiredDisabledOnly         = 0
        Kerberoastable                         = $PasswordsInHash.Kerberoastable.Count
        KerberoastableEnabledOnly              = 0
        KerberoastableDisabledOnly             = 0
        SmartCardUsersWithPassword             = $PasswordsInHash.SmartCardUsersWithPassword.Count
        SmartCardUsersWithPasswordEnabledOnly  = 0
        SmartCardUsersWithPasswordDisabledOnly = 0
    }
    $CountryStatistics = [ordered] @{
        DuplicatePasswordUsers = [ordered] @{}
        WeakPassword           = [ordered] @{}
    }
    $ContinentStatistics = [ordered] @{
        DuplicatePasswordUsers = [ordered] @{}
        WeakPassword           = [ordered] @{}
    }
    $CountryCodeStatistics = [ordered] @{
        DuplicatePasswordUsers = [ordered] @{}
        WeakPassword           = [ordered] @{}
    }
    $CountryToContinent = Convert-CountryToContinent

    $OutputUsers = foreach ($User in $AllUsers.Keys) {
        if ($AllUsers[$User].Country) {
            $Continent = $CountryToContinent[$AllUsers[$User].Country]
            if (-not $Continent) {
                $Continent = 'Unknown'
            }
        } else {
            $Continent = 'Unknown'
        }
        if ($AllUsers[$User].PasswordNotRequired) {
            $QualityStatistics.PasswordNotRequired++
            if ($AllUsers[$User].Enabled -eq $true) {
                $QualityStatistics.PasswordNotRequiredEnabledOnly++
            } else {
                $QualityStatistics.PasswordNotRequiredDisabledOnly++
            }
        }
        if ($AllUsers[$User].PasswordNeverExpires) {
            $QualityStatistics.PasswordNeverExpires++
            if ($AllUsers[$User].Enabled -eq $true) {
                $QualityStatistics.PasswordNeverExpiresEnabledOnly++
            } else {
                $QualityStatistics.PasswordNeverExpiresDisabledOnly++
            }
        }
        foreach ($Property in $PasswordsInHash.Keys) {
            if ($Property -eq 'DuplicatePasswordGroups') {
                if ($PasswordGroupsUsers[$User]) {
                    $AllUsers[$User][$Property] = $PasswordGroupsUsers[$User]
                    if ($AllUsers[$User].Enabled -eq $true) {
                        $QualityStatistics["$($Property)EnabledOnly"]++
                        $QualityStatistics.DuplicatePasswordUsersEnabledOnly++
                    } else {
                        $QualityStatistics["$($Property)DisabledOnly"]++
                        $QualityStatistics.DuplicatePasswordUsersDisabledOnly++
                    }
                    # we keep stats per country for weak passwords and duplicate passwords
                    $CountryStatistics['DuplicatePasswordUsers'][$AllUsers[$User].Country]++
                    $ContinentStatistics['DuplicatePasswordUsers'][$Continent]++
                    $CountryCodeStatistics['DuplicatePasswordUsers'][$AllUsers[$User].CountryCode]++

                } else {
                    $AllUsers[$User][$Property] = ''
                }
            } elseif ($Property -in $PropertiesToAdd) {
                if ($PasswordsInHash[$Property] -contains $User) {
                    $AllUsers[$User][$Property] = $true
                    if ($AllUsers[$User].Enabled -eq $true) {
                        $QualityStatistics["$($Property)EnabledOnly"]++
                    } else {
                        $QualityStatistics["$($Property)DisabledOnly"]++
                    }
                    # we keep stats per country for weak passwords and duplicate passwords
                    if ($Property -eq 'WeakPassword') {
                        $CountryStatistics[$Property][$AllUsers[$User].Country]++
                        $ContinentStatistics[$Property][$Continent]++
                        $CountryCodeStatistics[$Property][$AllUsers[$User].CountryCode]++
                    }
                } else {
                    $AllUsers[$User][$Property] = $false
                }
            }
        }
        [PSCustomObject] $AllUsers[$User]
    }
    if ($IncludeStatistics) {
        [ordered] @{
            Forest                       = $ForestInformation.Forest
            Domains                      = $ForestInformation.Domains
            Statistics                   = $QualityStatistics
            StatisticsCountry            = $CountryStatistics
            StatisticsCountryCode        = $CountryCodeStatistics
            StatisticsContinents         = $ContinentStatistics
            Users                        = $OutputUsers
            WeakPasswordsFileInformation = [ordered] @{
                WeakPasswordHashesStats       = $WeakPasswordHashesStats
                WeakPasswordHashesSortedStats = $WeakPasswordHashesSortedStats
                WeakPasswordsStats            = $WeakPasswordsStats
            }
        }
    } else {
        $OutputUsers
    }
}