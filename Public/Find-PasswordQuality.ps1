function Find-PasswordQuality {
    [CmdletBinding()]
    param(
        [string[]] $WeakPasswords,
        [switch] $IncludeStatistics
    )

    $PropertiesToAdd = @(
        'ClearTextPassword'
        'LMHash'
        'EmptyPassword'
        'WeakPassword'
        'DefaultComputerPassword'
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

    $ModuleExists = Get-Command -Module DSInternals -ErrorAction SilentlyContinue
    if (-not $ModuleExists) {
        Write-Color -Text "[e] ", "DSInternals module is not installed. Please install it using Install-Module DSInternals -Verbose" -Color Yellow, Red
        return
    }
    $AllUsers = Find-Password -AsHashTable -HashtableField NetBiosSamAccountName -ReturnObjectsType Users -AsHashTableObject -AddEmptyProperties $PropertiesToAdd

    Write-Color -Text "[i] ", "Discovering forest information" -Color Yellow, Gray, White, Yellow, White, Yellow, White
    $ForestInformation = Get-WinADForestDetails -PreferWritable -Forest $Forest -ExcludeDomains $ExcludeDomains -IncludeDomains $IncludeDomains -ExtendedForestInformation $ExtendedForestInformation

    $PasswordsInHash = [ordered] @{}
    $PasswordQuality = foreach ($Domain in $ForestInformation.Domains) {
        Write-Color -Text "[i] ", "Discovering DC for domain ", "$($Domain)", " in forest ", $ForestInformation.Name -Color Yellow, Gray, White, Yellow, White, Yellow, White
        $Server = $ForestInformation['QueryServers'][$Domain]['HostName'][0]

        Write-Color -Text "[i] ", "Getting replication data from ", "$($Domain)", " using ", $Server -Color Yellow, Gray, White, Yellow, White, Yellow, White

        $testPasswordQualitySplat = @{
            WeakPasswords = $WeakPasswords
        }
        Remove-EmptyValue -Hashtable $testPasswordQualitySplat

        try {
            Get-ADReplAccount -All -Server $Server -ErrorAction Stop
        } catch {
            Write-Color -Text "[e] ", "Unable to get replication data from ", "$($Domain)", " using ", $Server, ". Error: ", $_.Exception.Message -Color Red, Yellow, White, Yellow, Red, Red
        }
    }
    Write-Color -Text "[i] Testing password quality" -Color Yellow, Gray, White, Yellow, White, Yellow, White
    $Quality = $PasswordQuality | Test-PasswordQuality @testPasswordQualitySplat -IncludeDisabledAccounts

    Write-Color -Text "[i] Processing results" -Color Yellow, Gray, White, Yellow, White, Yellow, White
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
        DefaultComputerPassword                = $PasswordsInHash.DefaultComputerPassword.Count
        DefaultComputerPasswordEnabledOnly     = 0
        DefaultComputerPasswordDisabledOnly    = 0
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

    $OutputUsers = foreach ($User in $AllUsers.Keys) {
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
                } else {
                    $AllUsers[$User][$Property] = $false
                }
            }
        }
        [PSCustomObject] $AllUsers[$User]
    }
    if ($IncludeStatistics) {
        [ordered] @{
            Statistics = $QualityStatistics
            Users      = $OutputUsers
        }
    } else {
        $OutputUsers
    }
}