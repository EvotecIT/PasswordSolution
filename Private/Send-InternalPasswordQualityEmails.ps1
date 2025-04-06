function Send-InternalPasswordQualityEmails {
    [CmdletBinding()]
    param (
        [System.Collections.IDictionary] $Configuration,
        [Array] $Users,
        [System.Collections.IDictionary] $Statistics,
        [Array] $EmailRedirect,
        [int] $EmailLimit
    )

    $EmailParameters = $Configuration['EmailConfiguration']
    $EmailInformation = $Configuration['QualityEmail']

    $OutputData = [ordered] @{
        'Configuration' = $Configuration
        'Users'         = $Users
        'Statistics'    = $Statistics
        'EmailRedirect' = $EmailRedirect
        'EmailSent'     = [System.Collections.Generic.List[PSCustomObject]]::new()
    }

    $Count = 0
    [Array] $UsersProcessed = foreach ($User in $Users) {
        foreach ($Email in $EmailInformation) {
            $Matched = $false
            foreach ($Key in $Email.Keys) {
                if ($Key -eq 'Body', 'DuplicatePasswordGroupsType', 'DuplicatePasswordGroupsCount', 'Operator') {
                    continue
                }
                if ($Email['Operator'] -eq 'OR') {
                    # we treat OR operator to end the loop if any of the conditions is met
                    if ($Key -eq 'DuplicatePasswordGroups') {
                        if ($null -ne $Email['DuplicatePasswordGroupsCount']) {
                            if ($Email['DuplicatePasswordGroupsType'] -eq 'eq') {
                                if ($User.$Key.Count -eq $Email['DuplicatePasswordGroupsCount']) {
                                    $Matched = $true
                                    break
                                }
                            } elseif ($Email['DuplicatePasswordGroupsType'] -eq 'lt') {
                                if ($User.$Key.Count -lt $Email['DuplicatePasswordGroupsCount']) {
                                    $Matched = $true
                                    break
                                }
                            } elseif ($Email['DuplicatePasswordGroupsType'] -eq 'gt') {
                                if ($User.$Key.Count -gt $Email['DuplicatePasswordGroupsCount']) {
                                    $Matched = $true
                                    break
                                }
                            }
                        } else {
                            if ($User.$Key) {
                                $Matched = $true
                                break
                            }
                        }
                    } elseif ($Key -eq 'Domains') {
                        if ($null -ne $Email['Domains']) {
                            if ($User.$Key -in $Email['Domains']) {
                                $Matched = $true
                                break
                            }
                        } else {
                            if ($User.$Key) {
                                $Matched = $true
                                break
                            }
                        }
                    } elseif ($Key -eq 'OrganizationalUnit') {
                        if ($null -ne $Email['OrganizationalUnit']) {
                            if ($User.$Key -in $Email['OrganizationalUnit']) {
                                $Matched = $true
                                break
                            }
                        } else {
                            if ($User.$Key) {
                                $Matched = $true
                                break
                            }
                        }
                    } elseif ($Key -eq 'MemberOf') {
                        if ($null -ne $Email['MemberOf']) {
                            if ($User.$Key -in $Email['MemberOf']) {
                                $Matched = $true
                                break
                            }
                        } else {
                            if ($User.$Key) {
                                $Matched = $true
                                break
                            }
                        }
                    } elseif ($Key -eq 'Country') {
                        if ($null -ne $Email['Country']) {
                            if ($User.$Key -in $Email['Country']) {
                                $Matched = $true
                                break
                            }
                        } else {
                            if ($User.$Key) {
                                $Matched = $true
                                break
                            }
                        }
                    } else {
                        if ($User.$Key -eq $Email[$Key]) {
                            $Matched = $true
                            break
                        }
                    }
                } else {
                    # we treat AND operator to require all conditions to be met
                    if ($Key -eq 'DuplicatePasswordGroups') {
                        if ($null -ne $Email['DuplicatePasswordGroupsCount']) {
                            if ($Email['DuplicatePasswordGroupsType'] -eq 'eq') {
                                if ($User.$Key.Count -ne $Email['DuplicatePasswordGroupsCount']) {
                                    $Matched = $false
                                    break
                                }
                            } elseif ($Email['DuplicatePasswordGroupsType'] -eq 'lt') {
                                if ($User.$Key.Count -ge $Email['DuplicatePasswordGroupsCount']) {
                                    $Matched = $false
                                    break
                                }
                            } elseif ($Email['DuplicatePasswordGroupsType'] -eq 'gt') {
                                if ($User.$Key.Count -le $Email['DuplicatePasswordGroupsCount']) {
                                    $Matched = $false
                                    break
                                }
                            }
                        } else {
                            if (-not $User.$Key) {
                                $Matched = $false
                                break
                            }
                        }

                    } elseif ($Key -eq 'Domains') {
                        if ($null -ne $Email['Domains']) {
                            if ($User.$Key -notin $Email['Domains']) {
                                $Matched = $false
                                break
                            }
                        } else {
                            if (-not $User.$Key) {
                                $Matched = $false
                                break
                            }
                        }
                    } elseif ($Key -eq 'OrganizationalUnit') {
                        if ($null -ne $Email['OrganizationalUnit']) {
                            if ($User.$Key -notin $Email['OrganizationalUnit']) {
                                $Matched = $false
                                break
                            }
                        } else {
                            if (-not $User.$Key) {
                                $Matched = $false
                                break
                            }
                        }
                    } elseif ($Key -eq 'MemberOf') {
                        if ($null -ne $Email['MemberOf']) {
                            if ($User.$Key -notin $Email['MemberOf']) {
                                $Matched = $false
                                break
                            }
                        } else {
                            if (-not $User.$Key) {
                                $Matched = $false
                                break
                            }
                        }
                    } elseif ($Key -eq 'Country') {
                        if ($null -ne $Email['Country']) {
                            if ($User.$Key -notin $Email['Country']) {
                                $Matched = $false
                                break
                            }
                        } else {
                            if (-not $User.$Key) {
                                $Matched = $false
                                break
                            }
                        }
                    } else {
                        if ($User.$Key -ne $Email[$Key]) {
                            $Matched = $false
                            break
                        }
                    }
                }
            }
            if ($Matched) {
                $SourceParameters = @{
                    User       = $User
                    Email      = $Email
                    Statistics = $Statistics
                }
                # Send email to the user, as the conditions are met
                $Body = EmailBody -EmailBody $Email.Body -Parameter $SourceParameters
                $Subject = Add-ParametersToString -String $Email.Subject -Parameter $SourceParameters

                $EmailParametersMissing = @{
                    Body    = $Body
                    Subject = $Subject
                    To      = $User.EmailAddress
                }
                if ($EmailRedirect.Count -gt 0) {
                    $EmailParametersMissing['To'] = $EmailRedirect
                }

                try {
                    $DataSent = Send-EmailMessage @EmailParameters @EmailParametersMissing -ErrorAction Stop -WarningAction SilentlyContinue
                    $OutputData.EmailSent.Add($DataSent)
                } catch {
                    if ($_.Exception.Message -like "*Credential*") {
                        Write-Color -Text "[e] " , "Failed to send email to $($EmailParametersMissing.To) because error: $($_.Exception.Message)" -Color Yellow, White, Red
                        Write-Color -Text "[i] " , "Please make sure you have valid credentials in your configuration file (graph encryption issue?)" -Color Yellow, White, Red
                    } else {
                        Write-Color -Text "[e] " , "Failed to send email to $($EmailParametersMissing.To) because error: $($_.Exception.Message)" -Color Yellow, White, Red
                    }
                }

                [PSCustomObject] @{
                    UserPrincipalName          = $User.UserPrincipalName
                    SamAccountName             = $User.SamAccountName
                    Domain                     = $User.Domain
                    EmailSent                  = $DataSent.Status
                    EmailError                 = $DataSent.Error
                    EmailSentTo                = $DataSent.SentTo
                    EmailSentFrom              = $DataSent.SentFrom
                    Enabled                    = $User.Enabled
                    HasMailbox                 = $User.HasMailbox
                    EmailAddress               = $User.EmailAddress
                    SystemEmailAddress         = $User.SystemEmailAddress
                    DateExpiry                 = $User.DateExpiry
                    DaysToExpire               = $User.DaysToExpire
                    PasswordExpired            = $User.PasswordExpired
                    PasswordDays               = $User.PasswordDays
                    PasswordAtNextLogon        = $User.PasswordAtNextLogon
                    PasswordLastSet            = $User.PasswordLastSet
                    PasswordNotRequired        = $User.PasswordNotRequired
                    PasswordNeverExpires       = $User.PasswordNeverExpires
                    LastLogonDate              = $User.LastLogonDate
                    LastLogonDays              = $User.LastLogonDays
                    ClearTextPassword          = $User.ClearTextPassword
                    LMHash                     = $User.LMHash
                    EmptyPassword              = $User.EmptyPassword
                    WeakPassword               = $User.WeakPassword
                    AESKeysMissing             = $User.AESKeysMissing
                    PreAuthNotRequired         = $User.PreAuthNotRequired
                    DESEncryptionOnly          = $User.DESEncryptionOnly
                    Kerberoastable             = $User.Kerberoastable
                    DelegatableAdmins          = $User.DelegatableAdmins
                    SmartCardUsersWithPassword = $User.SmartCardUsersWithPassword
                    TimeToExecute              = $DataSent.TimeToExecute
                    #DuplicatePasswordGroups    = if ($null -ne $Email['DuplicatePasswordGroups']) { ($Email['DuplicatePasswordGroups'] | ForEach-Object { $_.Name }) } else { @() }
                    #ManagerSamAccountName      = if ($null -ne $Email['ManagerSamAccountName']) { ($Email['ManagerSamAccountName'] | ForEach-Object { $_.SamAccountName }) } else { @() }
                }

                if ($EmailLimit -gt 0) {
                    $Count++
                    if ($Count -ge $EmailLimit) {
                        break
                    }
                }
            }
        }
    }

    $OutputData.UsersProcessed = $UsersProcessed
    $OutputData
    <#
    PS C:\Support\GitHub\PasswordSolution> $users[0]

        UserPrincipalName          :
        SamAccountName             : Guest
        Domain                     : ad.evotec.xyz
        RuleName                   :
        RuleOptions                : {}
        Enabled                    : False
        HasMailbox                 : No
        EmailAddress               :
        SystemEmailAddress         :
        DateExpiry                 :
        DaysToExpire               :
        PasswordExpired            : False
        PasswordDays               :
        PasswordAtNextLogon        : False
        PasswordLastSet            :
        PasswordNotRequired        : True
        PasswordNeverExpires       : True
        LastLogonDate              :
        LastLogonDays              :
        ClearTextPassword          : False
        LMHash                     : False
        EmptyPassword              : True
        WeakPassword               : False
        AESKeysMissing             : False
        PreAuthNotRequired         : False
        DESEncryptionOnly          : False
        Kerberoastable             : False
        DelegatableAdmins          : False
        SmartCardUsersWithPassword : False
        DuplicatePasswordGroups    :
        Manager                    :
        ManagerDisplayName         :
        ManagerSamAccountName      :
        ManagerEmail               :
        ManagerStatus              : Missing
        ManagerLastLogonDays       :
        ManagerType                :
        DisplayName                : Guest
        Name                       : Guest
        GivenName                  :
        Surname                    :
        OrganizationalUnit         : CN=Users,DC=ad,DC=evotec,DC=xyz
        MemberOf                   : {CN=Guests,CN=Builtin,DC=ad,DC=evotec,DC=xyz}
        DistinguishedName          : CN=Guest,CN=Users,DC=ad,DC=evotec,DC=xyz
        ManagerDN                  :
        Country                    : Unknown
        CountryCode                : Unknown
        Type                       : User
        EmailFrom                  : AD
    #>


}