function Start-PasswordSolution {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $EmailParameters,
        [System.Collections.IDictionary] $ConfigurationParameters,
        [string] $OverwriteEmailProperty,
        [System.Collections.IDictionary] $UserSection,
        [System.Collections.IDictionary] $ManagerSection,
        [System.Collections.IDictionary] $AdminSection,
        [Array] $Rules,
        [scriptblock] $TemplatePreExpiry,
        [string] $TemplatePreExpirySubject,
        [scriptblock] $TemplatePostExpiry,
        [string] $TemplatePostExpirySubject,
        [scriptblock] $TemplateManager,
        [string] $TemplateManagerSubject,
        [scriptblock] $TemplateManagerNotCompliant,
        [string] $TemplateManagerNotCompliantSubject,
        [System.Collections.IDictionary] $DisplayConsole,
        [System.Collections.IDictionary] $HTMLOptions,
        [string] $FilePath,
        [string] $SearchPath
    )
    $Today = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    # Lets define Write-Color rules
    if ($null -eq $DisplayConsole) {
        $WriteParameters = @{
            ShowTime   = $true
            LogFile    = ""
            TimeFormat = "yyyy-MM-dd HH:mm:ss"
        }
    } else {
        $WriteParameters = $DisplayConsole
    }

    if ($SearchPath) {
        if (Test-Path -LiteralPath $SearchPath) {
            try {
                $SummarySearch = Import-Clixml -LiteralPath $SearchPath -ErrorAction Stop
                #$SummarySearch = Get-Content -LiteralPath $SearchPath -Raw | ConvertFrom-Json
            } catch {
                Write-Color @WriteParameters -Text "[e]", " Couldn't load the file $SearchPath", ". Skipping...", $_.Exception.Message -Color White, Yellow, White, Yellow, White, Yellow, White
            }
        }
    }
    if (-not $SummarySearch) {
        $SummarySearch = [ordered] @{
            EmailSent        = [ordered] @{

            }
            EmailEscalations = [ordered] @{

            }
        }
    }

    $Summary = [ordered] @{}
    $Summary['Notify'] = [ordered] @{}
    $Summary['NotifyManager'] = [ordered] @{}
    $Summary['Rules'] = [ordered] @{}

    $CachedUsers = Find-Password -AsHashTable -OverwriteEmailProperty $OverwriteEmailProperty
    foreach ($Rule in $Rules) {
        # Go for each rule and check if the user is in any of those rules
        if ($Rule.Enable -eq $true) {
            Write-Color @WriteParameters -Text "[i]", " Processing rule ", $Rule.Name, ' status: ', $Rule.Enable -Color Yellow, White, Green, White, Green, White, Green, White
            # Lets create summary for the rule
            if (-not $Summary['Rules'][$Rule.Name] ) {
                $Summary['Rules'][$Rule.Name] = [ordered] @{}
            }
            # this will make sure to expand array of multiple arrays of ints if provided
            # for example: (-150..-100),(-60..0), 1, 2, 3
            $Rule.Reminders = $Rule.Reminders | ForEach-Object { $_ }
            foreach ($User in $CachedUsers.Values) {
                if ($User.Enabled -eq $false) {
                    # We don't want to have disabled users
                    continue
                }
                if ($Rule.LimitOU.Count -gt 0) {
                    # Rule defined that only user withi specific OU has to be found
                    $FoundOU = $false
                    foreach ($OU in $Rule.LimitOU) {
                        if ($User.OrganizationalUnit -like $OU) {
                            $FoundOU = $true
                            break
                        }
                    }
                    if (-not $FoundOU) {
                        continue
                    }
                }
                if ($Rule.LimitGroup.Count -gt 0) {
                    # Rule defined that only user withi specific group has to be found
                    $FoundGroup = $false
                    foreach ($Group in $Rule.LimitGroup) {
                        if ($User.MemberOf -contains $Group) {
                            $FoundGroup = $true
                            break
                        }
                    }
                    if (-not $FoundGroup) {
                        continue
                    }
                }

                if ($Summary['Notify'][$User.DistinguishedName]) {
                    # User already exists in the notifications - rules are overlapping, we only take the first one
                    continue
                }


                if ($Rule.PasswordNeverExpires -eq $true) {
                    $DaysToPasswordExpiry = $Rule.PasswordNeverExpiresDays - $User.PasswordDays
                    $User.DaysToExpire = $DaysToPasswordExpiry
                }

                # Lets find users that expire
                if ($User.DaysToExpire -in $Rule.Reminders) {
                    Write-Color @WriteParameters -Text "[i]", " User ", $User.DisplayName, " (", $User.UserPrincipalName, ")", " days to expire: ", $User.DaysToExpire -Color Yellow, White, Yellow, White, Yellow, White, White, Blue
                    $Summary['Notify'][$User.DistinguishedName] = [ordered] @{
                        User = $User
                        Rule = $Rule
                    }
                    $Summary['Rules'][$Rule.Name][$User.DistinguishedName] = [ordered] @{
                        User = $User
                        Rule = $Rule
                    }

                    if ($Rule.SendToManager) {
                        if ($Rule.SendToManager.Manager -and $Rule.SendToManager.Manager.Enable -eq $true -and $User.ManagerStatus -eq 'Enabled' -and $User.ManagerEmail -like "*@*") {
                            # Manager is enabled and has an email, this is standard situation for manager in AD
                            # But before we go and do that, maybe user wants to send emails to managers if those users are in specific group or OU
                            if ($Rule.SendToManager.Manager.LimitOU.Count -gt 0) {
                                # Rule defined that only user withi specific OU has to be found
                                $FoundOU = $false
                                foreach ($OU in $Rule.SendToManager.Manager.LimitOU) {
                                    if ($User.OrganizationalUnit -like $OU) {
                                        $FoundOU = $true
                                        break
                                    }
                                }
                                if (-not $FoundOU) {
                                    continue
                                }
                            }
                            if ($Rule.SendToManager.Manager.LimitGroup.Count -gt 0) {
                                # Rule defined that only user withi specific group has to be found
                                $FoundGroup = $false
                                foreach ($Group in $Rule.SendToManager.Manager.LimitGroup) {
                                    if ($User.MemberOf -contains $Group) {
                                        $FoundGroup = $true
                                        break
                                    }
                                }
                                if (-not $FoundGroup) {
                                    continue
                                }
                            }
                            $Splat = [ordered] @{
                                SummaryDictionary = $Summary['NotifyManager']
                                Type              = 'ManagerDefault'
                                ManagerType       = 'Ok'
                                Key               = $User.ManagerDN
                                User              = $User
                                Rule              = $Rule
                                #Enabled           = $true
                            }
                            Add-ManagerInformation @Splat
                        } else {
                            # Not compliant (missing, disabled, no email), covers all the below options
                            if ($Rule.SendToManager.ManagerNotCompliant -and $Rule.SendToManager.ManagerNotCompliant.Enable -and $Rule.SendToManager.ManagerNotCompliant.Manager) {

                                # But before we go and do that, maybe user wants to send emails to managers if those users are in specific group or OU
                                if ($Rule.SendToManager.ManagerNotCompliant.LimitOU.Count -gt 0) {
                                    # Rule defined that only user withi specific OU has to be found
                                    $FoundOU = $false
                                    foreach ($OU in $Rule.SendToManager.Manager.LimitOU) {
                                        if ($User.OrganizationalUnit -like $OU) {
                                            $FoundOU = $true
                                            break
                                        }
                                    }
                                    if (-not $FoundOU) {
                                        continue
                                    }
                                }
                                if ($Rule.SendToManager.ManagerNotCompliant.LimitGroup.Count -gt 0) {
                                    # Rule defined that only user withi specific group has to be found
                                    $FoundGroup = $false
                                    foreach ($Group in $Rule.SendToManager.Manager.LimitGroup) {
                                        if ($User.MemberOf -contains $Group) {
                                            $FoundGroup = $true
                                            break
                                        }
                                    }
                                    if (-not $FoundGroup) {
                                        continue
                                    }
                                }


                                if ($Rule.SendToManager.ManagerNotCompliant.MissingEmail -and $User.ManagerStatus -eq 'Enabled') {
                                    # Manager is enabled but missing email
                                    $Splat = [ordered] @{
                                        SummaryDictionary = $Summary['NotifyManager']
                                        Type              = 'ManagerNotCompliant'
                                        ManagerType       = 'No email'

                                        Key               = $Rule.SendToManager.ManagerNotCompliant.Manager
                                        User              = $User
                                        Rule              = $Rule

                                    }
                                    Add-ManagerInformation @Splat
                                } elseif ($Rule.SendToManager.ManagerNotCompliant.Disabled -and $User.ManagerStatus -eq 'Disabled') {
                                    # Manager is disabled, regardless if he/she has email
                                    $Splat = [ordered] @{
                                        SummaryDictionary = $Summary['NotifyManager']
                                        Type              = 'ManagerNotCompliant'
                                        ManagerType       = 'Manager disabled'
                                        Key               = $Rule.SendToManager.ManagerNotCompliant.Manager
                                        User              = $User
                                        Rule              = $Rule

                                    }
                                    Add-ManagerInformation @Splat
                                } elseif ($Rule.SendToManager.ManagerNotCompliant.LastLogon -and $User.ManagerLastLogonDays -ge $Rule.SendToManager.ManagerNotCompliant.LastLogonDays) {
                                    # Manager Last Logon over X days
                                    $Splat = [ordered] @{
                                        SummaryDictionary = $Summary['NotifyManager']
                                        Type              = 'ManagerNotCompliant'
                                        ManagerType       = 'Manager not logging in'
                                        Key               = $Rule.SendToManager.ManagerNotCompliant.Manager
                                        User              = $User
                                        Rule              = $Rule

                                    }
                                    Add-ManagerInformation @Splat
                                } elseif ($Rule.SendToManager.ManagerNotCompliant.Missing -and $User.ManagerStatus -eq 'Missing') {
                                    # Manager is missing
                                    $Splat = [ordered] @{
                                        SummaryDictionary = $Summary['NotifyManager']
                                        Type              = 'ManagerNotCompliant'
                                        ManagerType       = 'Manager not set'
                                        Key               = $Rule.SendToManager.ManagerNotCompliant.Manager
                                        User              = $User
                                        Rule              = $Rule

                                    }
                                    Add-ManagerInformation @Splat
                                }
                            }
                        }
                    }
                }
            }
        } else {
            Write-Color @WriteParameters -Text "[i]", " Processing rule ", $Rule.Name, ' status: ', $Rule.Enable -Color Red, White, Red, White, Red, White, Red, White
        }
    }

    if ($UserSection.Enable) {
        Write-Color @WriteParameters -Text "[i] Sending notifications to users " -Color White, Yellow, White, Yellow, White, Yellow, White
        $CountUsers = 0
        [Array] $SummaryUsersEmails = foreach ($Notify in $Summary['Notify'].Values) {
            $CountUsers++
            $EmailSplat = [ordered] @{}

            if ($Notify.User.DaysToExpire -ge 0) {
                if ($Notify.Rule.TemplatePreExpiry) {
                    # User uses template per rule
                    $EmailSplat.Template = $Notify.Rule.TemplatePreExpiry
                } elseif ($TemplatePreExpiry) {
                    # User uses global template
                    $EmailSplat.Template = $TemplatePreExpiry
                } else {
                    # User uses built-in template
                    $EmailSplat.Template = {

                    }
                }
                if ($Notify.Rule.TemplatePreExpirySubject) {
                    $EmailSplat.Subject = $Notify.Rule.TemplatePreExpirySubject
                } elseif ($TemplatePreExpirySubject) {
                    $EmailSplat.Subject = $TemplatePreExpirySubject
                } else {
                    $EmailSplat.Subject = '[Password] Your password will expire on $DateExpiry ($TimeToExpire days)'
                }
            } else {
                if ($Notify.Rule.TemplatePostExpiry) {
                    $EmailSplat.Template = $Notify.Rule.TemplatePostExpiry
                } elseif ($TemplatePostExpiry) {
                    $EmailSplat.Template = $TemplatePostExpiry
                } else {
                    $EmailSplat.Template = {

                    }
                }
                if ($Notify.Rule.TemplatePostExpirySubject) {
                    $EmailSplat.Subject = $Notify.Rule.TemplatePostExpirySubject
                } elseif ($TemplatePostExpirySubject) {
                    $EmailSplat.Subject = $TemplatePostExpirySubject
                } else {
                    $EmailSplat.Subject = '[Password] Your password expired on $DateExpiry ($TimeToExpire days ago)'
                }
            }
            $EmailSplat.User = $Notify.User
            $EmailSplat.EmailParameters = $EmailParameters

            if ($UserSection.SendToDefaultEmail -ne $true) {
                $EmailSplat.EmailParameters.To = $Notify.User.EmailAddress
            }
            if ($Notify.User.EmailAddress -like "*@*") {
                # Regardless if we send email to default email or to user, if user doesn't have email address we shouldn't send an email
                $EmailResult = Send-PasswordEmail @EmailSplat
                [PSCustomObject] @{
                    UserPrincipalName    = $EmailSplat.User.UserPrincipalName
                    SamAccountName       = $EmailSplat.User.SamAccountName
                    Domain               = $EmailSplat.User.Domain
                    Rule                 = $Notify.Rule.Name
                    Status               = $EmailResult.Status
                    StatusWhen           = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    StatusError          = $EmailResult.Error
                    SentTo               = $EmailResult.SentTo
                    DateExpiry           = $EmailSplat.User.DateExpiry
                    DaysToExpire         = $EmailSplat.User.DaysToExpire
                    PasswordExpired      = $EmailSplat.User.PasswordExpired
                    PasswordNeverExpires = $EmailSplat.User.PasswordNeverExpires
                    PasswordLastSet      = $EmailSplat.User.PasswordLastSet
                }
                if ($UserSection.SendCountMaximum -gt 0) {
                    if ($UserSection.SendCountMaximum -le $CountUsers) {
                        Write-Color @WriteParameters -Text "[i]", " Send count maximum reached. There may be more accounts that match the rule." -Color Red, DarkMagenta
                        break
                    }
                }
            } else {
                # Email not sent
                $EmailResult = @{
                    Status = $false
                    Error  = 'No email address for user'
                    SentTo = ''
                }
                [PSCustomObject] @{
                    UserPrincipalName    = $EmailSplat.User.UserPrincipalName
                    SamAccountName       = $EmailSplat.User.SamAccountName
                    Domain               = $EmailSplat.User.Domain
                    Rule                 = $Notify.Rule.Name
                    Status               = $EmailResult.Status
                    StatusWhen           = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    StatusError          = $EmailResult.Error
                    SentTo               = $EmailResult.SentTo
                    DateExpiry           = $EmailSplat.User.DateExpiry
                    DaysToExpire         = $EmailSplat.User.DaysToExpire
                    PasswordExpired      = $EmailSplat.User.PasswordExpired
                    PasswordNeverExpires = $EmailSplat.User.PasswordNeverExpires
                    PasswordLastSet      = $EmailSplat.User.PasswordLastSet
                }
            }
            if ($EmailResult.Status -eq $true) {
                Write-Color @WriteParameters -Text "[i]", " Sending ", $Notify.User.DisplayName, " (", $Notify.User.UserPrincipalName, ")", " status: ", $EmailResult.Status, ", details: ", $EmailResult.Error -Color Yellow, White, Yellow, White, Yellow, White, White, Blue, White, Blue
            } else {
                Write-Color @WriteParameters -Text "[i]", " Sending ", $Notify.User.DisplayName, " (", $Notify.User.UserPrincipalName, ")", " status: ", $EmailResult.Status, ", details: ", $EmailResult.Error -Color Yellow, White, Yellow, White, Yellow, White, White, Red, White, Red
            }
        }
        Write-Color @WriteParameters -Text "[i] Sending notifications to users (sent: ", $SummaryUsersEmails.Count, " out of ", $Summary['Notify'].Values.Count, ")" -Color White, Yellow, White, Yellow, White, Yellow, White
    } else {
        Write-Color @WriteParameters -Text "[i] Sending notifications to users is ", "disabled!" -Color White, Yellow, DarkMagenta
    }
    if ($ManagerSection.Enable) {
        Write-Color @WriteParameters -Text "[i] Sending notifications to managers " -Color White, Yellow, White, Yellow, White, Yellow, White
        $CountManagers = 0
        [Array] $SummaryManagersEmails = foreach ($Manager in $Summary['NotifyManager'].Keys) {
            $CountManagers++
            if ($CachedUsers[$Manager]) {
                # This user is "findable" in AD
                $ManagerUser = $CachedUsers[$Manager]
            } else {
                # This user is provided by user in config file
                $ManagerUser = $Summary['NotifyManager'][$Manager]['Manager']
            }
            [Array] $ManagedUsers = $Summary['NotifyManager'][$Manager]['ManagerDefault'].Values.Output
            [Array] $ManagedUsersManagerNotCompliant = $Summary['NotifyManager'][$Manager]['ManagerNotCompliant'].Values.Output

            $EmailSplat = [ordered] @{}

            if ($Summary['NotifyManager'][$Manager].ManagerDefault.Count -gt 0) {
                if ($TemplateManager) {
                    # User uses global template
                    $EmailSplat.Template = $TemplateManager
                } else {
                    # User uses built-in template
                    $EmailSplat.Template = {

                    }
                }
                if ($TemplateManagerSubject) {
                    $EmailSplat.Subject = $TemplateManagerSubject
                } else {
                    $EmailSplat.Subject = "[Password Expiring] Dear Manager - Your accounts are expiring!"
                }
            } elseif ($Summary['NotifyManager'][$Manager].ManagerNotCompliant.Count -gt 0) {
                if ($TemplateManagerNotCompliant) {
                    # User uses global template
                    $EmailSplat.Template = $TemplateManagerNotCompliant
                } else {
                    # User uses built-in template
                    $EmailSplat.Template = {

                    }
                }
                if ($TemplateManagerNotCompliantSubject) {
                    $EmailSplat.Subject = $TemplateManagerNotCompliantSubject
                } else {
                    $EmailSplat.Subject = "[Password Escalation] Accounts are expiring with non-compliant manager"
                }
            }

            $EmailSplat.User = $ManagerUser
            $EmailSplat.ManagedUsers = $ManagedUsers
            $EmailSplat.ManagedUsersManagerNotCompliant = $ManagedUsersManagerNotCompliant
            #$EmailSplat.ManagedUsersManagerDisabled = $ManagedUsersManagerDisabled
            #$EmailSplat.ManagedUsersManagerMissing = $ManagedUsersManagerMissing
            #$EmailSplat.ManagedUsersManagerMissingEmail = $ManagedUsersManagerMissingEmail
            $EmailSplat.EmailParameters = $EmailParameters

            if ($ManagerSection.SendToDefaultEmail -ne $true) {
                $EmailSplat.EmailParameters.To = $ManagerUser.EmailAddress
            }

            $EmailResult = Send-PasswordEmail @EmailSplat
            [PSCustomObject] @{
                DisplayName              = $ManagerUser.DisplayName
                SamAccountName           = $ManagerUser.SamAccountName
                Domain                   = $ManagerUser.Domain
                Status                   = $EmailResult.Status
                StatusWhen               = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                SentTo                   = $EmailResult.SentTo
                StatusError              = $EmailResult.Error
                Accounts                 = $ManagedUsers.SamAccountName
                AccountsCount            = $ManagedUsers.Count
                Template                 = 'Unknown'
                ManagerNotCompliant      = $ManagedUsersManagerNotCompliant.SamAccountName
                ManagerNotCompliantCount = $ManagedUsersManagerNotCompliant.Count
                #ManagerDisabled          = $ManagedUsersManagerDisabled.SamAccountName
                #ManagerDisabledCount     = $ManagedUsersManagerDisabled.Count
                #ManagerMissing           = $ManagedUsersManagerMissing.SamAccountName
                #ManagerMissingCount      = $ManagedUsersManagerMissing.Count
                #ManagerMissingEmail      = $ManagedUsersManagerMissingEmail.SamAccountName
                #ManagerMissingEmailCount = $ManagedUsersManagerMissingEmail.Count
            }
            if ($ManagerSection.SendCountMaximum -gt 0) {
                if ($ManagerSection.SendCountMaximum -le $CountManagers) {
                    Write-Color @WriteParameters -Text "[i]", " Send count maximum reached. There may be more managers that match the rule." -Color Red, DarkMagenta
                    break
                }
            }
        }
        Write-Color @WriteParameters -Text "[i] Sending notifications to managers (sent: ", $SummaryManagersEmails.Count, " out of ", $Summary['NotifyManager'].Values.Count, ")" -Color White, Yellow, White, Yellow, White, Yellow, White
        #Write-Color @WriteParameters -Text "[i] Sending notifications to managers (sent: ", $SummaryManagersEmails.Count, ")" -Color White, Yellow, White, Yellow, White, Yellow, White
    } else {
        Write-Color @WriteParameters -Text "[i] Sending notifications to managers is ", "disabled!" -Color White, Yellow, DarkMagenta
    }

    # Create report
    New-HTML {
        New-TableOption -DataStore JavaScript -ArrayJoin -BoolAsString
        New-HTMLTab -Name 'All Users' {
            New-HTMLTable -DataTable $CachedUsers.Values -Filtering {
                New-TableCondition -Name 'PasswordExpired' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $true -ComparisonType string
                New-TableCondition -Name 'PasswordNeverExpires' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $false -ComparisonType string
            }
        }
        foreach ($Rule in  $Summary['Rules'].Keys) {
            New-HTMLTab -Name $Rule {
                New-HTMLTable -DataTable $Summary['Rules'][$Rule].Values.User -Filtering {
                    New-TableCondition -Name 'Enabled' -BackgroundColor LawnGreen -FailBackgroundColor BlueSmoke -Value $true -ComparisonType string
                    New-TableCondition -Name 'PasswordExpired' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $false -ComparisonType string
                    New-TableCondition -Name 'PasswordNeverExpires' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $false -ComparisonType string
                }
            }
        }
        New-HTMLTab -Name 'Email sent to users' {
            New-HTMLTable -DataTable $SummaryUsersEmails {
                New-TableHeader -Names 'Status', 'StatusError', 'SentTo', 'StatusWhen' -Title 'Email Summary'
                New-TableCondition -Name 'Status' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $true -ComparisonType string -HighlightHeaders 'Status', 'StatusWhen', 'StatusError', 'SentTo'
                New-TableCondition -Name 'PasswordExpired' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $false -ComparisonType string
                New-TableCondition -Name 'PasswordNeverExpires' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $false -ComparisonType string
            }
        }
        New-HTMLTab -Name 'Email sent to manager' {
            New-HTMLTable -DataTable $SummaryManagersEmails {
                New-TableHeader -Names 'Status', 'StatusError', 'SentTo', 'StatusWhen' -Title 'Email Summary'
                New-TableCondition -Name 'Status' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $true -ComparisonType string -HighlightHeaders 'Status', 'StatusWhen', 'StatusError', 'SentTo'
            }
        }
    } -ShowHTML:$HTMLOptions.ShowHTML -FilePath $FilePath -Online:$HTMLOptions.Online

    if ($SearchPath) {

        $SummarySearch['EmailSent'][$Today] = $SummaryUsersEmails
        $SummarySearch['EmailEscalations'][$Today] = $SummaryManagersEmails

        try {
            $SummarySearch | Export-Clixml -LiteralPath $SearchPath
            #$SummarySearch | ConvertTo-Json | Out-File -LiteralPath $SearchPath
        } catch {
            Write-Color @WriteParameters -Text "[e]", " Couldn't save to file $SearchPath", ". Error: ", $_.Exception.Message -Color White, Yellow, White, Yellow, White, Yellow, White
        }
    }
}