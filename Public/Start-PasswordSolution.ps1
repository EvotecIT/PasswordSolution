function Start-PasswordSolution {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary] $EmailParameters,
        [string] $OverwriteEmailProperty,
        [Parameter(Mandatory)][System.Collections.IDictionary] $UserSection,
        [Parameter(Mandatory)][System.Collections.IDictionary] $ManagerSection,
        [Parameter(Mandatory)][System.Collections.IDictionary] $SecuritySection,
        [Parameter(Mandatory)][System.Collections.IDictionary] $AdminSection,
        [Parameter(Mandatory)][Array] $Rules,
        [scriptblock] $TemplatePreExpiry,
        [string] $TemplatePreExpirySubject,
        [scriptblock] $TemplatePostExpiry,
        [string] $TemplatePostExpirySubject,
        [Parameter(Mandatory)][scriptblock] $TemplateManager,
        [Parameter(Mandatory)][string] $TemplateManagerSubject,
        [Parameter(Mandatory)][scriptblock] $TemplateSecurity,
        [Parameter(Mandatory)][string] $TemplateSecuritySubject,
        [Parameter(Mandatory)][scriptblock] $TemplateManagerNotCompliant,
        [Parameter(Mandatory)][string] $TemplateManagerNotCompliantSubject,
        [Parameter(Mandatory)][scriptblock] $TemplateAdmin,
        [Parameter(Mandatory)][string] $TemplateAdminSubject,
        [Parameter(Mandatory)][System.Collections.IDictionary] $Logging,
        [Parameter(Mandatory)][System.Collections.IDictionary] $HTMLOptions,
        [string] $FilePath,
        [string] $SearchPath
    )
    $TimeStart = Start-TimeLog
    $Script:Reporting = [ordered] @{}
    $Script:Reporting['Version'] = Get-GitHubVersion -Cmdlet 'Start-PasswordSolution' -RepositoryOwner 'evotecit' -RepositoryName 'PasswordSolution'
    $TodayDate = Get-Date
    $Today = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    if (-not $Logging) {
        $Logging = @{
            ShowTime   = $true
            LogFile    = ""
            TimeFormat = "yyyy-MM-dd HH:mm:ss"
        }
    }
    $PSDefaultParameterValues = @{
        "Write-Color:LogFile"    = $Logging.LogFile
        "Write-Color:ShowTime"   = $Logging.ShowTime
        "Write-Color:TimeFormat" = $Logging.TimeFormat
    }

    if ($Logging.LogFile) {
        $FolderPath = [io.path]::GetDirectoryName($Logging.LogFile)
        if (-not (Test-Path -LiteralPath $FolderPath)) {
            $null = New-Item -Path $FolderPath -ItemType Directory -Force
        }
        $CurrentLogs = Get-ChildItem -LiteralPath $FolderPath | Sort-Object -Property CreationTime -Descending | Select-Object -Skip $Logging.LogMaximum
        if ($CurrentLogs) {
            Write-Color -Text '[i] ', "Logs directory has more than ", $Logging.LogMaximum, " log files. Cleanup required..." -Color Yellow, DarkCyan, Red, DarkCyan
            foreach ($Log in $CurrentLogs) {
                try {
                    Remove-Item -LiteralPath $Log.FullName -Confirm:$false
                    Write-Color -Text '[+] ', "Deleted ", "$($Log.FullName)" -Color Yellow, White, Green
                } catch {
                    Write-Color -Text '[-] ', "Couldn't delete log file $($Log.FullName). Error: ', "$($_.Exception.Message) -Color Yellow, White, Red
                }
            }
        }
    }

    if ($SearchPath) {
        if (Test-Path -LiteralPath $SearchPath) {
            try {
                Write-Color -Text "[i]", " Loading file ", $SearchPath -Color White, Yellow, White, Yellow, White, Yellow, White
                $SummarySearch = Import-Clixml -LiteralPath $SearchPath -ErrorAction Stop
            } catch {
                Write-Color -Text "[e]", " Couldn't load the file $SearchPath", ". Skipping...", $_.Exception.Message -Color White, Yellow, White, Yellow, White, Yellow, White
            }
        }
    }
    if (-not $SummarySearch) {
        $SummarySearch = [ordered] @{
            EmailSent        = [ordered] @{

            }
            EmailManagers    = [ordered] @{

            }
            EmailEscalations = [ordered] @{

            }

        }
    }

    $Summary = [ordered] @{}
    $Summary['Notify'] = [ordered] @{}
    $Summary['NotifyManager'] = [ordered] @{}
    $Summary['NotifySecurity'] = [ordered] @{}
    $Summary['Rules'] = [ordered] @{}


    Write-Color -Text "[i]", " Starting process to find expiring users" -Color Yellow, White, Green, White, Green, White, Green, White
    $CachedUsers = Find-Password -AsHashTable -OverwriteEmailProperty $OverwriteEmailProperty
    foreach ($Rule in $Rules) {
        # Go for each rule and check if the user is in any of those rules
        if ($Rule.Enable -eq $true) {
            Write-Color -Text "[i]", " Processing rule ", $Rule.Name, ' status: ', $Rule.Enable -Color Yellow, White, Green, White, Green, White, Green, White
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
                if ($Rule.ExcludeOU.Count -gt 0) {
                    $FoundOU = $false
                    foreach ($OU in $Rule.ExcludeOU) {
                        if ($User.OrganizationalUnit -like $OU) {
                            $FoundOU = $true
                            break
                        }
                    }
                    # if OU is found we need to exclude the user
                    if ($FoundOU) {
                        continue
                    }
                }
                if ($Rule.IncludeOU.Count -gt 0) {
                    # Rule defined that only user withi specific OU has to be found
                    $FoundOU = $false
                    foreach ($OU in $Rule.IncludeOU) {
                        if ($User.OrganizationalUnit -like $OU) {
                            $FoundOU = $true
                            break
                        }
                    }
                    if (-not $FoundOU) {
                        continue
                    }
                }
                if ($Rule.ExcludeGroup.Count -gt 0) {
                    # Rule defined that only user withi specific group has to be found
                    $FoundGroup = $false
                    foreach ($Group in $Rule.ExcludeGroup) {
                        if ($User.MemberOf -contains $Group) {
                            $FoundGroup = $true
                            break
                        }
                    }
                    # If found, we need to skip user
                    if ($FoundGroup) {
                        continue
                    }
                }
                if ($Rule.IncludeGroup.Count -gt 0) {
                    # Rule defined that only user withi specific group has to be found
                    $FoundGroup = $false
                    foreach ($Group in $Rule.IncludeGroup) {
                        if ($User.MemberOf -contains $Group) {
                            $FoundGroup = $true
                            break
                        }
                    }
                    if (-not $FoundGroup) {
                        continue
                    }
                }
                if ($Rule.IncludeName.Count -gt 0) {
                    $IncludeName = $false
                    foreach ($Name in $Rule.IncludeName) {
                        foreach ($Property in $Rule.IncludeNameProperties) {
                            if ($User.$Property -like $Name) {
                                $IncludeName = $true
                                break
                            }
                        }
                        if ($IncludeName) {
                            break
                        }
                    }
                    if (-not $IncludeName) {
                        continue
                    }
                }
                if ($Summary['Notify'][$User.DistinguishedName] -and $Summary['Notify'][$User.DistinguishedName].ProcessManagersOnly -ne $true) {
                    # User already exists in the notifications - rules are overlapping, we only take the first one
                    # We also check for ProcessManagersOnly because we don't want first rule to ignore any other rules for users
                    continue
                }
                if ($Rule.IncludePasswordNeverExpires -and $Rule.IncludeExpiring) {
                    if ($User.PasswordNeverExpires -eq $true) {
                        $DaysToPasswordExpiry = $Rule.PasswordNeverExpiresDays - $User.PasswordDays
                        $User.DaysToExpire = $DaysToPasswordExpiry
                    }
                } elseif ($Rule.IncludeExpiring) {
                    if ($User.PasswordNeverExpires -eq $true) {
                        # we skip those that never expire
                        continue
                    }
                } elseif ($Rule.IncludePasswordNeverExpires) {
                    if ($User.PasswordNeverExpires -eq $true) {
                        $DaysToPasswordExpiry = $Rule.PasswordNeverExpiresDays - $User.PasswordDays
                        $User.DaysToExpire = $DaysToPasswordExpiry
                    } else {
                        # we skip users who expire
                        continue
                    }
                } else {
                    Write-Color -Text "[i]", " Processing rule ", $Rule.Name, " doesn't include IncludePasswordNeverExpires nor IncludeExpiring so skipping." -Color Yellow, White, Green, White, Green, White, Green, White
                    continue
                }
                # Lets find users that expire, and match our rule
                if ($User.DaysToExpire -in $Rule.Reminders) {
                    # check if we need to notify user or just manager
                    if (-not $Rule.ProcessManagersOnly) {
                        if ($Logging.NotifyOnUserMatchingRule) {
                            Write-Color -Text "[i]", " User ", $User.DisplayName, " (", $User.UserPrincipalName, ")", " days to expire: ", $User.DaysToExpire -Color Yellow, White, Yellow, White, Yellow, White, White, Blue
                        }
                        $Summary['Notify'][$User.DistinguishedName] = [ordered] @{
                            User                = $User
                            Rule                = $Rule
                            ProcessManagersOnly = $Rule.ProcessManagersOnly
                        }
                        $Summary['Rules'][$Rule.Name][$User.DistinguishedName] = [ordered] @{
                            User                = $User
                            Rule                = $Rule
                            ProcessManagersOnly = $Rule.ProcessManagersOnly
                        }
                    }
                }
                if ($Rule.SendToManager) {
                    if ($Rule.SendToManager.Manager -and $Rule.SendToManager.Manager.Enable -eq $true -and $User.ManagerStatus -eq 'Enabled' -and $User.ManagerEmail -like "*@*") {
                        $SendToManager = $true
                        # Manager is enabled and has an email, this is standard situation for manager in AD
                        # But before we go and do that, maybe user wants to send emails to managers if those users are in specific group or OU
                        if ($Rule.SendToManager.Manager.IncludeOU.Count -gt 0) {
                            # Rule defined that only user withi specific OU has to be found
                            $FoundOU = $false
                            foreach ($OU in $Rule.SendToManager.Manager.IncludeOU) {
                                if ($User.OrganizationalUnit -like $OU) {
                                    $FoundOU = $true
                                    break
                                }
                            }
                            if (-not $FoundOU) {
                                $SendToManager = $false
                            }
                        }
                        if ($SendToManager -and $Rule.SendToManager.Manager.ExcludeOU.Count -gt 0) {
                            $FoundOU = $false
                            foreach ($OU in $Rule.SendToManager.Manager.ExcludeOU) {
                                if ($User.OrganizationalUnit -like $OU) {
                                    $FoundOU = $true
                                    break
                                }
                            }
                            # if OU is found we need to exclude the user
                            if ($FoundOU) {
                                $SendToManager = $false
                            }
                        }
                        if ($SendToManager -and $Rule.SendToManager.Manager.ExcludeGroup.Count -gt 0) {
                            # Rule defined that only user withi specific group has to be found
                            $FoundGroup = $false
                            foreach ($Group in $Rule.SendToManager.Manager.ExcludeGroup) {
                                if ($User.MemberOf -contains $Group) {
                                    $FoundGroup = $true
                                    break
                                }
                            }
                            # if Group found, we need to skip this user
                            if ($FoundGroup) {
                                $SendToManager = $false
                            }
                        }
                        if ($SendToManager -and $Rule.SendToManager.Manager.IncludeGroup.Count -gt 0) {
                            # Rule defined that only user withi specific group has to be found
                            $FoundGroup = $false
                            foreach ($Group in $Rule.SendToManager.Manager.IncludeGroup) {
                                if ($User.MemberOf -contains $Group) {
                                    $FoundGroup = $true
                                    break
                                }
                            }
                            if (-not $FoundGroup) {
                                $SendToManager = $false
                            }
                        }

                        if ($SendToManager) {
                            $SendToManager = $false
                            if ($Rule.SendToManager.Manager.Reminders.Default.Enable -eq $true -and $null -eq $Rule.SendToManager.Manager.Reminders.Default.Reminder -and $User.DaysToExpire -in $Rule.Reminders) {
                                # Use default reminder as per user, not per manager
                                $SendToManager = $true
                            } elseif ($Rule.SendToManager.Manager.Reminders.Default.Enable -eq $true -and $User.DaysToExpire -in $Rule.SendToManager.Manager.Reminders.Default.Reminder) {
                                # User manager reminder as per manager config
                                $SendToManager = $true
                            }
                            if (-not $SendToManager -and $Rule.SendToManager.Manager.Reminders.OnDay -and $Rule.SendToManager.Manager.Reminders.OnDay.Enable -eq $true) {
                                foreach ($Day in $Rule.SendToManager.Manager.Reminders.OnDay.Days) {
                                    if ($Day -eq "$($TodayDate.DayOfWeek)") {
                                        if ($Rule.SendToManager.Manager.Reminders.OnDay.ComparisonType -eq 'lt') {
                                            if ($User.DaysToExpire -lt $Rule.SendToManager.Manager.Reminders.OnDay.Reminder) {
                                                $SendToManager = $true
                                                break
                                            }
                                        } elseif ($Rule.SendToManager.Manager.Reminders.OnDay.ComparisonType -eq 'gt') {
                                            if ($User.DaysToExpire -gt $Rule.SendToManager.Manager.Reminders.OnDay.Reminder) {
                                                $SendToManager = $true
                                                break
                                            }
                                        } elseif ($Rule.SendToManager.Manager.Reminders.OnDay.ComparisonType -eq 'eq') {
                                            if ($User.DaysToExpire -eq $Rule.SendToManager.Manager.Reminders.OnDay.Reminder) {
                                                $SendToManager = $true
                                                break
                                            }
                                        } elseif ($Rule.SendtoManager.Manager.Reminders.OnDay.ComparisonType -eq 'in') {
                                            if ($User.DaysToExpire -in $Rule.SendToManager.Manager.Reminders.OnDay.Reminder) {
                                                $SendToManager = $true
                                                break
                                            }
                                        }
                                    }
                                }
                            }
                            if (-not $SendToManager -and $Rule.SendToManager.Manager.Reminders.OnDayOfMonth -and $Rule.SendToManager.Manager.Reminders.OnDayOfMonth.Enable -eq $true) {
                                foreach ($Day in $Rule.SendToManager.Manager.Reminders.OnDayOfMonth.Days) {
                                    if ($Day -eq $TodayDate.Day) {
                                        if ($Rule.SendToManager.Manager.Reminders.OnDayOfMonth.ComparisonType -eq 'lt') {
                                            if ($User.DaysToExpire -lt $Rule.SendToManager.Manager.Reminders.OnDayOfMonth.Reminder) {
                                                $SendToManager = $true
                                                break
                                            }
                                        } elseif ($Rule.SendToManager.Manager.Reminders.OnDayOfMonth.ComparisonType -eq 'gt') {
                                            if ($User.DaysToExpire -gt $Rule.SendToManager.Manager.Reminders.OnDayOfMonth.Reminder) {
                                                $SendToManager = $true
                                                break
                                            }
                                        } elseif ($Rule.SendToManager.Manager.Reminders.OnDayOfMonth.ComparisonType -eq 'eq') {
                                            if ($User.DaysToExpire -eq $Rule.SendToManager.Manager.Reminders.OnDayOfMonth.Reminder) {
                                                $SendToManager = $true
                                                break
                                            }
                                        } elseif ($Rule.SendtoManager.Manager.Reminders.OnDayOfMonth.ComparisonType -eq 'in') {
                                            if ($User.DaysToExpire -in $Rule.SendToManager.Manager.Reminders.OnDayOfMonth.Reminder) {
                                                $SendToManager = $true
                                                break
                                            }
                                        }
                                    }
                                }
                            }
                            if ($SendToManager) {
                                $Splat = [ordered] @{
                                    SummaryDictionary = $Summary['NotifyManager']
                                    Type              = 'ManagerDefault'
                                    ManagerType       = 'Ok'
                                    Key               = $User.ManagerDN
                                    User              = $User
                                    Rule              = $Rule
                                }
                                Add-ManagerInformation @Splat
                            }
                        }
                    }
                }
                # Lets find users that have no manager, manager is not enabled or manager has no email
                if ($Rule.SendToManager -and $Rule.SendToManager.ManagerNotCompliant -and $Rule.SendToManager.ManagerNotCompliant.Enable -eq $true -and $Rule.SendToManager.ManagerNotCompliant.Manager) {
                    # Not compliant (missing, disabled, no email), covers all the below options
                    if ($Rule.SendToManager.ManagerNotCompliant -and $Rule.SendToManager.ManagerNotCompliant.Enable -and $Rule.SendToManager.ManagerNotCompliant.Manager) {
                        $ManagerNotCompliant = $false
                        if ($Rule.SendToManager.ManagerNotCompliant.Reminders) {
                            if ($Rule.SendToManager.ManagerNotCompliant.Reminders.Default -and $Rule.SendToManager.ManagerNotCompliant.Reminders.Default.Enable -eq $true) {
                                $Rule.SendToManager.ManagerNotCompliant.Reminders.Default.Reminder = $Rule.SendToManager.ManagerNotCompliant.Reminders.Default.Reminder | ForEach-Object { $_ }
                                if ($User.DaysToExpire -in $Rule.SendToManager.ManagerNotCompliant.Reminders.Default.Reminder) {
                                    $ManagerNotCompliant = $true
                                }
                            }
                            if ($Rule.SendToManager.ManagerNotCompliant.Reminders.OnDay -and $Rule.SendToManager.ManagerNotCompliant.Reminders.OnDay.Enable -eq $true) {
                                foreach ($Day in $Rule.SendToManager.ManagerNotCompliant.Reminders.OnDay.Days) {
                                    if ($Day -eq "$($TodayDate.DayOfWeek)") {
                                        if ($Rule.SendToManager.ManagerNotCompliant.Reminders.OnDay.ComparisonType -eq 'lt') {
                                            if ($User.DaysToExpire -lt $Rule.SendToManager.ManagerNotCompliant.Reminders.OnDay.Reminder) {
                                                $ManagerNotCompliant = $true
                                                break
                                            }
                                        } elseif ($Rule.SendToManager.ManagerNotCompliant.Reminders.OnDay.ComparisonType -eq 'gt') {
                                            if ($User.DaysToExpire -gt $Rule.SendToManager.ManagerNotCompliant.Reminders.OnDay.Reminder) {
                                                $ManagerNotCompliant = $true
                                                break
                                            }
                                        } elseif ($Rule.SendToManager.ManagerNotCompliant.Reminders.OnDay.ComparisonType -eq 'eq') {
                                            if ($User.DaysToExpire -eq $Rule.SendToManager.ManagerNotCompliant.Reminders.OnDay.Reminder) {
                                                $ManagerNotCompliant = $true
                                                break
                                            }
                                        } elseif ($Rule.SendtoManager.ManagerNotCompliant.Reminders.OnDay.ComparisonType -eq 'in') {
                                            if ($User.DaysToExpire -in $Rule.SendToManager.ManagerNotCompliant.Reminders.OnDay.Reminder) {
                                                $ManagerNotCompliant = $true
                                                break
                                            }
                                        }
                                    }
                                }
                            }
                            if ($Rule.SendToManager.ManagerNotCompliant.Reminders.OnDayOfMonth -and $Rule.SendToManager.ManagerNotCompliant.Reminders.OnDayOfMonth.Enable -eq $true) {
                                foreach ($Day in $Rule.SendToManager.ManagerNotCompliant.Reminders.OnDayOfMonth.Days) {
                                    if ($Day -eq $TodayDate.Day) {
                                        if ($Rule.SendToManager.ManagerNotCompliant.Reminders.OnDayOfMonth.ComparisonType -eq 'lt') {
                                            if ($User.DaysToExpire -lt $Rule.SendToManager.ManagerNotCompliant.Reminders.OnDayOfMonth.Reminder) {
                                                $ManagerNotCompliant = $true
                                                break
                                            }
                                        } elseif ($Rule.SendToManager.ManagerNotCompliant.Reminders.OnDayOfMonth.ComparisonType -eq 'gt') {
                                            if ($User.DaysToExpire -gt $Rule.SendToManager.ManagerNotCompliant.Reminders.OnDayOfMonth.Reminder) {
                                                $ManagerNotCompliant = $true
                                                break
                                            }
                                        } elseif ($Rule.SendToManager.ManagerNotCompliant.Reminders.OnDayOfMonth.ComparisonType -eq 'eq') {
                                            if ($User.DaysToExpire -eq $Rule.SendToManager.ManagerNotCompliant.Reminders.OnDayOfMonth.Reminder) {
                                                $ManagerNotCompliant = $true
                                                break
                                            }
                                        } elseif ($Rule.SendtoManager.ManagerNotCompliant.Reminders.OnDayOfMonth.ComparisonType -eq 'in') {
                                            if ($User.DaysToExpire -in $Rule.SendToManager.ManagerNotCompliant.Reminders.OnDayOfMonth.Reminder) {
                                                $ManagerNotCompliant = $true
                                                break
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if ($ManagerNotCompliant -eq $true) {
                            # But before we go and do that, maybe user wants to send emails to managers only if those users are in specific group or OU
                            if ($Rule.SendToManager.ManagerNotCompliant.IncludeOU.Count -gt 0) {
                                # Rule defined that only user withi specific OU has to be found
                                $FoundOU = $false
                                foreach ($OU in $Rule.SendToManager.ManagerNotCompliant.IncludeOU) {
                                    if ($User.OrganizationalUnit -like $OU) {
                                        $FoundOU = $true
                                        break
                                    }
                                }
                                if (-not $FoundOU) {
                                    continue
                                }
                            }
                            if ($Rule.SendToManager.ManagerNotCompliant.ExcludeOU.Count -gt 0) {
                                $FoundOU = $false
                                foreach ($OU in $Rule.SendToManager.ManagerNotCompliant.ExcludeOU) {
                                    if ($User.OrganizationalUnit -like $OU) {
                                        $FoundOU = $true
                                        break
                                    }
                                }
                                # if OU is found we need to exclude the user
                                if ($FoundOU) {
                                    continue
                                }
                            }
                            if ($Rule.SendToManager.ManagerNotCompliant.ExcludeGroup.Count -gt 0) {
                                # Rule defined that only user withi specific group has to be found
                                $FoundGroup = $false
                                foreach ($Group in $Rule.SendToManager.ManagerNotCompliant.ExcludeGroup) {
                                    if ($User.MemberOf -contains $Group) {
                                        $FoundGroup = $true
                                        break
                                    }
                                }
                                # if Group found, we need to skip this user
                                if ($FoundGroup) {
                                    continue
                                }
                            }
                            if ($Rule.SendToManager.ManagerNotCompliant.IncludeGroup.Count -gt 0) {
                                # Rule defined that only user withi specific group has to be found
                                $FoundGroup = $false
                                foreach ($Group in $Rule.SendToManager.ManagerNotCompliant.IncludeGroup) {
                                    if ($User.MemberOf -contains $Group) {
                                        $FoundGroup = $true
                                        break
                                    }
                                }
                                if (-not $FoundGroup) {
                                    continue
                                }
                            }
                            if ($Rule.SendToManager.ManagerNotCompliant.MissingEmail -and $User.ManagerStatus -in 'Enabled, bad email', 'No email') {
                                # Manager is enabled but missing email
                                $Splat = [ordered] @{
                                    SummaryDictionary = $Summary['NotifyManager']
                                    Type              = 'ManagerNotCompliant'
                                    ManagerType       = if ($User.ManagerStatus -eq 'Enabled, bad email') { 'Manager has bad email' } else { 'Manager has no email' }
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
                # Lets find users that require escalation
                if ($Rule.SendToManager -and $Rule.SendToManager.SecurityEscalation -and $Rule.SendToManager.SecurityEscalation.Enable -eq $true -and $Rule.SendToManager.SecurityEscalation.Manager) {
                    $SecurityEscalation = $false
                    if ($Rule.SendToManager.SecurityEscalation.Reminders) {
                        if ($Rule.SendToManager.SecurityEscalation.Reminders.Default -and $Rule.SendToManager.SecurityEscalation.Reminders.Default.Enable -eq $true) {
                            $Rule.SendToManager.SecurityEscalation.Reminders.Default.Reminder = $Rule.SendToManager.SecurityEscalation.Reminders.Default.Reminder | ForEach-Object { $_ }
                            if ($User.DaysToExpire -in $Rule.SendToManager.SecurityEscalation.Reminders.Default.Reminder) {
                                $SecurityEscalation = $true
                            }
                        }
                        if ($Rule.SendToManager.SecurityEscalation.Reminders.OnDay -and $Rule.SendToManager.SecurityEscalation.Reminders.OnDay.Enable -eq $true) {
                            foreach ($Day in $Rule.SendToManager.SecurityEscalation.Reminders.OnDay.Days) {
                                if ($Day -eq "$($TodayDate.DayOfWeek)") {
                                    if ($Rule.SendToManager.SecurityEscalation.Reminders.OnDay.ComparisonType -eq 'lt') {
                                        if ($User.DaysToExpire -lt $Rule.SendToManager.SecurityEscalation.Reminders.OnDay.Reminder) {
                                            $SecurityEscalation = $true
                                            break
                                        }
                                    } elseif ($Rule.SendToManager.SecurityEscalation.Reminders.OnDay.ComparisonType -eq 'gt') {
                                        if ($User.DaysToExpire -gt $Rule.SendToManager.SecurityEscalation.Reminders.OnDay.Reminder) {
                                            $SecurityEscalation = $true
                                            break
                                        }
                                    } elseif ($Rule.SendToManager.SecurityEscalation.Reminders.OnDay.ComparisonType -eq 'eq') {
                                        if ($User.DaysToExpire -eq $Rule.SendToManager.SecurityEscalation.Reminders.OnDay.Reminder) {
                                            $SecurityEscalation = $true
                                            break
                                        }
                                    } elseif ($Rule.SendtoManager.SecurityEscalation.Reminders.OnDay.ComparisonType -eq 'in') {
                                        if ($User.DaysToExpire -in $Rule.SendToManager.SecurityEscalation.Reminders.OnDay.Reminder) {
                                            $SecurityEscalation = $true
                                            break
                                        }
                                    }
                                }
                            }
                        }
                        if ($Rule.SendToManager.SecurityEscalation.Reminders.OnDayOfMonth -and $Rule.SendToManager.SecurityEscalation.Reminders.OnDayOfMonth.Enable -eq $true) {
                            foreach ($Day in $Rule.SendToManager.SecurityEscalation.Reminders.OnDayOfMonth.Days) {
                                if ($Day -eq $TodayDate.Day) {
                                    if ($Rule.SendToManager.SecurityEscalation.Reminders.OnDayOfMonth.ComparisonType -eq 'lt') {
                                        if ($User.DaysToExpire -lt $Rule.SendToManager.SecurityEscalation.Reminders.OnDayOfMonth.Reminder) {
                                            $SecurityEscalation = $true
                                            break
                                        }
                                    } elseif ($Rule.SendToManager.SecurityEscalation.Reminders.OnDayOfMonth.ComparisonType -eq 'gt') {
                                        if ($User.DaysToExpire -gt $Rule.SendToManager.SecurityEscalation.Reminders.OnDayOfMonth.Reminder) {
                                            $SecurityEscalation = $true
                                            break
                                        }
                                    } elseif ($Rule.SendToManager.SecurityEscalation.Reminders.OnDayOfMonth.ComparisonType -eq 'eq') {
                                        if ($User.DaysToExpire -eq $Rule.SendToManager.SecurityEscalation.Reminders.OnDayOfMonth.Reminder) {
                                            $SecurityEscalation = $true
                                            break
                                        }
                                    } elseif ($Rule.SendtoManager.SecurityEscalation.Reminders.OnDayOfMonth.ComparisonType -eq 'in') {
                                        if ($User.DaysToExpire -in $Rule.SendToManager.SecurityEscalation.Reminders.OnDayOfMonth.Reminder) {
                                            $SecurityEscalation = $true
                                            break
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if ($SecurityEscalation) {
                        if ($Rule.SendToManager.SecurityEscalation.IncludeOU.Count -gt 0) {
                            # Rule defined that only user withi specific OU has to be found
                            $FoundOU = $false
                            foreach ($OU in $Rule.SendToManager.SecurityEscalation.IncludeOU) {
                                if ($User.OrganizationalUnit -like $OU) {
                                    $FoundOU = $true
                                    break
                                }
                            }
                            if (-not $FoundOU) {
                                continue
                            }
                        }
                        if ($Rule.SendToManager.SecurityEscalation.ExcludeOU.Count -gt 0) {
                            $FoundOU = $false
                            foreach ($OU in $Rule.SendToManager.SecurityEscalation.ExcludeOU) {
                                if ($User.OrganizationalUnit -like $OU) {
                                    $FoundOU = $true
                                    break
                                }
                            }
                            # if OU is found we need to exclude the user
                            if ($FoundOU) {
                                continue
                            }
                        }
                        if ($Rule.SendToManager.SecurityEscalation.ExcludeGroup.Count -gt 0) {
                            # Rule defined that only user withi specific group has to be found
                            $FoundGroup = $false
                            foreach ($Group in $Rule.SendToManager.SecurityEscalation.ExcludeGroup) {
                                if ($User.MemberOf -contains $Group) {
                                    $FoundGroup = $true
                                    break
                                }
                            }
                            # if Group found, we need to skip this user
                            if ($FoundGroup) {
                                continue
                            }
                        }
                        if ($Rule.SendToManager.SecurityEscalation.IncludeGroup.Count -gt 0) {
                            # Rule defined that only user withi specific group has to be found
                            $FoundGroup = $false
                            foreach ($Group in $Rule.SendToManager.SecurityEscalation.IncludeGroup) {
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
                            SummaryDictionary = $Summary['NotifySecurity']
                            Type              = 'Security'
                            ManagerType       = 'Escalation'
                            Key               = $Rule.SendToManager.SecurityEscalation.Manager
                            User              = $User
                            Rule              = $Rule
                        }
                        Add-ManagerInformation @Splat
                    }
                }
            }
        } else {
            Write-Color -Text "[i]", " Processing rule ", $Rule.Name, ' status: ', $Rule.Enable -Color Red, White, Red, White, Red, White, Red, White
        }
    }

    if ($UserSection.Enable) {
        Write-Color -Text "[i] Sending notifications to users " -Color White, Yellow, White, Yellow, White, Yellow, White
        $CountUsers = 0
        [Array] $SummaryUsersEmails = foreach ($Notify in $Summary['Notify'].Values) {
            $CountUsers++
            $User = $Notify.User
            $Rule = $Notify.Rule

            # This shouldn't happen, but just in case, to be removed later on, as ProcessManagerOnly is skipping earlier on
            if ($Notify.ProcessManagersOnly -eq $true) {
                if ($Logging.NotifyOnSkipUserManagerOnly) {
                    Write-Color -Text "[i]", " Skipping User (Manager Only - $($Rule.Name)) ", $User.DisplayName, " (", $User.UserPrincipalName, ")", " days to expire: ", $User.DaysToExpire -Color Yellow, White, Magenta, White, Magenta, White, White, Blue
                }
                continue
            }

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
            } else {
                $EmailSplat.EmailParameters.To = $UserSection.DefaultEmail
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
            if ($Logging.NotifyOnUserSend) {
                Write-Color -Text "[i]", " Sending notifications to users ", $Notify.User.DisplayName, " (", $Notify.User.EmailAddress, ")", " status: ", $EmailResult.Status, " sent to: ", $EmailResult.SentTo, ", details: ", $EmailResult.Error -Color Yellow, White, Yellow, White, Yellow, White, White, Blue, White, Blue
            }
            if ($UserSection.SendCountMaximum -gt 0) {
                if ($UserSection.SendCountMaximum -le $CountUsers) {
                    Write-Color -Text "[i]", " Send count maximum reached. There may be more accounts that match the rule." -Color Red, DarkMagenta
                    break
                }
            }
        }
        Write-Color -Text "[i] Sending notifications to users (sent: ", $SummaryUsersEmails.Count, " out of ", $Summary['Notify'].Values.Count, ")" -Color White, Yellow, White, Yellow, White, Yellow, White
    } else {
        Write-Color -Text "[i] Sending notifications to users is ", "disabled!" -Color White, Yellow, DarkMagenta
    }
    if ($ManagerSection.Enable) {
        Write-Color -Text "[i] Sending notifications to managers " -Color White, Yellow, White, Yellow, White, Yellow, White
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
            $EmailSplat.EmailParameters = $EmailParameters

            if ($ManagerSection.SendToDefaultEmail -ne $true) {
                $EmailSplat.EmailParameters.To = $ManagerUser.EmailAddress
            } else {
                $EmailSplat.EmailParameters.To = $ManagerSection.DefaultEmail
            }
            if ($Logging.NotifyOnManagerSend) {
                Write-Color -Text "[i] Sending notifications to managers ", $ManagerUser.DisplayName, " (", $ManagerUser.EmailAddress, ") (SendToDefaultEmail: ", $ManagerSection.SendToDefaultEmail, ")" -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow
            }
            $EmailResult = Send-PasswordEmail @EmailSplat
            if ($Logging.NotifyOnManagerSend) {
                Write-Color -Text "[r] Sending notifications to managers ", $ManagerUser.DisplayName, " (", $ManagerUser.EmailAddress, ") (SendToDefaultEmail: ", $ManagerSection.SendToDefaultEmail, ") (status: ", $EmailResult.Status, " sent to: ", $EmailResult.SentTo, ")" -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow
            }

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
                    Write-Color -Text "[i]", " Send count maximum reached. There may be more managers that match the rule." -Color Red, DarkMagenta
                    break
                }
            }
        }
        Write-Color -Text "[i] Sending notifications to managers (sent: ", $SummaryManagersEmails.Count, " out of ", $Summary['NotifyManager'].Values.Count, ")" -Color White, Yellow, White, Yellow, White, Yellow, White
        #Write-Color -Text "[i] Sending notifications to managers (sent: ", $SummaryManagersEmails.Count, ")" -Color White, Yellow, White, Yellow, White, Yellow, White
    } else {
        Write-Color -Text "[i] Sending notifications to managers is ", "disabled!" -Color White, Yellow, DarkMagenta
    }
    if ($SecuritySection.Enable) {
        Write-Color -Text "[i] Sending notifications to security " -Color White, Yellow, White, Yellow, White, Yellow, White
        $CountSecurity = 0
        [Array] $SummaryEscalationEmails = foreach ($Manager in $Summary['NotifySecurity'].Keys) {
            $CountSecurity++
            # This user is provided by user in config file
            $ManagerUser = $Summary['NotifySecurity'][$Manager]['Manager']

            [Array] $ManagedUsers = $Summary['NotifySecurity'][$Manager]['Security'].Values.Output

            $EmailSplat = [ordered] @{}

            if ($Summary['NotifySecurity'][$Manager].Security.Count -gt 0) {
                # User uses global template
                $EmailSplat.Template = $TemplateSecurity
                if ($TemplateSecuritySubject) {
                    $EmailSplat.Subject = $TemplateSecuritySubject
                } else {
                    $EmailSplat.Subject = "[Password Expiring] Dear Security - Accounts expired"
                }
            } else {
                continue
            }

            $EmailSplat.User = $ManagerUser
            $EmailSplat.ManagedUsers = $ManagedUsers | Select-Object -Property 'Status', 'DisplayName', 'Enabled', 'SamAccountName', 'Domain', 'DateExpiry', 'DaysToExpire', 'PasswordLastSet', 'PasswordExpired'
            #$EmailSplat.ManagedUsersManagerNotCompliant = $ManagedUsersManagerNotCompliant
            #$EmailSplat.ManagedUsersManagerDisabled = $ManagedUsersManagerDisabled
            #$EmailSplat.ManagedUsersManagerMissing = $ManagedUsersManagerMissing
            #$EmailSplat.ManagedUsersManagerMissingEmail = $ManagedUsersManagerMissingEmail
            $EmailSplat.EmailParameters = $EmailParameters

            if ($ManagerSection.SendToDefaultEmail -ne $true) {
                $EmailSplat.EmailParameters.To = $ManagerUser.EmailAddress
            } else {
                $EmailSplat.EmailParameters.To = $SecuritySection.DefaultEmail
            }
            if ($Logging.NotifyOnSecuritySend) {
                Write-Color -Text "[i] Sending notifications to security ", $ManagerUser.DisplayName, " (", $ManagerUser.EmailAddress, ") (SendToDefaultEmail: ", $ManagerSection.SendToDefaultEmail, ")" -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow
            }
            $EmailResult = Send-PasswordEmail @EmailSplat
            if ($Logging.NotifyOnSecuritySend) {
                Write-Color -Text "[r] Sending notifications to security ", $ManagerUser.DisplayName, " (", $ManagerUser.EmailAddress, ") (SendToDefaultEmail: ", $ManagerSection.SendToDefaultEmail, ") (status: ", $EmailResult.Status, " sent to: ", $EmailResult.SentTo, ")" -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow
            }
            [PSCustomObject] @{
                DisplayName    = $ManagerUser.DisplayName
                SamAccountName = $ManagerUser.SamAccountName
                Domain         = $ManagerUser.Domain
                Status         = $EmailResult.Status
                StatusWhen     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                SentTo         = $EmailResult.SentTo
                StatusError    = $EmailResult.Error
                Accounts       = $ManagedUsers.SamAccountName
                AccountsCount  = $ManagedUsers.Count
                Template       = 'Unknown'
                # ManagerNotCompliant      = $ManagedUsersManagerNotCompliant.SamAccountName
                # ManagerNotCompliantCount = $ManagedUsersManagerNotCompliant.Count
                #ManagerDisabled          = $ManagedUsersManagerDisabled.SamAccountName
                #ManagerDisabledCount     = $ManagedUsersManagerDisabled.Count
                #ManagerMissing           = $ManagedUsersManagerMissing.SamAccountName
                #ManagerMissingCount      = $ManagedUsersManagerMissing.Count
                #ManagerMissingEmail      = $ManagedUsersManagerMissingEmail.SamAccountName
                #ManagerMissingEmailCount = $ManagedUsersManagerMissingEmail.Count
            }
            if ($ManagerSection.SendCountMaximum -gt 0) {
                if ($ManagerSection.SendCountMaximum -le $CountSecurity) {
                    Write-Color -Text "[i]", " Send count maximum reached. There may be more managers that match the rule." -Color Red, DarkMagenta
                    break
                }
            }
        }
        Write-Color -Text "[i] Sending notifications to security (sent: ", $SummaryEscalationEmails.Count, " out of ", $Summary['NotifySecurity'].Values.Count, ")" -Color White, Yellow, White, Yellow, White, Yellow, White
    } else {
        Write-Color -Text "[i] Sending notifications to security is ", "disabled!" -Color White, Yellow, DarkMagenta
    }

    if ($HTMLOptions.DisableWarnings -eq $true) {
        $WarningAction = 'SilentlyContinue'
    } else {
        $WarningAction = 'Continue'
    }

    $TimeEnd = Stop-TimeLog -Time $TimeStart -Option OneLiner

    if ($AdminSection.Enable) {
        Write-Color -Text "[i] Sending summary information " -Color White, Yellow, White, Yellow, White, Yellow, White
        $CountSecurity = 0
        [Array] $SummaryEmail = @(
            $CountSecurity++
            # This user is provided by user in config file
            $ManagerUser = $AdminSection.Manager

            $EmailSplat = [ordered] @{}
            # User uses global template
            $EmailSplat.Template = $TemplateAdmin
            $EmailSplat.Subject = $TemplateAdminSubject
            $EmailSplat.User = $ManagerUser

            $EmailSplat.SummaryUsersEmails = $SummaryUsersEmails
            $EmailSplat.SummaryManagersEmails = $SummaryManagersEmails
            $EmailSplat.SummaryEscalationEmails = $SummaryEscalationEmails
            $EmailSplat.TimeToProcess = $TimeEnd

            $EmailSplat.EmailParameters = $EmailParameters

            $EmailSplat.EmailParameters.To = $AdminSection.Manager.EmailAddress

            Write-Color -Text "[i] Sending summary information ", $ManagerUser.DisplayName, " (", $ManagerUser.EmailAddress, ") (SendToDefaultEmail: ", $ManagerSection.SendToDefaultEmail, ")" -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow

            $EmailResult = Send-PasswordEmail @EmailSplat

            Write-Color -Text "[r] Sending summary information ", $ManagerUser.DisplayName, " (", $ManagerUser.EmailAddress, ") (SendToDefaultEmail: ", $ManagerSection.SendToDefaultEmail, ") (status: ", $EmailResult.Status, " sent to: ", $EmailResult.SentTo, ")" -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow

            [PSCustomObject] @{
                DisplayName    = $ManagerUser.DisplayName
                SamAccountName = $ManagerUser.SamAccountName
                Domain         = $ManagerUser.Domain
                Status         = $EmailResult.Status
                StatusWhen     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                SentTo         = $EmailResult.SentTo
                StatusError    = $EmailResult.Error
                Template       = 'Unknown'
            }
        )
        Write-Color -Text "[i] Sending summary information (sent: ", $SummaryEmail.Count, ")" -Color White, Yellow, White, Yellow, White, Yellow, White
    } else {
        Write-Color -Text "[i] Sending summary information is ", "disabled!" -Color White, Yellow, DarkMagenta
    }

    Write-Color -Text "[i]", " Generating HTML report " -Color White, Yellow, Green
    # Create report
    New-HTML {
        New-HTMLHeader {
            New-HTMLSection -Invisible {
                New-HTMLSection {
                    New-HTMLText -Text "Report generated on $(Get-Date)" -Color Blue
                } -JustifyContent flex-start -Invisible
                New-HTMLSection {
                    New-HTMLText -Text "Password Solution - $($Script:Reporting['Version'])" -Color Blue
                } -JustifyContent flex-end -Invisible
            }
        }
        New-TableOption -DataStore JavaScript -ArrayJoin -BoolAsString
        if ($HTMLOptions.ShowAllUsers) {
            New-HTMLTab -Name 'All Users' {
                New-HTMLTable -DataTable $CachedUsers.Values -Filtering {
                    New-TableCondition -Name 'Enabled' -BackgroundColor LawnGreen -FailBackgroundColor BlueSmoke -Value $true -ComparisonType string -Operator eq
                    New-TableCondition -Name 'HasMailbox' -BackgroundColor LawnGreen -FailBackgroundColor BlueSmoke -Value $true -ComparisonType string -Operator eq
                    New-TableCondition -Name 'PasswordExpired' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $true -ComparisonType string
                    New-TableCondition -Name 'PasswordNeverExpires' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $false -ComparisonType string
                    New-TableCondition -Name 'ManagerStatus' -HighlightHeaders Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus -ComparisonType string -Value 'Missing', 'Disabled' -BackgroundColor Salmon -Operator in
                    New-TableCondition -Name 'ManagerStatus' -HighlightHeaders Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus -ComparisonType string -Value 'Enabled' -BackgroundColor LawnGreen
                    New-TableCondition -Name 'ManagerStatus' -HighlightHeaders Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus -ComparisonType string -Value 'Not available' -BackgroundColor BlueSmoke
                }
            }
        }
        if ($HTMLOptions.ShowRules) {
            foreach ($Rule in  $Summary['Rules'].Keys) {
                New-HTMLTab -Name $Rule {
                    New-HTMLTable -DataTable $Summary['Rules'][$Rule].Values.User -Filtering {
                        New-TableCondition -Name 'Enabled' -BackgroundColor LawnGreen -FailBackgroundColor BlueSmoke -Value $true -ComparisonType string
                        New-TableCondition -Name 'HasMailbox' -BackgroundColor LawnGreen -FailBackgroundColor BlueSmoke -Value $true -ComparisonType string -Operator eq
                        New-TableCondition -Name 'PasswordExpired' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $false -ComparisonType string
                        New-TableCondition -Name 'PasswordNeverExpires' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $false -ComparisonType string
                        New-TableCondition -Name 'ManagerStatus' -HighlightHeaders Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus -ComparisonType string -Value 'Missing', 'Disabled' -BackgroundColor Salmon -Operator in
                        New-TableCondition -Name 'ManagerStatus' -HighlightHeaders Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus -ComparisonType string -Value 'Enabled' -BackgroundColor LawnGreen
                        New-TableCondition -Name 'ManagerStatus' -HighlightHeaders Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus -ComparisonType string -Value 'Not available' -BackgroundColor BlueSmoke
                    }
                }
            }
        }
        if ($HTMLOptions.ShowUsersSent) {
            New-HTMLTab -Name 'Email sent to users' {
                New-HTMLTable -DataTable $SummaryUsersEmails {
                    New-TableHeader -Names 'Status', 'StatusError', 'SentTo', 'StatusWhen' -Title 'Email Summary'
                    New-TableCondition -Name 'Status' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $true -ComparisonType string -HighlightHeaders 'Status', 'StatusWhen', 'StatusError', 'SentTo'
                    New-TableCondition -Name 'PasswordExpired' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $false -ComparisonType string
                    New-TableCondition -Name 'PasswordNeverExpires' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $false -ComparisonType string
                } -Filtering
            }
        }
        if ($HTMLOptions.ShowManagersSent) {
            New-HTMLTab -Name 'Email sent to manager' {
                New-HTMLTable -DataTable $SummaryManagersEmails {
                    New-TableHeader -Names 'Status', 'StatusError', 'SentTo', 'StatusWhen' -Title 'Email Summary'
                    New-TableCondition -Name 'Status' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $true -ComparisonType string -HighlightHeaders 'Status', 'StatusWhen', 'StatusError', 'SentTo'
                } -Filtering
            }
        }
        if ($HTMLOptions.ShowEscalationSent) {
            New-HTMLTab -Name 'Email sent to Security' {
                New-HTMLTable -DataTable $SummaryEscalationEmails {
                    New-TableHeader -Names 'Status', 'StatusError', 'SentTo', 'StatusWhen' -Title 'Email Summary'
                    New-TableCondition -Name 'Status' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $true -ComparisonType string -HighlightHeaders 'Status', 'StatusWhen', 'StatusError', 'SentTo'
                } -Filtering
            }
        }
    } -ShowHTML:$HTMLOptions.ShowHTML -FilePath $FilePath -Online:$HTMLOptions.Online -WarningAction $WarningAction

    Write-Color -Text "[i]" , " Generating HTML report ", "Done" -Color White, Yellow, Green

    if ($SearchPath) {
        Write-Color -Text "[i]" , " Saving Search report " -Color White, Yellow, Green
        $SummarySearch['EmailSent'][$Today] = $SummaryUsersEmails
        $SummarySearch['EmailEscalations'][$Today] = $SummaryEscalationEmails
        $SummarySearch['EmailManagers'][$Today] = $SummaryManagersEmails

        try {
            $SummarySearch | Export-Clixml -LiteralPath $SearchPath
        } catch {
            Write-Color -Text "[e]", " Couldn't save to file $SearchPath", ". Error: ", $_.Exception.Message -Color White, Yellow, White, Yellow, White, Yellow, White
        }
        Write-Color -Text "[i]" , " Saving Search report ", "Done" -Color White, Yellow, Green
    }
}