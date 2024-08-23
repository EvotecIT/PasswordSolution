function Invoke-PasswordRuleProcessing {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Rule,
        [System.Collections.IDictionary] $Summary,
        [System.Collections.IDictionary] $CachedUsers,
        [System.Collections.IDictionary] $AllSkipped,
        [System.Collections.IDictionary] $Locations,
        [System.Collections.IDictionary] $Logging,
        [System.Collections.IDictionary] $UsersExternalSystem,
        [DateTime] $TodayDate
    )
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
                # Rule defined that only user within specific OU has to be found
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
                # Rule defined that only user within specific group has to be found
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
                # Rule defined that only user within specific group has to be found
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
            if ($Rule.ExcludeName.Count -gt 0) {
                $ExcludeName = $false
                foreach ($Name in $Rule.ExcludeName) {
                    foreach ($Property in $Rule.ExcludeNameProperties) {
                        if ($User.$Property -like $Name) {
                            $ExcludeName = $true
                            break
                        }
                    }
                    if ($ExcludeName) {
                        break
                    }
                }
                if ($ExcludeName) {
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

            if ($null -eq $User.DaysToExpire) {
                # This is to track users that our account may not have permissions over
                if ($Logging.NotifyOnUserDaysToExpireNull) {
                    Write-Color -Text @(
                        "[i]",
                        " User ",
                        $User.DisplayName,
                        " (",
                        $User.UserPrincipalName,
                        ")",
                        " days to expire not set. ",
                        "(",
                        "Password Last Set: ",
                        $User.PasswordLastSet,
                        ")",
                        " (Password at next logon: ",
                        $User.PasswordAtNextLogon, ")"
                    ) -Color Yellow, White, Yellow, White, Yellow, White, White, White, Yellow, DarkCyan, White, Yellow, DarkCyan, White
                }
                # if days to expire is not set, password last set is not set either
                # this means account either was never used or account we're using to has no permissions over that account
                $AllSkipped[$User.DistinguishedName] = $User

                $Location = $User.OrganizationalUnit
                if (-not $Location) {
                    $Location = 'Default'
                }
                if (-not $Locations[$Location]) {
                    $Locations[$Location] = [PSCustomObject] @{
                        Location     = $Location
                        Count        = 0
                        CountExpired = 0
                        Names        = [System.Collections.Generic.List[string]]::new()
                        NamesExpired = [System.Collections.Generic.List[string]]::new()
                    }
                }
                if ($User.PasswordExpired) {
                    $Locations[$Location].CountExpired++
                    $Locations[$Location].NamesExpired.Add($User.SamAccountName)
                } else {
                    $Locations[$Location].Count++
                    $Locations[$Location].Names.Add($User.SamAccountName)
                }
            }

            # Lets find users that expire, and match our rule
            if ($null -ne $User.DaysToExpire -and $User.DaysToExpire -in $Rule.Reminders) {
                # check if we need to notify user or just manager
                if (-not $Rule.ProcessManagersOnly) {
                    if ($Logging.NotifyOnUserMatchingRule) {
                        Write-Color -Text "[i]", " User ", $User.DisplayName, " (", $User.UserPrincipalName, ")", " days to expire: ", $User.DaysToExpire, " " -Color Yellow, White, Yellow, White, Yellow, White, White, Blue
                    }

                    # This is required for email notification to different email address
                    # User wanted to use different email address for notifications based on external property such as
                    # employeeID, employeeNumber, extensionAttributes, etc.
                    # normally this wouldn't be required if it's global setting, but if only a handful of users need to have their address changed
                    # the per rule overwriteemailproperty should be used
                    if ($Rule.OverwriteEmailProperty) {
                        $NewPropertyWithEmail = $Rule.OverwriteEmailProperty
                        if ($NewPropertyWithEmail -and $User.$NewPropertyWithEmail) {
                            $User.EmailAddress = $User.$NewPropertyWithEmail
                        }
                    }

                    if ($Rule.OverwriteEmailFromExternalUsers) {
                        $ExternalUser = $null
                        $ADProperty = $UsersExternalSystem.ActiveDirectoryProperty
                        $EmailProperty = $UsersExternalSystem.EmailProperty
                        $ExternalUser = $UsersExternalSystem['Users'][$User.$ADProperty]
                        if ($ExternalUser -and $ExternalUser.$EmailProperty -like '*@*') {
                            $User.EmailAddress = $ExternalUser.$EmailProperty
                        }
                    }

                    $Summary['Notify'][$User.DistinguishedName] = [ordered] @{
                        User                = $User
                        Rule                = $Rule
                        ProcessManagersOnly = $Rule.ProcessManagersOnly
                    }
                    # If we need to send an email to manager we need to update rules, just in case the user has not matched for user section
                    if ($Summary['Rules'][$Rule.Name][$User.DistinguishedName]) {
                        # User exists, update reason
                        $Summary['Rules'][$Rule.Name][$User.DistinguishedName].User.RuleOptions.Add('User')
                        $Summary['Rules'][$Rule.Name][$User.DistinguishedName].User.RuleName = $Rule.Name
                    } else {
                        # User doesn't exists in rules, add it
                        $Summary['Rules'][$Rule.Name][$User.DistinguishedName] = [ordered] @{
                            User                = $User
                            Rule                = $Rule
                            ProcessManagersOnly = $Rule.ProcessManagersOnly
                        }
                        $Summary['Rules'][$Rule.Name][$User.DistinguishedName].User.RuleOptions.Add('User')
                        $Summary['Rules'][$Rule.Name][$User.DistinguishedName].User.RuleName = $Rule.Name
                    }
                }
            }

            # this is to overwrite manager by using extensionAttribute or any other field in AD
            # it works on SamAccountName, DistinguishedName only
            if ($Rule.OverwriteManagerProperty) {
                $NewPropertyWithManager = $Rule.OverwriteManagerProperty
                if ($NewPropertyWithManager -and $User.$NewPropertyWithManager) {
                    $NewManager = $CachedUsers[$User.$NewPropertyWithManager]
                    if ($NewManager -and $NewManager.Mail -like "*@*") {
                        $User.ManagerEmail = $NewManager.Mail
                        $User.Manager = $NewManager.DisplayName
                        $User.ManagerSamAccountName = $NewManager.SamAccountName
                        $User.ManagerEnabled = $NewManager.Enabled
                        $User.ManagerLastLogon = $NewManager.LastLogonDate
                        if ($User.ManagerLastLogon) {
                            $User.ManagerLastLogonDays = $( - $($User.ManagerLastLogon - $Today).Days)
                        } else {
                            $User.ManagerLastLogonDays = $null
                        }
                        $User.ManagerType = $NewManager.ObjectClass
                        $User.ManagerDN = $NewManager.DistinguishedName
                    }
                }
            }

            # Lets find users that we need to notify manager about
            if ($null -ne $User.DaysToExpire -and $Rule.SendToManager) {
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
                        # Rule defined that only user within specific group has to be found
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
                        } elseif ($Rule.SendToManager.Manager.Reminders.Default.Enable -eq $true -and $Rule.SendToManager.Manager.Reminders.Default.Reminder -and $User.DaysToExpire -in $Rule.SendToManager.Manager.Reminders.Default.Reminder) {
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
                            if ($Logging.NotifyOnUserMatchingRuleForManager) {
                                Write-Color -Text "[i]", " User (manager rule) ", $User.DisplayName, " (", $User.UserPrincipalName, ")", " days to expire: ", $User.DaysToExpire, " " -Color Yellow, White, Yellow, White, Yellow, White, White, Blue
                            }
                            # If we need to send an email to manager we need to update rules, just in case the user has not matched for user section
                            if ($Summary['Rules'][$Rule.Name][$User.DistinguishedName]) {
                                # User exists, update reason
                                $Summary['Rules'][$Rule.Name][$User.DistinguishedName].User.RuleOptions.Add('Manager')
                                $Summary['Rules'][$Rule.Name][$User.DistinguishedName].User.RuleName = $Rule.Name
                            } else {
                                # User doesn't exists in rules, add it
                                $Summary['Rules'][$Rule.Name][$User.DistinguishedName] = [ordered] @{
                                    User                = $User
                                    Rule                = $Rule
                                    ProcessManagersOnly = $Rule.ProcessManagersOnly
                                }
                                $Summary['Rules'][$Rule.Name][$User.DistinguishedName].User.RuleOptions.Add('Manager')
                                $Summary['Rules'][$Rule.Name][$User.DistinguishedName].User.RuleName = $Rule.Name
                            }

                            # Push manager to list
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
                } else {
                    if ($Rule.SendToManager.Manager -and $Rule.SendToManager.Manager.Enable -eq $true) {
                        # Manager rule is enabled but manager is not enabled or has no email
                        if ($Logging.NotifyOnUserMatchingRuleForManagerButNotCompliant) {
                            Write-Color -Text "[i]", " User (manager rule) ", $User.DisplayName, " (", $User.UserPrincipalName, ")", " days to expire: ", $User.DaysToExpire, ", manager status: ", $User.ManagerStatus, ". Reason to skip: ", "No manager or manager is not enabled or manager has no email " -Color Yellow, White, Yellow, White, Yellow, White, White, Red, White, Red, White, Red
                        }
                    }
                }
            }
            # Lets find users that have no manager, manager is not enabled or manager has no email
            if ($Rule.SendToManager -and $Rule.SendToManager.ManagerNotCompliant -and $Rule.SendToManager.ManagerNotCompliant.Enable -eq $true -and $Rule.SendToManager.ManagerNotCompliant.Manager) {
                # Not compliant (missing, disabled, no email), covers all the below options
                if ($Rule.SendToManager.ManagerNotCompliant -and $Rule.SendToManager.ManagerNotCompliant.Enable -and $Rule.SendToManager.ManagerNotCompliant.Manager) {
                    $ManagerNotCompliant = $true
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
                            $ManagerNotCompliant = $false
                        }
                    }
                    if ($ManagerNotCompliant -and $Rule.SendToManager.ManagerNotCompliant.ExcludeOU.Count -gt 0) {
                        $FoundOU = $false
                        foreach ($OU in $Rule.SendToManager.ManagerNotCompliant.ExcludeOU) {
                            if ($User.OrganizationalUnit -like $OU) {
                                $FoundOU = $true
                                break
                            }
                        }
                        # if OU is found we need to exclude the user
                        if ($FoundOU) {
                            $ManagerNotCompliant = $false
                        }
                    }
                    if ($ManagerNotCompliant -and $Rule.SendToManager.ManagerNotCompliant.ExcludeGroup.Count -gt 0) {
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
                            $ManagerNotCompliant = $false
                        }
                    }
                    if ($ManagerNotCompliant -and $Rule.SendToManager.ManagerNotCompliant.IncludeGroup.Count -gt 0) {
                        # Rule defined that only user withi specific group has to be found
                        $FoundGroup = $false
                        foreach ($Group in $Rule.SendToManager.ManagerNotCompliant.IncludeGroup) {
                            if ($User.MemberOf -contains $Group) {
                                $FoundGroup = $true
                                break
                            }
                        }
                        if (-not $FoundGroup) {
                            $ManagerNotCompliant = $false
                        }
                    }

                    if ($Rule.SendToManager.ManagerNotCompliant.Reminders) {
                        $ManagerNotCompliant = $false
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
                        $ManagerNotCompliantMatched = $false
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

                            $ManagerNotCompliantMatched = $true
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

                            $ManagerNotCompliantMatched = $true
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

                            $ManagerNotCompliantMatched = $true
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

                            $ManagerNotCompliantMatched = $true
                        }

                        if ($ManagerNotCompliantMatched) {
                            if ($Logging.NotifyOnUserMatchingRuleForManagerNotCompliant) {
                                Write-Color -Text "[i]", " User (manager not compliant rule) ", $User.DisplayName, " (", $User.UserPrincipalName, ")", " days to expire: ", $User.DaysToExpire, " " -Color Yellow, White, Yellow, White, Yellow, White, White, Blue
                            }
                            # If we need to send an email to manager we need to update rules, just in case the user has not matched for user section
                            if ($Summary['Rules'][$Rule.Name][$User.DistinguishedName]) {
                                # User exists, update reason
                                $Summary['Rules'][$Rule.Name][$User.DistinguishedName].User.RuleOptions.Add('Manager Not Compliant')
                                $Summary['Rules'][$Rule.Name][$User.DistinguishedName].User.RuleName = $Rule.Name
                            } else {
                                # User doesn't exists in rules, add it
                                $Summary['Rules'][$Rule.Name][$User.DistinguishedName] = [ordered] @{
                                    User                = $User
                                    Rule                = $Rule
                                    ProcessManagersOnly = $Rule.ProcessManagersOnly
                                }
                                $Summary['Rules'][$Rule.Name][$User.DistinguishedName].User.RuleOptions.Add('Manager Not Compliant')
                                $Summary['Rules'][$Rule.Name][$User.DistinguishedName].User.RuleName = $Rule.Name
                            }
                        } else {
                            if ($User.ManagerStatus -eq 'Enabled') {
                                # do nothing
                            } else {
                                # This shouldn't happen, but just in case - we can log if this happens
                                if ($Logging.NotifyOnUserMatchingRuleForManagerNotCompliant) {
                                    Write-Color -Text "[i]", " User (manager not compliant rule not processed) ", $User.DisplayName, " (", $User.UserPrincipalName, ")", " days to expire: ", $User.DaysToExpire, " manager status: ", $User.ManagerStatus -Color Yellow, White, Yellow, White, Yellow, White, White, Blue
                                }
                            }
                        }
                    }
                }
            }
            # Lets find users that require escalation
            if ($null -ne $User.DaysToExpire -and $Rule.SendToManager -and $Rule.SendToManager.SecurityEscalation -and $Rule.SendToManager.SecurityEscalation.Enable -eq $true -and $Rule.SendToManager.SecurityEscalation.Manager) {
                $SecurityEscalation = $true
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
                        $SecurityEscalation = $false
                    }
                }
                if ($SecurityEscalation -and $Rule.SendToManager.SecurityEscalation.ExcludeOU.Count -gt 0) {
                    $FoundOU = $false
                    foreach ($OU in $Rule.SendToManager.SecurityEscalation.ExcludeOU) {
                        if ($User.OrganizationalUnit -like $OU) {
                            $FoundOU = $true
                            break
                        }
                    }
                    # if OU is found we need to exclude the user
                    if ($FoundOU) {
                        $SecurityEscalation = $false
                    }
                }
                if ($SecurityEscalation -and $Rule.SendToManager.SecurityEscalation.ExcludeGroup.Count -gt 0) {
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
                        $SecurityEscalation = $false
                    }
                }
                if ($SecurityEscalation -and $Rule.SendToManager.SecurityEscalation.IncludeGroup.Count -gt 0) {
                    # Rule defined that only user withi specific group has to be found
                    $FoundGroup = $false
                    foreach ($Group in $Rule.SendToManager.SecurityEscalation.IncludeGroup) {
                        if ($User.MemberOf -contains $Group) {
                            $FoundGroup = $true
                            break
                        }
                    }
                    if (-not $FoundGroup) {
                        $SecurityEscalation = $false
                    }
                }
                if ($Rule.SendToManager.SecurityEscalation.Reminders) {
                    $SecurityEscalation = $false
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
                    if ($Logging.NotifyOnUserMatchingRuleForSecurityEscalation) {
                        Write-Color -Text "[i]", " User (security escalation) ", $User.DisplayName, " (", $User.UserPrincipalName, ")", " days to expire: ", $User.DaysToExpire, " " -Color Yellow, White, Yellow, White, Yellow, White, White, Blue
                    }
                    # If we need to send an email to manager we need to update rules, just in case the user has not matched for user section
                    if ($Summary['Rules'][$Rule.Name][$User.DistinguishedName]) {
                        # User exists, update reason
                        $Summary['Rules'][$Rule.Name][$User.DistinguishedName].User.RuleOptions.Add('Security esclation')
                        $Summary['Rules'][$Rule.Name][$User.DistinguishedName].User.RuleName = $Rule.Name
                    } else {
                        # User doesn't exists in rules, add it
                        $Summary['Rules'][$Rule.Name][$User.DistinguishedName] = [ordered] @{
                            User                = $User
                            Rule                = $Rule
                            ProcessManagersOnly = $Rule.ProcessManagersOnly
                        }
                        $Summary['Rules'][$Rule.Name][$User.DistinguishedName].User.RuleOptions.Add('Security esclation')
                        $Summary['Rules'][$Rule.Name][$User.DistinguishedName].User.RuleName = $Rule.Name
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
        if ($null -ne $Rule.Name -and $null -ne $Rule.Enable) {
            Write-Color -Text "[i]", " Processing rule ", $Rule.Name, ' status: ', $Rule.Enable -Color Red, White, Red, White, Red, White, Red, White
        }
    }
}