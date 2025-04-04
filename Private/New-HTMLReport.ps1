function New-HTMLReport {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Report,
        [System.Collections.IDictionary] $EmailParameters,
        [System.Collections.IDictionary] $Logging,
        [string] $SearchPath,
        [Array] $Rules,
        [System.Collections.IDictionary] $UserSection,
        [System.Collections.IDictionary] $ManagerSection,
        [System.Collections.IDictionary] $SecuritySection,
        [System.Collections.IDictionary] $AdminSection,
        [System.Collections.IDictionary] $CachedUsers,
        [System.Collections.IDictionary] $Summary,
        [Array] $SummaryUsersEmails,
        [Array] $SummaryManagersEmails,
        [Array] $SummaryEscalationEmails,
        [System.Collections.IDictionary] $SummarySearch,
        [System.Collections.IDictionary] $Locations,
        [System.Collections.IDictionary] $AllSkipped,
        [System.Collections.IDictionary] $ExternalSystemReplacements,
        [ScriptBlock] $TemplateAdmin,
        [string] $TemplateAdminSubject,
        [Array] $FilterOrganizationalUnit,
        [Array] $SearchBase
    )
    $TranslateOperators = @{
        'lt' = 'Less than'
        'gt' = 'Greater than'
        'eq' = 'Equal to'
        'ne' = 'Not equal to'
        'le' = 'Less than or equal to'
        'ge' = 'Greater than or equal to'
        'in' = 'In'
    }

    Write-Color -Text "[i]", " Generating HTML report ", $Report.Title -Color White, Yellow, Green
    if ($Report.DisableWarnings -eq $true) {
        $WarningAction = 'SilentlyContinue'
    } else {
        $WarningAction = 'Continue'
    }
    if (-not $Report.Title) {
        $Report.Title = "Password Solution Report"
    }

    # Create report
    New-HTML {
        New-HTMLTabStyle -BorderRadius 0px -TextTransform capitalize -BackgroundColorActive SlateGrey
        New-HTMLSectionStyle -BorderRadius 0px -HeaderBackGroundColor Grey -RemoveShadow
        New-HTMLPanelStyle -BorderRadius 0px
        New-HTMLTableOption -DataStore JavaScript -BoolAsString -ArrayJoinString ', ' -ArrayJoin

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
        if ($Report.ShowConfiguration) {
            New-HTMLTab -Name "About" {
                New-HTMLTab -Name "Configuration" {
                    New-HTMLSection -Invisible {
                        New-HTMLSection -HeaderText "Email Configuration" {
                            New-HTMLList {
                                foreach ($Key in $EmailParameters.Keys) {
                                    if ($Key -eq 'Body') {

                                    } elseif ($Key -ne 'Password') {
                                        New-HTMLListItem -Text $Key, ": ", $EmailParameters[$Key] -FontWeight normal, normal, bold
                                    } else {
                                        New-HTMLListItem -Text $Key, ": ", "REDACTED" -FontWeight normal, normal, bold
                                    }
                                }
                            }
                        }
                        New-HTMLSection -HeaderText "Logging" {
                            New-HTMLList {
                                foreach ($Key in $Logging.Keys) {
                                    if ($Key -ne 'Password') {
                                        New-HTMLListItem -Text $Key, ": ", $Logging[$Key] -FontWeight normal, normal, bold
                                    } else {
                                        New-HTMLListItem -Text $Key, ": ", "REDACTED" -FontWeight normal, normal, bold
                                    }
                                }
                            }
                        }
                        New-HTMLSection -HeaderText "Other" {
                            New-HTMLList {
                                if ($Report.FilePath) {
                                    New-HTMLListItem -Text 'FilePath', ": ", $Report.FilePath -FontWeight normal, normal, bold
                                } else {
                                    New-HTMLListItem -Text 'FilePath', ": ", "Not set" -FontWeight normal, normal, bold
                                }
                                if ($Report.Email) {
                                    New-HTMLListItem -Text 'Email', ": ", $Report.Email -FontWeight normal, normal, bold
                                } else {
                                    New-HTMLListItem -Text 'Email', ": ", "Not set" -FontWeight normal, normal, bold
                                }
                                if ($SearchPath) {
                                    New-HTMLListItem -Text 'SearchPath', ": ", $SearchPath -FontWeight normal, normal, bold
                                } else {
                                    New-HTMLListItem -Text 'SearchPath', ": ", "Not set" -FontWeight normal, normal, bold
                                }
                                if ($FilterOrganizationalUnit.Count -gt 0) {
                                    New-HTMLListItem -Text 'FilterOrganizationalUnit', ": " {
                                        New-HTMLList {
                                            foreach ($OU in $FilterOrganizationalUnit) {
                                                New-HTMLListItem -Text 'OU', ": ", $OU -FontWeight normal, normal, bold
                                            }
                                        }
                                    } -FontWeight normal, normal, bold

                                } else {
                                    New-HTMLListItem -Text 'FilterOrganizationalUnit', ": ", "Not set" -FontWeight normal, normal, bold
                                }
                                if ($SearchBase.Count -gt 0) {
                                    New-HTMLListItem -Text 'SearchBase', ": " {
                                        New-HTMLList {
                                            foreach ($OU in $SearchBase) {
                                                New-HTMLListItem -Text 'OU', ": ", $OU -FontWeight normal, normal, bold
                                            }
                                        }
                                    } -FontWeight normal, normal, bold
                                } else {
                                    New-HTMLListItem -Text 'SearchBase', ": ", "Not set" -FontWeight normal, normal, bold
                                }
                            }
                        }
                    }

                    New-HTMLSection -Invisible {
                        New-HTMLSection -HeaderText "User Section" {
                            New-HTMLList {
                                New-HTMLListItem -Text "Enabled: ", $UserSection.Enable -FontWeight normal, bold -TextDecoration underline, none
                                New-HTMLListItem -Text "SendCountMaximum: ", $UserSection.SendCountMaximum -FontWeight normal, bold -TextDecoration underline, none
                                New-HTMLListItem -Text "SendToDefaultEmail: ", $UserSection.SendToDefaultEmail -FontWeight normal, bold -TextDecoration underline, none
                                New-HTMLListItem -Text "DefaultEmail: ", ($UserSection.DefaultEmail -join ", ") -FontWeight normal, bold -TextDecoration underline, none
                            }
                        }
                        New-HTMLSection -HeaderText "Manager Section" {
                            New-HTMLList {
                                New-HTMLListItem -Text "Enabled: ", $ManagerSection.Enable -FontWeight normal, bold -TextDecoration underline, none
                                New-HTMLListItem -Text "SendCountMaximum: ", $ManagerSection.SendCountMaximum -FontWeight normal, bold -TextDecoration underline, none
                                New-HTMLListItem -Text "SendToDefaultEmail: ", $ManagerSection.SendToDefaultEmail -FontWeight normal, bold -TextDecoration underline, none
                                New-HTMLListItem -Text "DefaultEmail: ", ($ManagerSection.DefaultEmail -join ", ") -FontWeight normal, bold -TextDecoration underline, none
                            }
                        }
                        New-HTMLSection -HeaderText "Security Section" {
                            New-HTMLList {
                                New-HTMLListItem -Text "Enabled: ", $SecuritySection.Enable -FontWeight normal, bold -TextDecoration underline, none
                                New-HTMLListItem -Text "SendCountMaximum: ", $SecuritySection.SendCountMaximum -FontWeight normal, bold -TextDecoration underline, none
                                New-HTMLListItem -Text "SendToDefaultEmail: ", $SecuritySection.SendToDefaultEmail -FontWeight normal, bold -TextDecoration underline, none
                                New-HTMLListItem -Text "DefaultEmail: ", ($SecuritySection.DefaultEmail -join ", ") -FontWeight normal, bold -TextDecoration underline, none
                                New-HTMLListItem -Text "Attach CSV: ", ($SecuritySection.AttachCSV -join ",") -FontWeight normal, bold -TextDecoration underline, none
                            }
                        }
                        New-HTMLSection -HeaderText "Admin Section" {
                            New-HTMLList {
                                New-HTMLListItem -Text "Enabled: ", $AdminSection.Enable -FontWeight normal, bold -TextDecoration underline, none
                                New-HTMLListItem -Text "Subject: ", $TemplateAdminSubject -FontWeight normal, bold -TextDecoration underline, none
                                New-HTMLListItem -Text "Manager: ", $AdminSection.Manager.DisplayName -FontWeight normal, bold -TextDecoration underline, none
                                New-HTMLListItem -Text "Manager Email: ", ($AdminSection.Manager.EmailAddress -join ", ") -FontWeight normal, bold -TextDecoration underline, none
                            }
                        }
                    }
                }
                New-HTMLTab -Name 'Rules Configuration' {
                    New-HTMLText -Text "There are ", $Rules.Count, " rules defined in the Password Solution. ", "Please keep in mind that order of the rules matter." -FontWeight normal, bold, normal -Color None, Blue, None

                    foreach ($Rule in $Rules) {
                        if ($Rule.Enable) {
                            $SectionColor = 'SpringGreen'
                        } else {
                            $SectionColor = 'Coral'
                        }
                        New-HTMLSection -HeaderText "Rule $($Rule.Name)" -CanCollapse -HeaderBackGroundColor $SectionColor {
                            New-HTMLList {
                                if ($Rule.Enable) {
                                    New-HTMLListItem -Text "Rule ", $Rule.Name, " is ", "enabled" -FontWeight normal, bold, normal, bold, normal, normal -Color None, None, None, Green
                                } else {
                                    New-HTMLListItem -Text "Rule ", $Rule.Name, " is ", "disabled" -FontWeight normal, bold, normal, bold, normal, normal -Color None, None, None, Red
                                }
                                New-HTMLList {
                                    New-HTMLListItem -Text "Notify till expiry on ", $($Rule.Reminders -join ","), " day " -FontWeight normal, bold, normal
                                    if ($Rule.IncludeExpiring) {
                                        New-HTMLListItem -Text "Include expiring accounts is ", "enabled" -FontWeight bold, bold -Color None, Green
                                    } else {
                                        New-HTMLListItem -Text "Include expiring accounts is ", "disabled" -FontWeight bold, bold -Color None, Red
                                    }
                                    if ($Rule.IncludePasswordNeverExpires) {
                                        New-HTMLListItem -Text "Include passwords never expiring with ", $Rule.PasswordNeverExpiresDays, " days rule" -FontWeight bold -Color Amethyst
                                    } else {
                                        New-HTMLListItem -Text "Do not include passwords that never expire." -FontWeight bold -Color Blue
                                    }
                                    if ($Rule.IncludeName.Count -gt 0 -and $Rule.IncludeNameProperties.Count -gt 0) {
                                        New-HTMLListItem -Text "Apply naming rule to require that account contains of of names ", $($Rule.IncludeName -join ", "), " in at least one property ", ($Rule.IncludeNameProperties -join ", ") -FontWeight normal, bold, normal, bold, normal -Color None, Blue, None, Blue
                                    } else {
                                        New-HTMLListItem -Text "Do not apply special name rules" -Color Blue -FontWeight bold
                                    }
                                    if ($Rule.IncludeOU) {
                                        New-HTMLListItem -Text "Apply Organizational Unit inclusion on ", ($Rule.IncludeOU -join ", ") -FontWeight normal, bold -Color None, Blue
                                    } else {
                                        New-HTMLListItem -Text "Do not apply Organizational Unit limit" -Color Blue -FontWeight bold
                                    }
                                    if ($Rule.ExcludeOU) {
                                        New-HTMLListItem -Text "Apply Organizational Unit exclusion on ", $Rule.ExcludeOU -FontWeight normal, bold -Color None, Green
                                    } else {
                                        New-HTMLListItem -Text "Do not exclude any Organizational Unit" -Color Blue -FontWeight bold
                                    }
                                    if ($Rule.IncludeGroup) {
                                        New-HTMLListItem -Text "Appply Group Membership inclusion (direct only) ", ($Rule.IncludeGroup -join ", ")
                                    } else {
                                        New-HTMLListItem -Text "Do not apply Group Membership limit"
                                    }
                                    if ($Rule.ExcludeGroup) {
                                        New-HTMLListItem -Text "Apply Group Membership exclusion (direct only): ", ($Rule.ExcludeGroup -join ", ")
                                    } else {
                                        New-HTMLListItem -Text "Do not apply Group Membership exclusion"
                                    }
                                    New-HTMLListItem -Text "Send to manager" -NestedListItems {
                                        New-HTMLList {
                                            if ($Rule.SendToManager.Manager.Enable) {
                                                New-HTMLListItem -Text "Manager ", " is ", 'enabled' -FontWeight bold, normal, bold -Color None, None, Green {
                                                    New-HTMLList {
                                                        New-HTMLListItem -Text "Rules: " {
                                                            New-HTMLList {
                                                                if ($Rule.SendToManager.Manager.Reminders.Default.Enable) {
                                                                    if ($Rule.SendToManager.Manager.Reminders.Default.Reminder) {
                                                                        New-HTMLListItem -Text "Default ", "is enabled", " sent on ", $($Rule.SendToManager.Manager.Reminders.Default.Reminder -join ", "), " days to expiry of user." -FontWeight normal, bold, normal, bold, normal -Color None, Green, None, Green
                                                                    } else {
                                                                        New-HTMLListItem -Text "Default ", "is enabled", " sent on ", $($Rule.Reminders -join ", "), " days to expiry of user." -FontWeight normal, bold, normal, bold, normal -Color None, Green, None, Green
                                                                    }
                                                                } else {
                                                                    New-HTMLListItem -Text "Default rule is ", "disabled" -FontWeight bold, bold -Color None, Red
                                                                }
                                                                if ($Rule.SendToManager.Manager.Reminders.OnDay.Enable) {
                                                                    New-HTMLListItem -Text @(
                                                                        "On day of the week ", "is ", "enabled"
                                                                        " on days: ", ($Rule.SendToManager.Manager.Reminders.OnDay.Days -join ", "),
                                                                        " with comparison ", $TranslateOperators[$Rule.SendToManager.Manager.Reminders.OnDay.ComparisonType],
                                                                        ' value ', $Rule.SendToManager.Manager.Reminders.OnDay.Reminder
                                                                    ) -FontWeight bold, normal, bold, normal, bold, normal, bold, normal, bold -Color None, None, Green, None, Green, None, Green, None, Green
                                                                } else {
                                                                    New-HTMLListItem -Text "On day of week rule is ", "disabled" -FontWeight bold, bold -Color None, Red
                                                                }
                                                                if ($Rule.SendToManager.Manager.Reminders.OnDayOfMonth.Enable) {
                                                                    New-HTMLListItem -Text @(
                                                                        "On day of the month rule ", "is", " enabled",
                                                                        " on days ", ($Rule.SendToManager.Manager.Reminders.OnDayOfMonth.Days -join ","),
                                                                        " with comparison ", $TranslateOperators[$Rule.SendToManager.Manager.Reminders.OnDayOfMonth.ComparisonType],
                                                                        ' value ', $Rule.SendToManager.Manager.Reminders.OnDayOfMonth.Reminder
                                                                    ) -FontWeight bold, normal, bold, normal, bold, normal, bold, normal, bold -Color None, None, Green, None, Green, None, Green, None, Green
                                                                } else {
                                                                    New-HTMLListItem -Text "On day of month rule is ", "disabled" -FontWeight bold, bold -Color None, Red
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            } else {
                                                New-HTMLListItem -Text "Manager ", " is ", 'disabled' -FontWeight bold, normal, bold -Color None, None, Red
                                            }
                                            if ($Rule.SendToManager.ManagerNotCompliant.Enable) {
                                                New-HTMLListItem -Text "Manager Escalation", " is ", 'enabled' -FontWeight bold, normal, bold -Color None, None, Green {
                                                    New-HTMLList {
                                                        New-HTMLListItem -Text "Manager Name: ", $Rule.SendToManager.ManagerNotCompliant.Manager.DisplayName -FontWeight normal, bold -TextDecoration underline, none
                                                        New-HTMLListItem -Text "Manager Email Address: ", $Rule.SendToManager.ManagerNotCompliant.Manager.EmailAddress -FontWeight normal, bold -TextDecoration underline, none
                                                    }
                                                    New-HTMLList {
                                                        New-HTMLListItem -Text "Rules: " {
                                                            New-HTMLList {
                                                                if ($Rule.SendToManager.ManagerNotCompliant.Reminders.Default.Enable) {
                                                                    if ($Rule.SendToManager.ManagerNotCompliant.Reminders.Default.Reminder) {
                                                                        New-HTMLListItem -Text "Default ", "is enabled", " sent on ", $($Rule.SendToManager.ManagerNotCompliant.Reminders.Default.Reminder -join ", "), " days to expiry of user." -FontWeight normal, bold, normal, bold, normal -Color None, Green, None, Green
                                                                    } else {
                                                                        New-HTMLListItem -Text "Default ", "is enabled", " sent on ", $($Rule.Reminders -join ", "), " days to expiry of user." -FontWeight normal, bold, normal, bold, normal -Color None, Green, None, Green
                                                                    }
                                                                } else {
                                                                    New-HTMLListItem -Text "Default rule is ", "disabled" -FontWeight bold, bold -Color None, Red
                                                                }
                                                                if ($Rule.SendToManager.ManagerNotCompliant.Reminders.OnDay.Enable) {
                                                                    New-HTMLListItem -Text @(
                                                                        "On day of the week ", "is ", "enabled"
                                                                        " on days: ", ($Rule.SendToManager.ManagerNotCompliant.Reminders.OnDay.Days -join ", "),
                                                                        " with comparison ", $TranslateOperators[$Rule.SendToManager.ManagerNotCompliant.Reminders.OnDay.ComparisonType],
                                                                        ' value ', $Rule.SendToManager.ManagerNotCompliant.Reminders.OnDay.Reminder
                                                                    ) -FontWeight bold, normal, bold, normal, bold, normal, bold, normal, bold -Color None, None, Green, None, Green, None, Green, None, Green
                                                                } else {
                                                                    New-HTMLListItem -Text "On day of week rule is ", "disabled" -FontWeight bold, bold -Color None, Red
                                                                }
                                                                if ($Rule.SendToManager.ManagerNotCompliant.Reminders.OnDayOfMonth.Enable) {
                                                                    New-HTMLListItem -Text @(
                                                                        "On day of the month rule ", "is", " enabled",
                                                                        " on days ", ($Rule.SendToManager.ManagerNotCompliant.Reminders.OnDayOfMonth.Days -join ", "),
                                                                        " with comparison ", $TranslateOperators[$Rule.SendToManager.ManagerNotCompliant.Reminders.OnDayOfMonth.ComparisonType],
                                                                        ' value ', $Rule.SendToManager.ManagerNotCompliant.Reminders.OnDayOfMonth.Reminder
                                                                    ) -FontWeight bold, normal, bold, normal, bold, normal, bold, normal, bold -Color None, None, Green, None, Green, None, Green, None, Green
                                                                } else {
                                                                    New-HTMLListItem -Text "On day of month rule is ", "disabled" -FontWeight bold, bold -Color None, Red
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            } else {
                                                New-HTMLListItem -Text "Manager Escalation", " is ", "disabled" -FontWeight bold, normal, bold -Color None, None, Red
                                            }
                                            if ($Rule.SendToManager.SecurityEscalation.Enable) {
                                                New-HTMLListItem -Text "Security Escalation ", "is", " enabled" -FontWeight bold, normal, bold -Color None, None, Green {
                                                    New-HTMLList {
                                                        New-HTMLListItem -Text "Manager Name: ", $Rule.SendToManager.SecurityEscalation.Manager.DisplayName -FontWeight normal, bold -TextDecoration underline, none
                                                        New-HTMLListItem -Text "Manager Email Address: ", $Rule.SendToManager.SecurityEscalation.Manager.EmailAddress -FontWeight normal, bold -TextDecoration underline, none
                                                    }
                                                    New-HTMLList {
                                                        New-HTMLListItem -Text "Rules: " {
                                                            New-HTMLList {
                                                                <#
                                                                        if ($Rule.SendToManager.SecurityEscalation.Reminders.Default.Enable) {
                                                                            New-HTMLListItem -Text "Default: ", $Rule.SendToManager.SecurityEscalation.Reminders.Default.Enable
                                                                        } else {
                                                                            New-HTMLListItem -Text "Default rule is ", "disabled" -FontWeight bold, bold -Color None, Red
                                                                        }
                                                                        #>
                                                                if ($Rule.SendToManager.SecurityEscalation.Reminders.Default.Enable) {
                                                                    if ($Rule.SendToManager.SecurityEscalation.Reminders.Default.Reminder) {
                                                                        New-HTMLListItem -Text "Default ", "is enabled", " sent on ", $($Rule.SendToManager.SecurityEscalation.Reminders.Default.Reminder -join ", "), " days to expiry of user." -FontWeight normal, bold, normal, bold, normal -Color None, Green, None, Green
                                                                    } else {
                                                                        New-HTMLListItem -Text "Default ", "is enabled", " sent on ", $($Rule.Reminders -join ", "), " days to expiry of user." -FontWeight normal, bold, normal, bold, normal -Color None, Green, None, Green
                                                                    }
                                                                } else {
                                                                    New-HTMLListItem -Text "Default rule is ", "disabled" -FontWeight bold, bold -Color None, Red
                                                                }
                                                                if ($Rule.SendToManager.SecurityEscalation.Reminders.OnDay.Enable) {
                                                                    New-HTMLListItem -Text @(
                                                                        "On day of the week ", "is ", "enabled"
                                                                        " on days: ", ($Rule.SendToManager.SecurityEscalation.Reminders.OnDay.Days -join ", "),
                                                                        " with comparison ", $TranslateOperators[$Rule.SendToManager.SecurityEscalation.Reminders.OnDay.ComparisonType],
                                                                        ' value ', $Rule.SendToManager.SecurityEscalation.Reminders.OnDay.Reminder
                                                                    ) -FontWeight bold, normal, bold, normal, bold, normal, bold, normal, bold -Color None, None, Green, None, Green, None, Green, None, Green
                                                                } else {
                                                                    New-HTMLListItem -Text "On day of week rule is ", "disabled" -FontWeight bold, bold -Color None, Red
                                                                }
                                                                if ($Rule.SendToManager.SecurityEscalation.Reminders.OnDayOfMonth.Enable) {
                                                                    New-HTMLListItem -Text @(
                                                                        "On day of the month rule ", "is", " enabled",
                                                                        " on days ", ($Rule.SendToManager.SecurityEscalation.Reminders.OnDayOfMonth.Days -join ", "),
                                                                        " with comparison ", $TranslateOperators[$Rule.SendToManager.SecurityEscalation.Reminders.OnDayOfMonth.ComparisonType],
                                                                        ' value ', $Rule.SendToManager.SecurityEscalation.Reminders.OnDayOfMonth.Reminder
                                                                    ) -FontWeight bold, normal, bold, normal, bold, normal, bold, normal, bold -Color None, None, Green, None, Green, None, Green, None, Green
                                                                } else {
                                                                    New-HTMLListItem -Text "On day of month rule is ", "disabled" -FontWeight bold, bold -Color None, Red
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            } else {
                                                New-HTMLListItem -Text "Security Escalation", " is ", "disabled" -FontWeight bold, normal, bold -Color None, None, Red
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        if ($Report.ShowAllUsers) {
            $AllUsers = foreach ($User in $CachedUsers.Values) {
                if ($User.Type -eq 'Contact') {
                    continue
                }
                $User
            }
            New-HTMLTab -Name 'All Users' {
                New-HTMLTable -DataTable $AllUsers -Filtering {
                    New-TableCondition -Name 'Enabled' -BackgroundColor LawnGreen -FailBackgroundColor BlueSmoke -Value $true -ComparisonType string -Operator eq
                    New-TableCondition -Name 'HasMailbox' -BackgroundColor LawnGreen -FailBackgroundColor BlueSmoke -Value $true -ComparisonType string -Operator eq
                    New-TableCondition -Name 'PasswordExpired' -BackgroundColor LawnGreen -Value $false -ComparisonType string
                    New-TableCondition -Name 'PasswordExpired' -BackgroundColor Salmon -Value $true -ComparisonType string
                    New-TableCondition -Name 'PasswordNeverExpires' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $false -ComparisonType string
                    New-TableCondition -Name 'PasswordAtNextLogon' -BackgroundColor BlueSmoke -Value $true -ComparisonType string
                    New-TableCondition -Name 'PasswordAtNextLogon' -BackgroundColor LawnGreen -Value $false -ComparisonType string
                    New-TableCondition -Name 'ManagerStatus' -HighlightHeaders Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus -ComparisonType string -Value 'Missing', 'Disabled' -BackgroundColor Salmon -Operator in
                    New-TableCondition -Name 'ManagerStatus' -HighlightHeaders Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus -ComparisonType string -Value 'Enabled' -BackgroundColor LawnGreen
                    New-TableCondition -Name 'ManagerStatus' -HighlightHeaders Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus -ComparisonType string -Value 'Not available' -BackgroundColor BlueSmoke
                } -ExcludeProperty $Report.ExcludeProperties -ScrollX
            }
        }
        if ($Report.ShowRules) {
            if ($Report.NestedRules) {
                # nested rules view under single tab
                if ($Summary['Rules'].Keys.Count -gt 0) {
                    New-HTMLTab -Name 'Rules Information' {
                        foreach ($Rule in  $Summary['Rules'].Keys) {
                            if ((Measure-Object -InputObject $Summary['Rules'][$Rule].Values.User).Count -gt 0) {
                                $Color = 'LawnGreen'
                                $IconSolid = 'Star'
                            } else {
                                $Color = 'Salmon'
                                $IconSolid = 'Stop'
                            }
                            New-HTMLTab -Name $Rule -TextColor $Color -IconColor $Color -IconSolid $IconSolid {
                                New-HTMLTable -DataTable $Summary['Rules'][$Rule].Values.User -Filtering {
                                    New-TableCondition -Name 'Enabled' -BackgroundColor LawnGreen -FailBackgroundColor BlueSmoke -Value $true -ComparisonType string
                                    New-TableCondition -Name 'HasMailbox' -BackgroundColor LawnGreen -FailBackgroundColor BlueSmoke -Value $true -ComparisonType string -Operator eq
                                    New-TableCondition -Name 'PasswordExpired' -BackgroundColor LawnGreen -Value $false -ComparisonType string
                                    New-TableCondition -Name 'PasswordExpired' -BackgroundColor Salmon -Value $true -ComparisonType string
                                    New-TableCondition -Name 'PasswordNeverExpires' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $false -ComparisonType string
                                    New-TableCondition -Name 'PasswordAtNextLogon' -BackgroundColor BlueSmoke -Value $true -ComparisonType string
                                    New-TableCondition -Name 'PasswordAtNextLogon' -BackgroundColor LawnGreen -Value $false -ComparisonType string
                                    New-TableCondition -Name 'ManagerStatus' -HighlightHeaders Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus -ComparisonType string -Value 'Missing', 'Disabled' -BackgroundColor Salmon -Operator in
                                    New-TableCondition -Name 'ManagerStatus' -HighlightHeaders Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus -ComparisonType string -Value 'Enabled' -BackgroundColor LawnGreen
                                    New-TableCondition -Name 'ManagerStatus' -HighlightHeaders Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus -ComparisonType string -Value 'Not available' -BackgroundColor BlueSmoke
                                } -ExcludeProperty $Report.ExcludeProperties -ScrollX
                            }
                        }
                    }
                }
            } else {
                foreach ($Rule in  $Summary['Rules'].Keys) {
                    if ((Measure-Object -InputObject $Summary['Rules'][$Rule].Values.User).Count -gt 0) {
                        $Color = 'LawnGreen'
                        $IconSolid = 'Star'
                    } else {
                        $Color = 'Salmon'
                        $IconSolid = 'Stop'
                    }
                    New-HTMLTab -Name $Rule -TextColor $Color -IconColor $Color -IconSolid $IconSolid {
                        New-HTMLTable -DataTable $Summary['Rules'][$Rule].Values.User -Filtering {
                            New-TableCondition -Name 'Enabled' -BackgroundColor LawnGreen -FailBackgroundColor BlueSmoke -Value $true -ComparisonType string
                            New-TableCondition -Name 'HasMailbox' -BackgroundColor LawnGreen -FailBackgroundColor BlueSmoke -Value $true -ComparisonType string -Operator eq
                            New-TableCondition -Name 'PasswordExpired' -BackgroundColor LawnGreen -Value $false -ComparisonType string
                            New-TableCondition -Name 'PasswordExpired' -BackgroundColor Salmon -Value $true -ComparisonType string
                            New-TableCondition -Name 'PasswordNeverExpires' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $false -ComparisonType string
                            New-TableCondition -Name 'PasswordAtNextLogon' -BackgroundColor BlueSmoke -Value $true -ComparisonType string
                            New-TableCondition -Name 'PasswordAtNextLogon' -BackgroundColor LawnGreen -Value $false -ComparisonType string
                            New-TableCondition -Name 'ManagerStatus' -HighlightHeaders Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus -ComparisonType string -Value 'Missing', 'Disabled' -BackgroundColor Salmon -Operator in
                            New-TableCondition -Name 'ManagerStatus' -HighlightHeaders Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus -ComparisonType string -Value 'Enabled' -BackgroundColor LawnGreen
                            New-TableCondition -Name 'ManagerStatus' -HighlightHeaders Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus -ComparisonType string -Value 'Not available' -BackgroundColor BlueSmoke
                        } -ExcludeProperty $Report.ExcludeProperties -ScrollX
                    }
                }
            }
        }
        if ($Report.ShowUsersSent) {
            if ((Measure-Object -InputObject $SummaryUsersEmails).Count -gt 0) {
                $Color = 'BrightTurquoise'
                $IconSolid = 'sticky-note'
            } else {
                $Color = 'Amaranth'
                $IconSolid = 'stop-circle'
            }
            New-HTMLTab -Name 'Email sent to users' -TextColor $Color -IconColor $Color -IconSolid $IconSolid {
                New-HTMLTable -DataTable $SummaryUsersEmails {
                    New-TableHeader -Names 'Status', 'StatusError', 'SentTo', 'StatusWhen' -Title 'Email Summary'
                    New-TableCondition -Name 'Status' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $true -ComparisonType string -HighlightHeaders 'Status', 'StatusWhen', 'StatusError', 'SentTo'
                    New-TableCondition -Name 'PasswordExpired' -BackgroundColor LawnGreen -Value $false -ComparisonType string
                    New-TableCondition -Name 'PasswordExpired' -BackgroundColor Salmon -Value $true -ComparisonType string
                    New-TableCondition -Name 'PasswordNeverExpires' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $false -ComparisonType string

                    New-TableCondition -Name 'Disabled' -BackgroundColor LawnGreen -Value $true -ComparisonType string -HighlightHeaders 'Disabled', 'DisabledError'
                    New-TableCondition -Name 'Disabled' -BackgroundColor Salmon -Value $false -ComparisonType string
                    New-TableCondition -Name 'DisabledError' -BackgroundColor ColumbiaBlue -Value 'WhatIf' -ComparisonType string -HighlightHeaders 'Disabled', 'DisabledError'
                } -Filtering -ScrollX
            }
        }
        if ($Report.ShowManagersSent) {
            if ((Measure-Object -InputObject $SummaryManagersEmails).Count -gt 0) {
                $Color = 'BrightTurquoise'
                $IconSolid = 'sticky-note'
            } else {
                $Color = 'Amaranth'
                $IconSolid = 'stop-circle'
            }
            New-HTMLTab -Name 'Email sent to manager' -TextColor $Color -IconColor $Color -IconSolid $IconSolid {
                New-HTMLTable -DataTable $SummaryManagersEmails {
                    New-TableHeader -Names 'Status', 'StatusError', 'SentTo', 'StatusWhen' -Title 'Email Summary'
                    New-TableCondition -Name 'Status' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $true -ComparisonType string -HighlightHeaders 'Status', 'StatusWhen', 'StatusError', 'SentTo'

                    New-TableCondition -Name 'DisabledAccountsError' -BackgroundColor LawnGreen -Value 'Not disabled', 'WhatIf', '' -ComparisonType string -Operator notin -HighlightHeaders 'DisabledAccounts', 'DisabledAccountsCount', 'DisabledAccountsError'
                } -Filtering -ScrollX
            }
        }
        if ($Report.ShowEscalationSent) {
            if ((Measure-Object -InputObject $SummaryEscalationEmails).Count -gt 0) {
                $Color = 'BrightTurquoise'
                $IconSolid = 'sticky-note'
            } else {
                $Color = 'Amaranth'
                $IconSolid = 'stop-circle'
            }
            New-HTMLTab -Name 'Email sent to Security' -TextColor $Color -IconColor $Color -IconSolid $IconSolid {
                New-HTMLTable -DataTable $SummaryEscalationEmails {
                    New-TableHeader -Names 'Status', 'StatusError', 'SentTo', 'StatusWhen' -Title 'Email Summary'
                    New-TableCondition -Name 'Status' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $true -ComparisonType string -HighlightHeaders 'Status', 'StatusWhen', 'StatusError', 'SentTo'
                } -Filtering -ScrollX
            }
        }
        if ($Report.ShowExternalSystemReplacementsUsers) {
            if ($ExternalSystemReplacements.Users.Count -gt 0) {
                $Color = 'BrightTurquoise'
                $IconSolid = 'sticky-note'
            } else {
                $Color = 'Amaranth'
                $IconSolid = 'stop-circle'
            }
            New-HTMLTab -Name 'External System Users' -TextColor $Color -IconColor $Color -IconSolid $IconSolid {
                New-HTMLTable -DataTable $ExternalSystemReplacements.Users {
                    #New-TableHeader -Names 'Status', 'StatusError', 'SentTo', 'StatusWhen' -Title 'Email Summary'
                    #New-TableCondition -Name 'Status' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $true -ComparisonType string -HighlightHeaders 'Status', 'StatusWhen', 'StatusError', 'SentTo'
                } -Filtering
            }
        }
        if ($Report.ShowExternalSystemReplacementsManagers) {
            if ($ExternalSystemReplacements.Managers.Count -gt 0) {
                $Color = 'BrightTurquoise'
                $IconSolid = 'sticky-note'
            } else {
                $Color = 'Amaranth'
                $IconSolid = 'stop-circle'
            }
            New-HTMLTab -Name 'External System Managers' -TextColor $Color -IconColor $Color -IconSolid $IconSolid {
                New-HTMLTable -DataTable $ExternalSystemReplacements.Managers {
                    #New-TableHeader -Names 'Status', 'StatusError', 'SentTo', 'StatusWhen' -Title 'Email Summary'
                    #New-TableCondition -Name 'Status' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $true -ComparisonType string -HighlightHeaders 'Status', 'StatusWhen', 'StatusError', 'SentTo'
                } -Filtering
            }
        }
        if ($Report.ShowSearchUsers) {
            [Array] $UsersSent = $SummarySearch['EmailSent'].Values #| ForEach-Object { if ($_ -ne $null) { $_ } }
            if ($UsersSent.Count -gt 0) {
                $Color = 'BrightTurquoise'
                $IconSolid = 'sticky-note'
            } else {
                $Color = 'Amaranth'
                $IconSolid = 'stop-circle'
            }
            New-HTMLTab -Name 'History Emails To Users' -TextColor $Color -IconColor $Color -IconSolid $IconSolid {
                New-HTMLTable -DataTable $UsersSent {
                    New-TableHeader -Names 'Status', 'StatusError', 'SentTo', 'StatusWhen' -Title 'Email Summary'
                    New-TableCondition -Name 'Status' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $true -ComparisonType string -HighlightHeaders 'Status', 'StatusWhen', 'StatusError', 'SentTo'
                    New-TableCondition -Name 'PasswordExpired' -BackgroundColor LawnGreen -Value $false -ComparisonType string
                    New-TableCondition -Name 'PasswordExpired' -BackgroundColor Salmon -Value $true -ComparisonType string
                    New-TableCondition -Name 'PasswordNeverExpires' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $false -ComparisonType string

                    New-TableCondition -Name 'Disabled' -BackgroundColor LawnGreen -Value $true -ComparisonType string -HighlightHeaders 'Disabled', 'DisabledError'
                    New-TableCondition -Name 'Disabled' -BackgroundColor Salmon -Value $false -ComparisonType string
                    New-TableCondition -Name 'DisabledError' -BackgroundColor ColumbiaBlue -Value 'WhatIf' -ComparisonType string -HighlightHeaders 'Disabled', 'DisabledError'
                } -Filtering -AllProperties -ScrollX
            }
        }
        if ($Report.ShowSearchManagers) {
            [Array] $ShowSearchManagers = $SummarySearch['EmailManagers'].Values #| ForEach-Object { if ($_ -ne $null) { $_ } }
            if ($ShowSearchManagers.Count -gt 0) {
                $Color = 'BrightTurquoise'
                $IconSolid = 'sticky-note'
            } else {
                $Color = 'Amaranth'
                $IconSolid = 'stop-circle'
            }
            New-HTMLTab -Name 'History Emails To Managers' -TextColor $Color -IconColor $Color -IconSolid $IconSolid {
                New-HTMLTable -DataTable $ShowSearchManagers {
                    New-TableHeader -Names 'Status', 'StatusError', 'SentTo', 'StatusWhen' -Title 'Email Summary'
                    New-TableCondition -Name 'Status' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $true -ComparisonType string -HighlightHeaders 'Status', 'StatusWhen', 'StatusError', 'SentTo'
                    New-TableCondition -Name 'DisabledAccountsError' -BackgroundColor LawnGreen -Value 'Not disabled', 'WhatIf', '' -ComparisonType string -Operator notin -HighlightHeaders 'DisabledAccounts', 'DisabledAccountsCount', 'DisabledAccountsError'
                } -Filtering -AllProperties -ScrollX
            }
        }
        if ($Report.ShowSearchEscalations) {
            [Array] $ShowSearchEscalations = $SummarySearch['EmailEscalations'].Values #| ForEach-Object { if ($_ -ne $null) { $_ } }
            if ($ShowSearchEscalations.Count -gt 0) {
                $Color = 'BrightTurquoise'
                $IconSolid = 'sticky-note'
            } else {
                $Color = 'Amaranth'
                $IconSolid = 'stop-circle'
            }
            New-HTMLTab -Name 'History Email To Security' -TextColor $Color -IconColor $Color -IconSolid $IconSolid {
                New-HTMLTable -DataTable $ShowSearchEscalations {
                    New-TableHeader -Names 'Status', 'StatusError', 'SentTo', 'StatusWhen' -Title 'Email Summary'
                    New-TableCondition -Name 'Status' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $true -ComparisonType string -HighlightHeaders 'Status', 'StatusWhen', 'StatusError', 'SentTo'
                } -Filtering -AllProperties -ScrollX
            }
        }
        if ($Report.ShowSkippedUsers) {
            New-HTMLTab -Name 'Skipped Users' -IconSolid users {
                $SkippedUsers = foreach ($User in  $AllSkipped.Values) {
                    if ($User.Type -ne 'Contact') {
                        $User
                    }
                }
                New-HTMLPanel -AlignContentText center {
                    New-HTMLText -FontSize 15pt -Text "Those users have no password date set. This means account running expiration checks doesn't have permissions or acccout never had password set or account is set to change password on logon. "
                } -Invisible
                New-HTMLTable -DataTable $SkippedUsers -Filtering {
                    New-TableCondition -Name 'Enabled' -BackgroundColor LawnGreen -FailBackgroundColor BlueSmoke -Value $true -ComparisonType string -Operator eq
                    New-TableCondition -Name 'HasMailbox' -BackgroundColor LawnGreen -FailBackgroundColor BlueSmoke -Value $true -ComparisonType string -Operator eq
                    New-TableCondition -Name 'PasswordExpired' -BackgroundColor LawnGreen -Value $false -ComparisonType string
                    New-TableCondition -Name 'PasswordExpired' -BackgroundColor Salmon -Value $true -ComparisonType string
                    New-TableCondition -Name 'PasswordNeverExpires' -BackgroundColor LawnGreen -FailBackgroundColor Salmon -Value $false -ComparisonType string
                    New-TableCondition -Name 'PasswordAtNextLogon' -BackgroundColor BlueSmoke -Value $true -ComparisonType string
                    New-TableCondition -Name 'PasswordAtNextLogon' -BackgroundColor LawnGreen -Value $false -ComparisonType string
                    New-TableCondition -Name 'ManagerStatus' -HighlightHeaders Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus -ComparisonType string -Value 'Missing', 'Disabled' -BackgroundColor Salmon -Operator in
                    New-TableCondition -Name 'ManagerStatus' -HighlightHeaders Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus -ComparisonType string -Value 'Enabled' -BackgroundColor LawnGreen
                    New-TableCondition -Name 'ManagerStatus' -HighlightHeaders Manager, ManagerSamAccountName, ManagerEmail, ManagerStatus -ComparisonType string -Value 'Not available' -BackgroundColor BlueSmoke
                } -ScrollX
            }
        }
        if ($Report.ShowSkippedLocations) {
            New-HTMLTab -Name 'Skipped Locations' -IconSolid building {
                New-HTMLPanel -AlignContentText center {
                    New-HTMLText -FontSize 15pt -Text "Users in those Organizational Units have no password date set. This means account running expiration checks doesn't have permissions or acccout never had password set or account is set to change password on logon. "
                } -Invisible
                New-HTMLTable -DataTable $Locations.Values -Filtering {
                    New-TableHeader -ResponsiveOperations none -Names 'Names', 'NamesExpired'
                } -ScrollX
            }
        }
    } -ShowHTML:$Report.ShowHTML -FilePath $Report.FilePath -Online:$Report.Online -WarningAction $WarningAction -TitleText $Report.Title

    Write-Color -Text "[i]" , " Generating HTML report ", $Report.Title, ". Done" -Color White, Yellow, Green
}