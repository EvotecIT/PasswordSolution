Import-Module .\PasswordSolution.psd1 -Force

$Date = Get-Date

$GraphCredentials = @{
    ClientID     = '0fb383f1-8bfe-4c68-8ce2-5f6aa1d602fe'
    DirectoryID  = 'ceb371f6-8745-4876-a040-69f2d10a9d1a'
    ClientSecret = Get-Content -Raw -LiteralPath "C:\Support\Important\O365-GraphEmailTestingKey.txt"
}

$PasswordSolution = [ordered] @{
    # Graph based credentials
    EmailParameters                    = @{
        Credential = ConvertTo-GraphCredential -ClientID $GraphCredentials.ClientID -ClientSecret $GraphCredentials.ClientSecret -DirectoryID $GraphCredentials.DirectoryID
        Graph      = $true
        Priority   = 'Normal'
        From       = 'przemyslaw.klys+testgithub@evotec.pl'
        #To         = 'przemyslaw.klys+testgithub@evotec.pl' # your default email field (IMPORTANT)
        WhatIf     = $false
        ReplyTo    = 'contact+testgithub@evotec.pl'
    }
    # Standard SMTP credentials
    # EmailParameters                    = [ordered] @{
    #     UserName       = 'ADAutomations@evotec.pl'
    #     Password       = Get-Content -LiteralPath D:\Secrets\WO_SVC_ADAutomations.txt
    #     From           = 'ADAutomations@evotec.pl'
    #     Server         = 'smtp.office365.com'
    #     Priority       = 'High'
    #     UseSsl         = $true
    #     Port           = 587
    #     Verbose        = $false
    #     WhatIf         = $true
    #     AsSecureString = $true
    # }
    OverwriteEmailProperty             = 'extensionAttribute13'
    UserSection                        = @{
        Enable                 = $true
        SendCountMaximum       = 3
        SendToDefaultEmail     = $true # if enabled $EmailParameters are used (good for testing)
        DefaultEmail           = 'przemyslaw.klys+testgithub@evotec.pl' # your default email field (IMPORTANT)
    }
    ManagerSection                     = @{
        Enable                 = $true
        SendCountMaximum       = 3
        SendToDefaultEmail     = $true # if enabled $EmailParameters are used (good for testing)
        DefaultEmail           = 'przemyslaw.klys+testgithub@evotec.pl' # your default email field (IMPORTANT)
    }
    SecuritySection                    = @{
        Enable             = $true
        SendCountMaximum   = 3
        SendToDefaultEmail = $true # if enabled $EmailParameters are used (good for testing)
        DefaultEmail       = 'przemyslaw.klys+testgithub@evotec.pl' # your default email field (IMPORTANT)
        AttachCSV          = $true
    }
    AdminSection                       = @{
        Enable  = $true # doesn't processes this section at all
        Email   = 'przemyslaw.klys+testgithub@evotec.pl'
        Subject = "[Reporting Evotec] Summary of password reminders"
        Manager = [ordered] @{
            DisplayName  = 'Administrators'
            EmailAddress = 'przemyslaw.klys+testgithub@evotec.pl'
        }
    }
    Rules                              = @(
        # rules are new way to define things. You can define more than one rule and limit it per group/ou
        # the primary rule above can be set or doesn't have to, all parameters from rules below can be used across different rules
        # only one email will be sent even if the rules are overlapping, the first one wins
        #region "admins"
        [ordered] @{
            Name                        = 'Administrative Accounts'
            Enable                      = $false # doesn't processes this section at all if $false
            Reminders                   = -45, -30, -15, -7, 0, 1, 2, 3, 7, 15, 30, 60
            #Reminders                   = @(-200..-1), 0, 1, 2, 3, 4, 5, 12, 13, 14, 15, 28, 30, @(30..60), @(61..370)
            # this means we want to process only users that NeverExpire
            IncludeExpiring             = $true
            IncludePasswordNeverExpires = $true
            PasswordNeverExpiresDays    = 90
            IncludeNameProperties       = 'SamAccountName'
            IncludeName                 = @(
                "ADM_*"
                "SADM_*"
                "PADM_*"
                "MADM_*"
                "NADM_*"
                "ADM0_*"
                "ADM1_*"
                "ADM2_*"
            )
            SendToManager               = @{
                Manager             = [ordered] @{
                    Enable    = $true
                    Reminders = @{
                        OnDay = @{
                            Enable         = $true
                            Days           = 'Monday'
                            Reminder       = 10
                            ComparisonType = 'lt' # lt = less then, gt = greater then, eq = equal, in = inside
                        }
                    }
                }
                # Manager not compliant will be processed regardless of Reminder for Users
                ManagerNotCompliant = [ordered] @{
                    Enable        = $false
                    Manager       = [ordered] @{
                        DisplayName  = 'Global Service Desk'
                        EmailAddress = 'servicedesk@evotec.pl'
                    }
                    Disabled      = $true
                    Missing       = $true
                    MissingEmail  = $true
                    LastLogon     = $true
                    LastLogonDays = 90
                    Reminders     = @{
                        OnDayOfMonth = @{
                            Enable         = $true
                            Days           = 10, 21
                            Reminder       = 50
                            ComparisonType = 'lt' # lt = less then, gt = greater then, eq = equal, in = inside
                        }
                    }
                }
            }
        }
        #endregion admins
        #region "ITR01"
        [ordered] @{
            Name                        = 'ITR01 SVC'
            Enable                      = $false # doesn't processes this section at all if $false
            Reminders                   = -45, -30, -15, -7, 0, 1, 2, 3, 7, 15, 30, 60
            IncludeExpiring             = $true
            IncludePasswordNeverExpires = $true
            PasswordNeverExpiresDays    = 360
            IncludeNameProperties       = 'DisplayName', 'SamAccountName', 'Name', 'UserPrincipalName'
            IncludeName                 = @(
                "*SVC_*"
            )
            # limit group or limit OU can limit people with password never expire to certain users only
            IncludeOU                   = @(
                '*OU=ITR01,DC=*'
            )
            # It's important to use single quotes to not activate variables
            SendToManager               = @{
                Manager             = [ordered] @{
                    Enable    = $true
                    # it uses manager from AD in this section
                    Reminders = @{
                        Default = @{
                            Enable = $true
                        }
                        OnDay   = @{
                            Enable         = $true
                            Days           = 'Monday', 'Thursday'
                            Reminder       = 15
                            ComparisonType = 'lt' # lt = less then, gt = greater then, eq = equal, in = inside
                        }
                    }
                }
                # Security escalation will be processed regardless of Reminder for Users
                # Meaning Reminders definded below can be different then what users get
                SecurityEscalation  = [ordered] @{
                    Enable    = $true
                    Manager   = [ordered] @{
                        DisplayName  = 'IT Security'
                        EmailAddress = 'security@evotec.pl'
                    }
                    Reminders = @{
                        OnDayOfMonth = @{
                            Enable         = $true
                            Days           = 10, 21
                            Reminder       = -1
                            ComparisonType = 'lt' # lt = less then, gt = greater then, eq = equal, in = inside
                        }
                    }
                }
                # Manager not compliant will be processed regardless of Reminder for Users
                ManagerNotCompliant = [ordered] @{
                    Enable        = $true
                    Manager       = [ordered] @{
                        DisplayName  = 'ITR01 Service Desk'
                        EmailAddress = 'przemyslaw.klys+testgithub@evotec.pl'
                    }
                    Disabled      = $true
                    Missing       = $true
                    MissingEmail  = $true
                    LastLogon     = $true
                    LastLogonDays = 90
                    Reminders     = @{
                        OnDayOfMonth = @{
                            Enable         = $true
                            Days           = 10, 21
                            Reminder       = 50
                            ComparisonType = 'lt' # lt = less then, gt = greater then, eq = equal, in = inside
                        }
                    }
                }
            }
        }
        [ordered] @{
            Name                        = 'ITR01 USR'
            Enable                      = $false # doesn't processes this section at all if $false
            Reminders                   = -45, -30, -15, -7, 0, 1, 2, 3, 7, 15, 30, 60
            #Reminders                   = @(-200..-1), 0, 1, 2, 3, 4, 5, 12, 13, 14, 15, 28, 30, @(30..60), @(61..370)
            # this means we want to process only users that NeverExpire
            IncludeExpiring             = $true
            IncludePasswordNeverExpires = $true
            PasswordNeverExpiresDays    = 360
            IncludeNameProperties       = 'DisplayName', 'SamAccountName', 'Name', 'UserPrincipalName'
            IncludeName                 = @(
                "*_USR_*"
            )
            # limit group or limit OU can limit people with password never expire to certain users only
            IncludeOU                   = @(
                '*OU=ITR01,DC=*'
            )
            SendToManager               = @{
                Manager             = [ordered] @{
                    Enable    = $true
                    Reminders = @{
                        Default = @{
                            Enable = $true
                        }
                        OnDay   = @{
                            Enable         = $true
                            Days           = 'Monday', 'Thursday'
                            Reminder       = 15
                            ComparisonType = 'lt' # lt = less then, gt = greater then, eq = equal, in = inside
                        }
                    }

                }
                # Security escalation will be processed regardless of Reminder for Users
                # Meaning Reminders definded below can be different then what users get
                SecurityEscalation  = [ordered] @{
                    Enable    = $false
                    Manager   = [ordered] @{
                        DisplayName  = 'IT Security'
                        EmailAddress = 'security@evotec.pl'
                    }
                    Reminders = @{
                        OnDayOfMonth = @{
                            Enable         = $true
                            Days           = 10, 21
                            Reminder       = -1
                            ComparisonType = 'lt' # lt = less then, gt = greater then, eq = equal, in = inside
                        }
                    }
                }
                ManagerNotCompliant = [ordered] @{
                    Enable        = $false
                    Manager       = [ordered] @{
                        DisplayName  = 'ITR01 Service Desk'
                        EmailAddress = 'przemyslaw.klys+testgithub@evotec.pl'
                    }
                    Disabled      = $true
                    Missing       = $true
                    MissingEmail  = $true
                    LastLogon     = $true
                    LastLogonDays = 90
                    Reminders     = @{
                        OnDayOfMonth = @{
                            Enable         = $true
                            Days           = 10, 21
                            Reminder       = 50
                            ComparisonType = 'lt' # lt = less then, gt = greater then, eq = equal, in = inside
                        }
                    }
                }
            }
        }
        # [ordered] @{
        #     Name            = 'ITR01'
        #     Enable          = $true # doesn't processes this section at all if $false
        #     Reminders       = 60, 30, 15, 7, 3, 2, 1, 0, -7, -15, -30, -45
        #     # this means we want to process only users that NeverExpire
        #     IncludeExpiring = $true
        #     # limit group or limit OU can limit people with password never expire to certain users only
        #     IncludeOU       = @(
        #         '*OU=ITR01,DC=*'
        #     )
        # }

        [ordered] @{
            Name                  = 'ITR01'
            Enable                = $true # doesn't processes this section at all if $false
            Reminders             = 60, 30, 15, 7, 3, 2, 1, 0, -7, -15, -30, -45
            # this means we want to process only users that NeverExpire
            IncludeExpiring       = $true
            # limit group or limit OU can limit people with password never expire to certain users only
            IncludeOU             = @(
                '*OU=ITR01,DC=*'
            )
            IncludeNameProperties = 'SamAccountName'
            IncludeName           = @(
                "HACO"
            )
            SendToManager         = @{
                Manager = [ordered] @{
                    Enable    = $true
                    Reminders = @{
                        Default = @{
                            Enable = $true
                        }
                        OnDay   = @{
                            Enable         = $true
                            Days           = 'Monday', 'Thursday'
                            Reminder       = 15
                            ComparisonType = 'lt' # lt = less then, gt = greater then, eq = equal, in = inside
                        }
                    }

                }
            }
        }
        #endregion
        #region "All others"
        [ordered] @{
            Name            = 'All others'
            Enable          = $false # doesn't processes this section at all if $false
            Reminders       = @(500..-500), 60, 30, 15, 7, 3, 2, 1, 0, -7, -15, -30, -45
            IncludeExpiring = $true
            SendToManager   = @{
                Manager             = [ordered] @{
                    Enable    = $true
                    # it uses manager from AD in this section
                    Reminders = @{
                        Default = @{
                            Enable = $true
                        }
                        OnDay   = @{
                            Enable         = $true
                            Days           = 'Monday', 'Thursday'
                            Reminder       = 15
                            ComparisonType = 'lt' # lt = less then, gt = greater then, eq = equal, in = inside
                        }
                    }
                }
                # Security escalation will be processed regardless of Reminder for Users
                # Meaning Reminders definded below can be different then what users get
                SecurityEscalation  = [ordered] @{
                    Enable    = $true
                    Manager   = [ordered] @{
                        DisplayName  = 'IT Security'
                        EmailAddress = 'przemyslaw.klys+testgithub@evotec.pl'
                    }
                    Reminders = @{
                        OnDayOfMonth = @{
                            Enable         = $true
                            Days           = 1, 21
                            Reminder       = -1
                            ComparisonType = 'lt' # lt = less then, gt = greater then, eq = equal, in = inside
                        }
                    }
                }
                # Manager not compliant will be processed regardless of Reminder for Users
                ManagerNotCompliant = [ordered] @{
                    Enable        = $true
                    Manager       = [ordered] @{
                        DisplayName  = 'ITR01 Service Desk'
                        EmailAddress = 'przemyslaw.klys+testgithub@evotec.pl'
                    }
                    Disabled      = $true
                    Missing       = $true
                    MissingEmail  = $true
                    LastLogon     = $true
                    LastLogonDays = 90
                    Reminders     = @{
                        OnDayOfMonth = @{
                            Enable         = $true
                            Days           = 1, 21
                            Reminder       = 50
                            ComparisonType = 'lt' # lt = less then, gt = greater then, eq = equal, in = inside
                        }
                    }
                }
            }
        }
        #endregion "All others"
    )
    # Keep in mind those are script block not hashtable
    TemplatePreExpiry                  = {
        EmailImage -Source 'https://evotec.pl/wp-content/uploads/2015/05/Logo-evotec-012.png' -UrlLink '' -AlternativeText 'Evotec Logo' -Width '200' -Inline #-Height '100px'
        EmailText -LineBreak
        EmailText -Text "Dear ", "$DisplayName," -LineBreak
        EmailText -Text "Your password will expire in  $DaysToExpire days and if you do not change it, you will not be able to connect to the Evotec Network and IT services. "

        EmailText -Text "Depending on your situation, please follow one of the methods below to change your password." -LineBreak

        EmailText -Text "If you are connected to the Evotec Network (either directly or through VPN):"
        EmailList {
            EmailListItem -Text "Press CTRL+ALT+DEL"
            EmailListItem -Text "Choose Change password"
            EmailListItem -Text "Type in your old password and then type the new one according to the [KGD: 2-96-IS-POL-01113513.](http://search.evotec.local/Open/2-96-IS-POL-01113513)"
            EmailListItem -Text "After the change is complete you will be prompted with information that the password has been changed"
        }

        EmailText -Text "If you are not connected to the Evotec Network:"
        EmailList {
            EmailListItem -Text "Open [Password Change Link](https://account.activedirectory.windowsazure.com/ChangePassword.aspx) using your web browser"
            EmailListItem -Text "Login using your current credentials"
            EmailListItem -Text "On the change password form, type your old password and the new password that you want to set (twice)"
            EmailListItem -Text "Click Submit"
        }
        EmailText -Text "Please also remember to modify your password on the email configuration of your Smartphone or Tablet." -LineBreak
        EmailText -Text "Kind regards,"
        EmailText -Text "IT Service Desk"
    }
    TemplatePreExpirySubject           = '[Password Expiring] Your password will expire on $DateExpiry ($DaysToExpire days)'
    TemplatePostExpiry                 = {
        EmailImage -Source 'https://evotec.pl/wp-content/uploads/2015/05/Logo-evotec-012.png' -UrlLink '' -AlternativeText 'Evotec Logo' -Width '200' -Inline #-Height '100px'
        EmailText -LineBreak
        EmailText -Text "Dear ", "$DisplayName," -LineBreak
        EmailText -Text "Your password already expired on $PasswordLastSet. If you do not change it, you will not be able to connect to the Evotec Network and IT services. "

        EmailText -Text "Depending on your situation, please follow one of the methods below to change your password." -LineBreak

        EmailText -Text "If you are connected to the Evotec Network (either directly or through VPN):"
        EmailList {
            EmailListItem -Text "Press CTRL+ALT+DEL"
            EmailListItem -Text "Choose Change password"
            EmailListItem -Text "Type in your old password and then type the new one according to the [KGD: 2-96-IS-POL-01113513.](http://search.Evotec.local/Open/2-96-IS-POL-01113513)"
            EmailListItem -Text "After the change is complete you will be prompted with information that the password has been changed"
        }

        EmailText -Text "If you are not connected to the Evotec Network:"
        EmailList {
            EmailListItem -Text "Open [Password Change Link](https://account.activedirectory.windowsazure.com/ChangePassword.aspx) using your web browser"
            EmailListItem -Text "Login using your current credentials"
            EmailListItem -Text "On the change password form, type your old password and the new password that you want to set (twice)"
            EmailListItem -Text "Click Submit"
        }
        EmailText -Text "Please also remember to modify your password on the email configuration of your Smartphone or Tablet." -LineBreak
        EmailText -Text "Kind regards,"
        EmailText -Text "IT Service Desk"
    }
    TemplatePostExpirySubject          = '[Password Expiring] Your password expired on $DateExpiry ($DaysToExpire days ago)'
    TemplateManager                    = {
        EmailImage -Source 'https://evotec.pl/wp-content/uploads/2015/05/Logo-evotec-012.png' -UrlLink '' -AlternativeText 'Evotec Logo' -Width '200' -Inline #-Height '100px'
        EmailText -LineBreak
        EmailText -Text "Hello $ManagerDisplayName," -LineBreak

        EmailText -Text "Below is a summary of accounts where the password is due to expire soon. These accounts are either:"
        EmailList {
            EmailListItem -Text 'Managed by you'
            EmailListItem -Text 'You are the manager of the owner of these accounts.'
        }
        EmailText -Text "Where you are the owner, please action the password change on each account outlined below, according to the rules specified by Password Policy [KGD: 2-96-IS-POL-01113513](http://search.Evotec.local/Open/2-96-IS-POL-01113513)." -LineBreak

        EmailTable -DataTable $ManagerUsersTable -HideFooter

        EmailText -LineBreak
        EmailText -Text @(
            "Please note that for Service Accounts, even though the ",
            "'password never expires' "
            "flag remains set to "
            "'true' "
            ", the password MUST be changed before the expiry date specified in the above table. "
            "It is the responsibility of the manager of the account to ensure that this takes place. "
        ) -FontWeight normal, bold, normal, bold, normal, normal -LineBreak

        EmailText -Text @(
            "Please make an effort "
            "to change password yourself using known methods rather than asking the Service Desk to change the password for you. "
            "If password is changed by Service Desk agent, there are at least 2 people knowing the password - Service Desk Agent and You! "
            "Do you really want the Service Desk agent to know the password to critical system you manage/own? "
            "Be responsible!"
        ) -FontWeight bold, normal, normal, normal, bold -LineBreak -Color None, None, None, None, Red

        EmailText -Text "One of the ways to change the password is: " -FontWeight bold
        EmailList {
            EmailListItem -Text "Press CTRL+ALT+DEL"
            EmailListItem -Text "Choose Change password"
            EmailListItem -Text "In the account name - change it to the account you want to change password for." -FontWeight bold
            EmailListItem -Text "Type in current password for the account and then type the new one according to the rules specified in the password policy: [KGD: 2-96-IS-POL-01113513.](http://search.Evotec.local/Open/2-96-IS-POL-01113513)"
            EmailListItem -Text "After the change is complete you will be provided with information that the password has been changed"
        }
        EmailText -Text "Failure to take action could result in loss of service/escalation to the IT Security team." -LineBreak -FontWeight bold
        EmailText -Text "Kind regards,"
        EmailText -Text "IT Service Desk"
    }
    TemplateManagerSubject             = "[Passsword Expiring] Accounts you manage/own are expiring or already expired"
    TemplateSecurity                   = {
        EmailImage -Source 'https://evotec.pl/wp-content/uploads/2015/05/Logo-evotec-012.png' -UrlLink '' -AlternativeText 'Evotec Logo' -Width '200' -Inline #-Height '100px'
        EmailText -LineBreak
        EmailText -Text "Hello ", "$ManagerDisplayName", "," -LineBreak -FontWeight normal, bold, normal

        EmailText -Text "Below is a summary of ", "all service accounts", " where the passwords have exceeded the time limit stipulated in the password policy KGD. These accounts are all in violation of the KGD and immediate action/escalation should take place." -LineBreak -FontWeight normal, bold, normal

        EmailText -Text "It has been agreed that the ", "password never expires", " flag has been set to ", "true", " to avoid business disruption/loss of service. As a result we require your escalation to the managers of the account to take immediate action to change the password ASAP." -LineBreak -FontWeight normal, bold, normal, bold, normal
        EmailText -Text "Numerous automated reminders have been sent to the Manager, but no response/action has been taken yet." -LineBreak

        EmailText -Text "Please reach out directly to the manager/site to ensure that these passwords are changed immediately." -LineBreak

        EmailText -Text "If there is still lack of responses/action taken, it will be in your (IT Security) discretion to disable the account(s) question and take any appropriate action." -LineBreak -FontWeight bold

        EmailTable -DataTable $ManagerUsersTable -HideFooter

        EmailText -LineBreak
        EmailText -Text "Many thanks in advance." -LineBreak
        EmailText -Text "Kind regards,"
        EmailText -Text "IT Service Desk"
    }
    TemplateSecuritySubject            = "[Passsword Expired] Following accounts are expired!"
    TemplateManagerNotCompliant        = {
        EmailImage -Source 'https://evotec.pl/wp-content/uploads/2015/05/Logo-evotec-012.png' -UrlLink '' -AlternativeText 'Evotec Logo' -Width '200' -Inline #-Height '100px'

        EmailText -LineBreak
        EmailText -Text "Hello $ManagerDisplayName," -LineBreak

        EmailText -Text "Below is a summary of accounts where there is missing 'critical' information. These accounts are either:"

        EmailList {
            EmailListItem -Text "Missing a Manager in the AD - please add an active manager"
            EmailListItem -Text "The Manager in AD is Disabled - please add an active manager"
            EmailListItem -Text "Manager Last logon >90 days - please confirm if the manager is still an employee at Evotec/change the manager to an active manager"
            EmailListItem -Text "Manager is missing email - add manager email"
        }
        EmailText -Text "Please contact the respective local IT Service Desk (outlined in the below table) to update this Manager's attributes in the AD directly. The suggested action to take can be found in the below table." -LineBreak

        EmailTable -DataTable $ManagerUsersTableManagerNotCompliant -HideFooter

        EmailText -LineBreak
        EmailText -Text "Please note that all Service Accounts must have a Manager set in the AD, in order to fall within the Password Policy compliance notifications that are sent globally." -LineBreak

        EmailText -Text "Kind regards," -LineBreak
        EmailText -Text "IT Service Desk" -LineBreak
    }
    TemplateManagerNotCompliantSubject = "[Password Escalation] Accounts are expiring with non-compliant manager"
    TemplateAdmin                      = {
        EmailImage -Source 'https://evotec.pl/wp-content/uploads/2015/05/Logo-evotec-012.png' -UrlLink '' -AlternativeText 'Evotec Logo' -Width '200' -Inline #-Height '100px'

        EmailText -LineBreak
        EmailText -Text "Hello $ManagerDisplayName," -LineBreak

        EmailText -Text "Here's the summary of password notifications:"

        EmailList {
            EmailListItem -Text "Found users matching rule to send emails: ", $SummaryUsersEmails.Count
            EmailListItem -Text "Sent emails to users: ", ($SummaryUsersEmails | Where-Object { $_.Status -eq $true }).Count
            EmailListItem -Text "Couldn't send emails because of no email: ", ($SummaryUsersEmails | Where-Object { $_.Status -eq $false -and $_.StatusError -eq 'No email address for user' }).Count
            EmailListItem -Text "Couldn't send emails because other reasons: ", ($SummaryUsersEmails | Where-Object { $_.Status -eq $false -and $_.StatusError -ne 'No email address for user' }).Count
            EmailListItem -Text "Sent emails to managers: ", $SummaryManagersEmails.Count
            EmailListItem -Text "Sent emails to security: ", $SummaryEscalationEmails.Count
        }

        EmailText -Text "It took ", $TimeToProcess , " seconds to process the template." -LineBreak

        EmailText -Text "Hope everything works correctly! ", " You can take a look at [Password Solution Report](https://adcompliance.Evotec.local/CustomReports/PasswordSolution_$(($Date).ToString('yyyy-MM-dd_HH_mm_ss')).html) for details." -LineBreak

        EmailText -Text "Kind regards," -LineBreak
        EmailText -Text "IT Service Desk" -LineBreak
    }
    TemplateAdminSubject               = '[Password Summary] Passwords summary'
    Logging                            = @{
        # Logging to file and to screen
        ShowTime                     = $true
        LogFile                      = "$PSScriptRoot\Logs\PasswordSolution_$(($Date).ToString('yyyy-MM-dd_HH_mm_ss')).log"
        TimeFormat                   = "yyyy-MM-dd HH:mm:ss"
        LogMaximum                   = 365
        NotifyOnSkipUserManagerOnly  = $false
        NotifyOnSecuritySend         = $true
        NotifyOnManagerSend          = $true
        NotifyOnUserSend             = $true
        NotifyOnUserMatchingRule     = $true
        NotifyOnUserDaysToExpireNull = $true
        EmailDateFormat              = "yyyy-MM-dd HH:mm:ss"
        EmailDateFormatUTCConversion = $true
    }
    HTMLReports                        = @(
        # Accepts a list of reports to generate. Can be multiple reprorts having different sections, or just one having it all
        [ordered] @{
            Enable                = $true
            ShowHTML              = $true
            Title                 = "Password Solution Summary"
            Online                = $true
            DisableWarnings       = $true
            ShowConfiguration     = $true
            ShowAllUsers          = $true
            ShowRules             = $true
            ShowUsersSent         = $true
            ShowManagersSent      = $true
            ShowEscalationSent    = $true
            ShowSkippedUsers      = $true
            ShowSkippedLocations  = $true
            ShowSearchUsers       = $true
            ShowSearchManagers    = $true
            ShowSearchEscalations = $true
            FilePath              = "$PSScriptRoot\Reporting\PasswordSolution_$(($Date).ToString('yyyy-MM-dd_HH_mm_ss')).html"
            AttachToEmail         = $true
        }
    )
    SearchPath                         = "$PSScriptRoot\Search\SearchLog_$((Get-Date).ToString('yyyy-MM')).xml"
}

Start-PasswordSolution @PasswordSolution

