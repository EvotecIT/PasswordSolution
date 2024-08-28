Clear-Host
Import-Module .\PasswordSolution.psd1 -Force

$Date = Get-Date

Connect-MgGraph -Scopes "User.Read.All" -NoWelcome

Start-PasswordSolution {
    $Options = @{
        # Logging to file and to screen
        ShowTime                                          = $false
        LogFile                                           = "$PSScriptRoot\Logs\PasswordSolution_$(($Date).ToString('yyyy-MM-dd_HH_mm_ss')).log"
        TimeFormat                                        = "yyyy-MM-dd HH:mm:ss"
        LogMaximum                                        = 365
        NotifyOnSkipUserManagerOnly                       = $true
        NotifyOnSecuritySend                              = $true
        NotifyOnManagerSend                               = $true
        NotifyOnUserSend                                  = $true
        NotifyOnUserMatchingRule                          = $true
        NotifyOnUserDaysToExpireNull                      = $true
        NotifyOnUserMatchingRuleForManager                = $true
        NotifyOnUserMatchingRuleForManagerButNotCompliant = $true
        SearchPath                                        = "$PSScriptRoot\Search\SearchLog_$((Get-Date).ToString('yyyy-MM')).xml"
        EmailDateFormat                                   = "yyyy-MM-dd"
        EmailDateFormatUTCConversion                      = $true
        # FilterOrganizationalUnit                          = @(
        #     "*OU=Accounts,OU=Administration,DC=ad,DC=evotec,DC=xyz"
        #     "*OU=Administration,DC=ad,DC=evotec,DC=xyz"
        # )
    }
    New-PasswordConfigurationOption @Options

    New-PasswordConfigurationEntra -Enable

    $GraphCredentials = @{
        ClientID     = '0fb383f1-8bfe-4c68-8ce2-5f6aa1d602fe'
        DirectoryID  = 'ceb371f6-8745-4876-a040-69f2d10a9d1a'
        ClientSecret = Get-Content -Raw -LiteralPath "C:\Support\Important\O365-GraphEmailTestingKey.txt"
    }
    # (full support for Mailozaurr parameters)
    $EmailParameters = @{
        Credential = ConvertTo-GraphCredential -ClientID $GraphCredentials.ClientID -ClientSecret $GraphCredentials.ClientSecret -DirectoryID $GraphCredentials.DirectoryID
        Graph      = $true
        Priority   = 'Normal'
        From       = 'przemyslaw.klys@test.pl'
        WhatIf     = $true
        ReplyTo    = 'contact+testgithub@test.pl'
    }
    New-PasswordConfigurationEmail @EmailParameters

    # Configure behavior for different types of actions
    New-PasswordConfigurationType -Type User -Enable -SendCountMaximum 10 -DefaultEmail 'przemyslaw.klys+testgithub1@evotec.pl'
    New-PasswordConfigurationType -Type Manager -Enable -SendCountMaximum 10 -DefaultEmail 'przemyslaw.klys+testgithub2@evotec.pl'
    New-PasswordConfigurationType -Type Security -Enable -SendCountMaximum 1 -DefaultEmail 'przemyslaw.klys+testgithub3@evotec.pl' #-AttachCSV

    # Configure reporting
    $Report = [ordered] @{
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
        ShowSkippedUsers      = $false
        ShowSkippedLocations  = $false
        ShowSearchUsers       = $true
        ShowSearchManagers    = $true
        ShowSearchEscalations = $true
        NestedRules           = $false
        FilePath              = "$PSScriptRoot\Reporting\PasswordSolution_$(($Date).ToString('yyyy-MM-dd_HH_mm_ss')).html"
        AttachToEmail         = $true
    }
    New-PasswordConfigurationReport @Report

    # Configure rules for different types of users
    # New-PasswordConfigurationRule -Name 'Administrative Accounts' -Enable -IncludeExpiring -IncludePasswordNeverExpires -PasswordNeverExpiresDays 90 {
    #     # follow expiration days of a user
    #     New-PasswordConfigurationRuleReminder -Type 'Manager'
    #     # use a custom expiration days, and send only on specific days 1st, 10th and 15th of a month
    #     New-PasswordConfigurationRuleReminder -Type 'Manager' -DayOfMonth 1, 10, 15 -ExpirationDays -45, -30, -15, -7, 0, 1, 2, 3, 7, 15, 30, 60
    #     # use a custom expiration days (only if it's less then 10 days left), and send only on specific days of a week
    #     New-PasswordConfigurationRuleReminder -Type 'Manager' -DayOfWeek Monday, Wednesday, Friday -ExpirationDays 10 -ComparisonType 'lt'
    # } -ExpirationDays -45, -30, -15, -7, 0, 1, 2, 3, 7, 15, 30, 60 -IncludeNameProperties 'SamAccountName' -IncludeName = @(
    #     "ADM_*"
    #     "SADM_*"
    #     "PADM_*"
    #     "MADM_*"
    #     "NADM_*"
    #     "ADM0_*"
    #     "ADM1_*"
    #     "ADM2_*"
    # )

    # Configure rules for different types of users
    $newPasswordConfigurationRuleSplat = @{
        Name                             = 'Administrative Accounts'
        Enable                           = $true
        IncludeExpiring                  = $true
        IncludePasswordNeverExpires      = $true
        PasswordNeverExpiresDays         = 90
        ReminderConfiguration            = {
            # follow expiration days of a user
            New-PasswordConfigurationRuleReminder -Type 'Manager' -DayOfWeek Monday, Wednesday, Friday -ExpirationDays 60 -ComparisonType lt #-ExpirationDays @(-200..-1), 0, 1, 2, 3, 7, 15, 30, 60 -ComparisonType lt
            New-PasswordConfigurationRuleReminder -Type 'ManagerNotCompliant' -DayOfWeek Friday -ExpirationDays 300 -ComparisonType lt
            New-PasswordConfigurationRuleReminder -Type 'Security' -DayOfWeek Monday -ExpirationDays -1 -ComparisonType lt
        }
        #ReminderDays                     = '-45', '-30', '-15', '-7' #, 0, 1, 2, 3, 7, 15, 30, 60, 29, 28
        # IncludeOU                        = @(
        #     "*OU=Accounts,OU=Administration,DC=ad,DC=evotec,DC=xyz"
        #     "*OU=Administration,DC=ad,DC=evotec,DC=xyz"
        # )
        IncludeName                      = @(
            'Przem*'
        )
        IncludeNameProperties            = 'DisplayName', 'SamAccountName'
        ManagerReminder                  = $true
        ProcessManagersOnly              = $true
        ManagerNotCompliant              = $true
        ManagerNotCompliantDisplayName   = 'Global Service Desk'
        ManagerNotCompliantEmailAddress  = 'przemyslaw.klys@test.pl'
        ManagerNotCompliantDisabled      = $true
        ManagerNotCompliantMissing       = $true
        ManagerNotCompliantMissingEmail  = $true
        ManagerNotCompliantLastLogonDays = 90
        SecurityEscalation               = $true
        SecurityEscalationDisplayName    = 'IT Security'
        SecurityEscalationEmailAddress   = 'przemyslaw.klys@test.pl'
    }

    New-PasswordConfigurationRule @newPasswordConfigurationRuleSplat


    # New-PasswordConfigurationRule -Name 'All others' -Enable -ReminderDays @(500..-500), 60, 59, 30, 15, 7, 3, 2, 1, 0, -7, -15, -30, -45 {
    #     # follow expiration days of a user, you need to enable ManagerReminder for this functionality to work
    #     New-PasswordConfigurationRuleReminder -Type 'Manager' -ExpirationDays -45, -30, -15, -7, 0, 1, 2, 3, 7, 15, 30, 60
    #     # use a custom expiration days, and send only on specific days 1st, 10th and 15th of a month
    #     New-PasswordConfigurationRuleReminder -Type 'Manager' -DayOfMonth 1, 10, 15 -ExpirationDays -45, -30, -15, -7, 0, 1, 2, 3, 7, 15, 30, 60
    #     # use a custom expiration days (only if it's less then 10 days left), and send only on specific days of a week
    #     New-PasswordConfigurationRuleReminder -Type 'Manager' -DayOfWeek Monday, Wednesday, Friday -ExpirationDays 10 -ComparisonType 'lt'
    # } -IncludeExpiring -OverwriteEmailProperty 'extensionAttribute5' -OverwriteManagerProperty 'extensionAttribute1' -ManagerReminder

    # Template to user when sending email to user before password expires
    New-PasswordConfigurationTemplate -Type PreExpiry -Template {
        EmailImage -Source 'https://evotec.pl/wp-content/uploads/2015/05/Logo-evotec-012.png' -UrlLink '' -AlternativeText 'Evotec Logo' -Width '200' -Inline #-Height '100px'
        EmailText -LineBreak
        EmailText -Text "Dear ", "$DisplayName," -LineBreak
        EmailText -Text "Your password will expire in  $DaysToExpire days and if you do not change it, you will not be able to connect to the Evotec Network and IT services. "

        EmailText -Text "Depending on your situation, please follow one of the methods below to change your password." -LineBreak

        EmailText -Text "If you are connected to the Evotec Network (either directly or through VPN):"
        EmailList {
            EmailListItem -Text "Press CTRL+ALT+DEL"
            EmailListItem -Text "Choose Change password"
            EmailListItem -Text "Type in your old password and then type the new one according to the Policy (at least 8 characters, at least one uppercase letter, at least one lowercase letter, at least one number, at least one special character)"
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
    } -Subject '[Password Expiring] Your password will expire on $DateExpiry ($DaysToExpire days)'
    # Template to user when sending email to user after password expires
    New-PasswordConfigurationTemplate -Type PostExpiry -Template {
        EmailImage -Source 'https://evotec.pl/wp-content/uploads/2015/05/Logo-evotec-012.png' -UrlLink '' -AlternativeText 'Evotec Logo' -Width '200' -Inline #-Height '100px'
        EmailText -LineBreak
        EmailText -Text "Dear ", "$DisplayName," -LineBreak
        EmailText -Text "Your password already expired on $PasswordLastSet. If you do not change it, you will not be able to connect to the Evotec Network and IT services. "

        EmailText -Text "Depending on your situation, please follow one of the methods below to change your password." -LineBreak

        EmailText -Text "If you are connected to the Evotec Network (either directly or through VPN):"
        EmailList {
            EmailListItem -Text "Press CTRL+ALT+DEL"
            EmailListItem -Text "Choose Change password"
            EmailListItem -Text "Type in your old password and then type the new one according to the Policy"
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
    } -Subject '[Password Expiring] Your password expired on $DateExpiry ($DaysToExpire days ago)'
    # Template to security team with all service accounts that have expired passwords and password never expires set to true
    New-PasswordConfigurationTemplate -Type Security {
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
    } -Subject "[Passsword Expired] Following accounts are expired!"
    # Template to manager with all accounts that have expired passwords
    New-PasswordConfigurationTemplate -Type Manager {
        EmailImage -Source 'https://evotec.pl/wp-content/uploads/2015/05/Logo-evotec-012.png' -UrlLink '' -AlternativeText 'Evotec Logo' -Width '200' -Inline #-Height '100px'
        EmailText -LineBreak
        EmailText -Text "Hello $ManagerDisplayName," -LineBreak

        EmailText -Text "Below is a summary of accounts where the password is due to expire soon. These accounts are either:"
        EmailList {
            EmailListItem -Text 'Managed by you'
            EmailListItem -Text 'You are the manager of the owner of these accounts.'
        }
        EmailText -Text "Where you are the owner, please action the password change on each account outlined below, according to the rules specified by Password Policy." -LineBreak

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
            EmailListItem -Text "Type in current password for the account and then type the new one according to the rules specified in the password policy."
            EmailListItem -Text "After the change is complete you will be provided with information that the password has been changed"
        }
        EmailText -Text "Failure to take action could result in loss of service/escalation to the IT Security team." -LineBreak -FontWeight bold
        EmailText -Text "Kind regards,"
        EmailText -Text "IT Service Desk"
    } -Subject "[Passsword Expiring] Accounts you manage/own are expiring or already expired"
    # Template to Service Desk with information about manager missing, disabled, last logon >90 days, missing email for service accounts
    New-PasswordConfigurationTemplate -Type ManagerNotCompliant {
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
    } -Subject "[Password Escalation] Accounts are expiring with non-compliant manager"
    # Template to Admins with information summarizing what happened
    New-PasswordConfigurationTemplate -Type Admin {
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
    } -Subject '[Password Summary] Passwords summary'
}