function Set-PasswordConfiguration {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Logging,
        [scriptblock] $ConfigurationDSL,
        [scriptblock] $TemplatePreExpiry,
        [string] $TemplatePreExpirySubject,
        [scriptblock] $TemplatePostExpiry,
        [string] $TemplatePostExpirySubject,
        [scriptblock] $TemplateManager,
        [string] $TemplateManagerSubject,
        [scriptblock] $TemplateSecurity,
        [string] $TemplateSecuritySubject,
        [scriptblock] $TemplateManagerNotCompliant,
        [string] $TemplateManagerNotCompliantSubject,
        [scriptblock] $TemplateAdmin,
        [string] $TemplateAdminSubject,
        [System.Collections.IDictionary] $EmailParameters,
        [System.Collections.IDictionary] $UserSection,
        [System.Collections.IDictionary] $ManagerSection,
        [System.Collections.IDictionary] $SecuritySection,
        [System.Collections.IDictionary] $AdminSection,
        [System.Collections.IDictionary] $UsersExternalSystem,
        [Array] $HTMLReports,
        [Array] $Rules,
        [string] $SearchPath,
        [string] $OverwriteEmailProperty,
        [string] $OverwriteManagerProperty,
        [string[]] $FilterOrganizationalUnit
    )

    if (-not $Rules) {
        $Rules = @() # not worth the effort for generic list
    }
    if (-not $HTMLReports) {
        $HTMLReports = @() # not worth the effort for generic list
    }

    if ($ConfigurationDSL) {
        try {
            $ConfigurationExecuted = & $ConfigurationDSL
            foreach ($Configuration in $ConfigurationExecuted) {
                if ($Configuration.Type -eq 'PasswordConfigurationOption') {
                    if ($Configuration.Settings.SearchPath) {
                        $SearchPath = $Configuration.Settings.SearchPath
                    }
                    if ($Configuration.Settings.OverwriteEmailProperty) {
                        $OverwriteEmailProperty = $Configuration.Settings.OverwriteEmailProperty
                    }
                    if ($Configuration.Settings.OverwriteManagerProperty) {
                        $OverwriteManagerProperty = $Configuration.Settings.OverwriteManagerProperty
                    }
                    if ($Configuration.Settings.FilterOrganizationalUnit) {
                        $FilterOrganizationalUnit = $Configuration.Settings.FilterOrganizationalUnit
                    }
                    foreach ($Setting in $Configuration.Settings.Keys) {
                        if ($Setting -notin 'SearchPath', 'OverwriteEmailProperty', 'OverwriteManagerProperty', 'FilterOrganizationalUnit') {
                            $Logging[$Setting] = $Configuration.Settings[$Setting]
                        }
                    }
                } elseif ($Configuration.Type -eq 'PasswordConfigurationEmail') {
                    $EmailParameters = $Configuration.Settings
                } elseif ($Configuration.Type -eq 'PasswordConfigurationTypeUser') {
                    $UserSection = $Configuration.Settings
                } elseif ($Configuration.Type -eq 'PasswordConfigurationTypeManager') {
                    $ManagerSection = $Configuration.Settings
                } elseif ($Configuration.Type -eq 'PasswordConfigurationTypeSecurity') {
                    $SecuritySection = $Configuration.Settings
                } elseif ($Configuration.Type -eq 'PasswordConfigurationTypeAdmin') {
                    $AdminSection = $Configuration.Settings
                } elseif ($Configuration.Type -eq 'PasswordConfigurationReport') {
                    $HTMLReports += $Configuration.Settings
                } elseif ($Configuration.Type -eq 'PasswordConfigurationRule') {
                    if ($Configuration.Error) {
                        return
                    }
                    $Rules += $Configuration.Settings
                } elseif ($Configuration.Type -eq "PasswordConfigurationTemplatePreExpiry") {
                    $TemplatePreExpiry = $Configuration.Settings.Template
                    $TemplatePreExpirySubject = $Configuration.Settings.Subject
                } elseif ($Configuration.Type -eq "PasswordConfigurationTemplatePostExpiry") {
                    $TemplatePostExpiry = $Configuration.Settings.Template
                    $TemplatePostExpirySubject = $Configuration.Settings.Subject
                } elseif ($Configuration.Type -eq "PasswordConfigurationTemplateManager") {
                    $TemplateManager = $Configuration.Settings.Template
                    $TemplateManagerSubject = $Configuration.Settings.Subject
                } elseif ($Configuration.Type -eq "PasswordConfigurationTemplateSecurity") {
                    $TemplateSecurity = $Configuration.Settings.Template
                    $TemplateSecuritySubject = $Configuration.Settings.Subject
                } elseif ($Configuration.Type -eq "PasswordConfigurationTemplateManagerNotCompliant") {
                    $TemplateManagerNotCompliant = $Configuration.Settings.Template
                    $TemplateManagerNotCompliantSubject = $Configuration.Settings.Subject
                } elseif ($Configuration.Type -eq "PasswordConfigurationTemplateAdmin") {
                    $TemplateAdmin = $Configuration.Settings.Template
                    $TemplateAdminSubject = $Configuration.Settings.Subject
                } elseif ($Configuration.Type -eq 'ExternalUsers') {
                    $UsersExternalSystem = $Configuration
                }
            }
        } catch {
            Write-Color -Text "[e]", " Processing configuration failed because of error in line ", $_.InvocationInfo.ScriptLineNumber, " in ", $_.InvocationInfo.InvocationName, " with message: ", $_.Exception.Message -Color Yellow, White, Red
            return
        }
    }

    if (-not $TemplatePreExpiry) {
        Write-Color -Text "[i]", " TemplatePreExpiry not defined. Using default template (built-in)" -Color Yellow, Red
        $TemplatePreExpiry = {
            EmailText -LineBreak
            EmailText -Text "Dear ", "$DisplayName," -LineBreak
            EmailText -Text "Your password will expire in  $DaysToExpire days and if you do not change it, you will not be able to connect to the Network and IT services. "

            EmailText -Text "Depending on your situation, please follow one of the methods below to change your password." -LineBreak

            EmailText -Text "If you are connected to the Internal Network (either directly or through VPN):"
            EmailList {
                EmailListItem -Text "Press CTRL+ALT+DEL"
                EmailListItem -Text "Choose Change password"
                EmailListItem -Text "Type in your old password and then type the new one according to the password policy (twice)"
                EmailListItem -Text "After the change is complete you will be prompted with information that the password has been changed"
            }

            EmailText -Text "If you are not connected to the Internal Network:"
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
    }
    if (-not $TemplatePreExpirySubject) {
        Write-Color -Text "[i]", " TemplatePreExpirySubject not defined. Using default template (built-in)" -Color Yellow, Red
        $TemplatePreExpirySubject = '[Password Expiring] Your password will expire on $DateExpiry ($DaysToExpire days)'
    }

    if (-not $TemplatePostExpiry) {
        Write-Color -Text "[i]", " TemplatePostExpiry not defined. Using default template (built-in)" -Color Yellow, Red
        # Template to user when sending email to user after password expires
        $TemplatePostExpiry = {
            EmailText -LineBreak
            EmailText -Text "Dear ", "$DisplayName," -LineBreak
            EmailText -Text "Your password already expired on $PasswordLastSet. If you do not change it, you will not be able to connect to the Network and IT services. "

            EmailText -Text "Depending on your situation, please follow one of the methods below to change your password." -LineBreak

            EmailText -Text "If you are connected to the Network (either directly or through VPN):"
            EmailList {
                EmailListItem -Text "Press CTRL+ALT+DEL"
                EmailListItem -Text "Choose Change password"
                EmailListItem -Text "Type in your old password and then type the new one according to the password policy (twice)"
                EmailListItem -Text "After the change is complete you will be prompted with information that the password has been changed"
            }

            EmailText -Text "If you are not connected to the Internal Network:"
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
    }
    if (-not $TemplatePostExpirySubject) {
        Write-Color -Text "[i]", " TemplatePostExpirySubject not defined. Using default template (built-in)" -Color Yellow, Red
        $TemplatePostExpirySubject = '[Password Expired] Your password expired on $DateExpiry ($DaysToExpire days ago)'
    }
    # Template to security team with all service accounts that have expired passwords and password never expires set to true
    if (-not $TemplateSecurity) {
        Write-Color -Text "[i]", " TemplateSecurity not defined. Using default template (built-in)" -Color Yellow, Red
        $TemplateSecurity = {
            EmailText -LineBreak
            EmailText -Text "Hello ", "$ManagerDisplayName", "," -LineBreak -FontWeight normal, bold, normal

            EmailText -Text @(
                "Below is a summary of ", "all service accounts",
                " where the passwords have exceeded the time limit stipulated in the password policy. These accounts are all in violation of the policy and immediate action/escalation should take place."
            ) -LineBreak -FontWeight normal, bold, normal

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
    }
    if (-not $TemplateSecuritySubject) {
        Write-Color -Text "[i]", " TemplateSecuritySubject not defined. Using default template (built-in)" -Color Yellow, Red
        $TemplateSecuritySubject = "[Passsword Expired] Following accounts are expired!"
    }
    if (-not $TemplateManager) {
        Write-Color -Text "[i]", " TemplateManager not defined. Using default template (built-in)" -Color Yellow, Red

        $TemplateManager = {
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
        }
    }
    if (-not $TemplateManagerSubject) {
        Write-Color -Text "[i]", " TemplateManagerSubject not defined. Using default template (built-in)" -Color Yellow, Red
        $TemplateManagerSubject = "[Passsword Expiring] Accounts you manage/own are expiring or already expired"
    }

    # Template to Service Desk with information about manager missing, disabled, last logon >90 days, missing email for service accounts
    if (-not $TemplateManagerNotCompliant) {
        $TemplateManagerNotCompliant = {
            EmailText -LineBreak
            EmailText -Text "Hello $ManagerDisplayName," -LineBreak

            EmailText -Text "Below is a summary of accounts where there is missing 'critical' information. These accounts are either:"

            EmailList {
                EmailListItem -Text "Missing a Manager in the AD - please add an active manager"
                EmailListItem -Text "The Manager in AD is Disabled - please add an active manager"
                EmailListItem -Text "Manager Last logon >90 days - please confirm if the manager is still an employee/change the manager to an active manager"
                EmailListItem -Text "Manager is missing email - add manager email"
            }
            EmailText -Text "Please contact the respective local IT Service Desk (outlined in the below table) to update this Manager's attributes in the AD directly. The suggested action to take can be found in the below table." -LineBreak

            EmailTable -DataTable $ManagerUsersTableManagerNotCompliant -HideFooter

            EmailText -LineBreak

            EmailText -Text "Kind regards," -LineBreak
            EmailText -Text "IT Service Desk" -LineBreak
        }
    }
    if (-not $TemplateManagerNotCompliantSubject) {
        $TemplateManagerNotCompliantSubject = "[Password Escalation] Accounts are expiring with non-compliant manager"
    }

    if (-not $TemplateAdmin) {
        # Template to Admins with information summarizing what happened
        $TemplateAdmin = {
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

            EmailText -Text "Hope everything works correctly! " -LineBreak

            EmailText -Text "Kind regards," -LineBreak
            EmailText -Text "IT Service Desk" -LineBreak
        }
        if (-not $TemplateAdminSubject) {
            $TemplateAdminSubject = '[Password Summary] Passwords summary'
        }
    }

    # Lets return information to the caller
    $OutputInformation = [ordered] @{
        EmailParameters                    = $EmailParameters
        UserSection                        = $UserSection
        ManagerSection                     = $ManagerSection
        SecuritySection                    = $SecuritySection
        AdminSection                       = $AdminSection
        HTMLReports                        = $HTMLReports
        Rules                              = $Rules
        SearchPath                         = $SearchPath
        OverwriteEmailProperty             = $OverwriteEmailProperty
        OverwriteManagerProperty           = $OverwriteManagerProperty
        Logging                            = $Logging
        TemplatePreExpiry                  = $TemplatePreExpiry
        TemplatePreExpirySubject           = $TemplatePreExpirySubject
        TemplatePostExpiry                 = $TemplatePostExpiry
        TemplatePostExpirySubject          = $TemplatePostExpirySubject
        TemplateManager                    = $TemplateManager
        TemplateManagerSubject             = $TemplateManagerSubject
        TemplateSecurity                   = $TemplateSecurity
        TemplateSecuritySubject            = $TemplateSecuritySubject
        TemplateManagerNotCompliant        = $TemplateManagerNotCompliant
        TemplateManagerNotCompliantSubject = $TemplateManagerNotCompliantSubject
        TemplateAdmin                      = $TemplateAdmin
        TemplateAdminSubject               = $TemplateAdminSubject
        UsersExternalSystem                = $UsersExternalSystem
        FilterOrganizationalUnit           = $FilterOrganizationalUnit
    }
    $OutputInformation
}