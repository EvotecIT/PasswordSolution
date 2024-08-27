function Send-PasswordUserNofifications {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $UserSection,
        [System.Collections.IDictionary] $Summary,
        [System.Collections.IDictionary] $Logging,
        [ScriptBlock] $TemplatePreExpiry,
        [string] $TemplatePreExpirySubject,
        [scriptBlock] $TemplatePostExpiry,
        [string] $TemplatePostExpirySubject,
        [System.Collections.IDictionary] $EmailParameters
    )
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
                    $EmailSplat.Subject = '[Password] Your password will expire on $DateExpiry ($DaysToExpire days)'
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
                    $EmailSplat.Subject = '[Password] Your password expired on $DateExpiry ($DaysToExpire days ago)'
                }
            }
            $EmailSplat.User = $Notify.User
            $EmailSplat.EmailParameters = $EmailParameters

            $EmailSplat.EmailDateFormat = $Logging.EmailDateFormat
            $EmailSplat.EmailDateFormatUTCConversion = $Logging.EmailDateFormatUTCConversion

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
                    EmailFrom            = $EmailSplat.User.EmailFrom
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
                    EmailFrom            = $EmailSplat.User.EmailFrom
                }
            }
            if ($Logging.NotifyOnUserSend) {
                if ($EmailResult.SentTo) {
                    Write-Color -Text "[i]", " Sending notifications to user ", $Notify.User.DisplayName, " (", $Notify.User.EmailAddress, ")", " status: ", $EmailResult.Status, " sent to: ", $EmailResult.SentTo, ", details: ", $EmailResult.Error -Color Yellow, White, Yellow, White, Yellow, White, White, Blue, White, Blue
                } else {
                    Write-Color -Text "[i]", " Skipping notifications to user ", $Notify.User.DisplayName, " (", $Notify.User.EmailAddress, ")", " status: ", $EmailResult.Status, " details: ", $EmailResult.Error -Color Yellow, White, Yellow, White, Yellow, White, White, Blue, White, Blue
                }
            }
            if ($UserSection.SendCountMaximum -gt 0) {
                if ($UserSection.SendCountMaximum -le $CountUsers) {
                    Write-Color -Text "[i]", " Send count maximum reached. There may be more accounts that match the rule." -Color Red, DarkRed
                    break
                }
            }
        }
        Write-Color -Text "[i] Sending notifications to users (sent: ", $SummaryUsersEmails.Count, " out of ", $Summary['Notify'].Values.Count, ")" -Color White, Yellow, White, Yellow, White, Yellow, White
        $SummaryUsersEmails
    } else {
        Write-Color -Text "[i] Sending notifications to users is ", "disabled!" -Color White, Yellow, DarkRed
    }

}