function Send-PasswordSecurityNotifications {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $SecuritySection,
        [System.Collections.IDictionary] $Summary,
        [ScriptBlock]  $TemplateSecurity,
        [string] $TemplateSecuritySubject,
        [System.Collections.IDictionary] $Logging
    )

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

            if ($SecuritySection.AttachCSV -and $ManagedUsers.Count -gt 0) {
                $ManagedUsers | Export-Csv -LiteralPath $Env:TEMP\ManagedUsersSecurity.csv -NoTypeInformation -Force -Encoding UTF8 -ErrorAction Stop
                $EmailSplat.Attachments = @(
                    if (Test-Path -LiteralPath "$Env:TEMP\ManagedUsersSecurity.csv") {
                        "$Env:TEMP\ManagedUsersSecurity.csv"
                    }
                )
            }
            $EmailSplat.User = $ManagerUser
            $EmailSplat.ManagedUsers = $ManagedUsers | Select-Object -Property 'Status', 'DisplayName', 'Enabled', 'SamAccountName', 'Domain', 'DateExpiry', 'DaysToExpire', 'PasswordLastSet', 'PasswordExpired'
            #$EmailSplat.ManagedUsersManagerNotCompliant = $ManagedUsersManagerNotCompliant
            #$EmailSplat.ManagedUsersManagerDisabled = $ManagedUsersManagerDisabled
            #$EmailSplat.ManagedUsersManagerMissing = $ManagedUsersManagerMissing
            #$EmailSplat.ManagedUsersManagerMissingEmail = $ManagedUsersManagerMissingEmail
            $EmailSplat.EmailParameters = $EmailParameters

            $EmailSplat.EmailDateFormat = $Logging.EmailDateFormat
            $EmailSplat.EmailDateFormatUTCConversion = $Logging.EmailDateFormatUTCConversion

            if ($SecuritySection.SendToDefaultEmail -ne $true) {
                $EmailSplat.EmailParameters.To = $ManagerUser.EmailAddress
            } else {
                $EmailSplat.EmailParameters.To = $SecuritySection.DefaultEmail
            }
            if ($Logging.NotifyOnSecuritySend) {
                Write-Color -Text "[i] Sending notifications to security ", $ManagerUser.DisplayName, " (", $ManagerUser.EmailAddress, ") (SendToDefaultEmail: ", $ManagerSection.SendToDefaultEmail, ")" -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow
            }
            $EmailResult = Send-PasswordEmail @EmailSplat
            if ($Logging.NotifyOnSecuritySend) {
                if ($EmailResult.Error) {
                    Write-Color -Text "[r] Sending notifications to security ", $ManagerUser.DisplayName, " (", $ManagerUser.EmailAddress, ") (SendToDefaultEmail: ", $ManagerSection.SendToDefaultEmail, ") (status: ", $EmailResult.Status, " sent to: ", $EmailResult.SentTo, ", error: ", $EmailResult.Error, ")" -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow
                } else {
                    Write-Color -Text "[r] Sending notifications to security ", $ManagerUser.DisplayName, " (", $ManagerUser.EmailAddress, ") (SendToDefaultEmail: ", $ManagerSection.SendToDefaultEmail, ") (status: ", $EmailResult.Status, " sent to: ", $EmailResult.SentTo, ")" -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow
                }
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
            if ($SecuritySection.SendCountMaximum -gt 0) {
                if ($SecuritySection.SendCountMaximum -le $CountSecurity) {
                    Write-Color -Text "[i]", " Send count maximum reached. There may be more managers that match the rule." -Color Red, DarkRed
                    break
                }
            }
        }
        Write-Color -Text "[i] Sending notifications to security (sent: ", $SummaryEscalationEmails.Count, " out of ", $Summary['NotifySecurity'].Values.Count, ")" -Color White, Yellow, White, Yellow, White, Yellow, White
        $SummaryEscalationEmails
    } else {
        Write-Color -Text "[i] Sending notifications to security is ", "disabled!" -Color White, Yellow, DarkRed
    }

}