function Send-PasswordManagerNofifications {
    [CmdletBinding()]
    param(
        $ManagerSection,
        $Summary,
        $CachedUsers,
        $TemplateManager,
        $TemplateManagerSubject,
        $TemplateManagerNotCompliant,
        $TemplateManagerNotCompliantSubject,
        $EmailParameters,
        $Logging
    )
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
        $SummaryManagersEmails
    } else {
        Write-Color -Text "[i] Sending notifications to managers is ", "disabled!" -Color White, Yellow, DarkMagenta
    }
}