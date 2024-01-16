function Send-PasswordAdminNotifications {
    [CmdletBinding()]
    param(
        $AdminSection,
        $TemplateAdmin,
        $TemplateAdminSubject,
        $TimeEnd,
        $EmailParameters,
        $HtmlAttachments,
        $Logging
    )

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

            $EmailSplat.EmailDateFormat = $Logging.EmailDateFormat
            $EmailSplat.EmailDateFormatUTCConversion = $Logging.EmailDateFormatUTCConversion

            if ($HtmlAttachments.Count -gt 0) {
                $EmailSplat.Attachments = $HtmlAttachments
            }
            Write-Color -Text "[i] Sending summary information ", $ManagerUser.DisplayName, " (", $ManagerUser.EmailAddress, ") (SendToDefaultEmail: ", $ManagerSection.SendToDefaultEmail, ")" -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow

            $EmailResult = Send-PasswordEmail @EmailSplat

            Write-Color -Text "[r] Sending summary information ", $ManagerUser.DisplayName, " (", $ManagerUser.EmailAddress, ") (SendToDefaultEmail: ", $ManagerSection.SendToDefaultEmail, ") (status: ", $EmailResult.Status, " sent to: ", $EmailResult.SentTo, ")" -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow
            if ($EmailResult.Error) {
                Write-Color -Text "[r] Error: ", $EmailResult.Error -Color White, Yellow, White, Yellow, White, Yellow, White, Yellow, White, Yellow
            }
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
}