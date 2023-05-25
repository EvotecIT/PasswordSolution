function Export-SearchInformation {
    [CmdletBinding()]
    param(
        [string] $SearchPath,
        [System.Collections.IDictionary] $SummarySearch,
        [string] $Today,
        [Array] $SummaryUsersEmails,
        [Array] $SummaryManagersEmails,
        [Array] $SummaryEscalationEmails
    )

    if ($SearchPath) {
        Write-Color -Text "[i]" , " Saving Search report " -Color White, Yellow, Green
        if ($SummaryUsersEmails) {
            $SummarySearch['EmailSent'][$Today] += $SummaryUsersEmails
        }
        if ($SummaryEscalationEmails) {
            $SummarySearch['EmailEscalations'][$Today] += $SummaryEscalationEmails
        }
        if ($SummaryManagersEmails) {
            $SummarySearch['EmailManagers'][$Today] += $SummaryManagersEmails
        }
        try {
            $SummarySearch | Export-Clixml -LiteralPath $SearchPath -ErrorAction Stop
        } catch {
            Write-Color -Text "[e]", " Couldn't save to file $SearchPath", ". Error: ", $_.Exception.Message -Color White, Yellow, White, Yellow, White, Yellow, White
        }
        Write-Color -Text "[i]" , " Saving Search report ", "Done" -Color White, Yellow, Green
    }
}