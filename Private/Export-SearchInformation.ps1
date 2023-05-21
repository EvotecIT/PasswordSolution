﻿function Export-SearchInformation {
    [CmdletBinding()]
    param(
        [string] $SearchPath,
        [System.Collections.IDictionary] $SummarySearch,
        [string] $Today
    )

    if ($SearchPath) {
        Write-Color -Text "[i]" , " Saving Search report " -Color White, Yellow, Green
        $SummarySearch['EmailSent'][$Today] = $SummaryUsersEmails
        $SummarySearch['EmailEscalations'][$Today] = $SummaryEscalationEmails
        $SummarySearch['EmailManagers'][$Today] = $SummaryManagersEmails

        try {
            $SummarySearch | Export-Clixml -LiteralPath $SearchPath
        } catch {
            Write-Color -Text "[e]", " Couldn't save to file $SearchPath", ". Error: ", $_.Exception.Message -Color White, Yellow, White, Yellow, White, Yellow, White
        }
        Write-Color -Text "[i]" , " Saving Search report ", "Done" -Color White, Yellow, Green
    }
}