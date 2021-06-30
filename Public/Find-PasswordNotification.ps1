function Find-PasswordNotification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SearchPath,
        [switch] $Manager
    )
    if ($SearchPath) {
        if (Test-Path -LiteralPath $SearchPath) {
            try {
                $SummarySearch = Import-Clixml -LiteralPath $SearchPath -ErrorAction Stop
                #$SummarySearch = Get-Content -LiteralPath $SearchPath -Raw | ConvertFrom-Json
            } catch {
                Write-Color -Text "[e]", " Couldn't load the file $SearchPath", ". Skipping...", $_.Exception.Message -Color White, Yellow, White, Yellow, White, Yellow, White
            }
            if ($SummarySearch -and $Manager) {
                $SummarySearch.EmailEscalations.Values
            } elseif ($SummarySearch -and $Manager -eq $false) {
                $SummarySearch.EmailSent.Values
            }
        }
    }
}