function Import-SearchInformation {
    [CmdletBinding()]
    param(
        [string] $SearchPath
    )
    if ($SearchPath) {
        if (Test-Path -LiteralPath $SearchPath) {
            try {
                Write-Color -Text "[i]", " Loading file ", $SearchPath -Color White, Yellow, White, Yellow, White, Yellow, White
                $SummarySearch = Import-Clixml -LiteralPath $SearchPath -ErrorAction Stop
            } catch {
                Write-Color -Text "[e]", " Couldn't load the file $SearchPath", ". Skipping...", $_.Exception.Message -Color White, Yellow, White, Yellow, White, Yellow, White
            }
        }
    }
    if (-not $SummarySearch) {
        $SummarySearch = [ordered] @{
            EmailSent        = [ordered] @{}
            EmailManagers    = [ordered] @{}
            EmailEscalations = [ordered] @{}
        }
    } 
    $SummarySearch

}