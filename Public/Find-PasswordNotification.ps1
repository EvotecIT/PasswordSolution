function Find-PasswordNotification {
    <#
    .SYNOPSIS
    Searches thru XML logs created by Password Solution

    .DESCRIPTION
    Searches thru XML logs created by Password Solution

    .PARAMETER SearchPath
    Path to file where the XML log is located

    .PARAMETER Manager
    Search thru manager escalations

    .EXAMPLE
    Find-PasswordNotification -SearchPath $PSScriptRoot\Search\SearchLog.xml | Format-Table

    .EXAMPLE
    Find-PasswordNotification -SearchPath "$PSScriptRoot\Search\SearchLog_2021-06.xml" -Manager | Format-Table

    .NOTES
    General notes
    #>
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