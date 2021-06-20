function Find-PasswordNotification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SearchPath,
        [System.Collections.IDictionary] $DisplayConsole,
        [switch] $Manager
    )
    #$Today = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    # Lets define Write-Color rules
    if ($null -eq $DisplayConsole) {
        $WriteParameters = @{
            ShowTime   = $true
            LogFile    = ""
            TimeFormat = "yyyy-MM-dd HH:mm:ss"
        }
    } else {
        $WriteParameters = $DisplayConsole
    }

    if ($SearchPath) {
        if (Test-Path -LiteralPath $SearchPath) {
            try {
                $SummarySearch = Import-Clixml -LiteralPath $SearchPath -ErrorAction Stop
                #$SummarySearch = Get-Content -LiteralPath $SearchPath -Raw | ConvertFrom-Json
            } catch {
                Write-Color @WriteParameters -Text "[e]", " Couldn't load the file $SearchPath", ". Skipping...", $_.Exception.Message -Color White, Yellow, White, Yellow, White, Yellow, White
            }
            if ($SummarySearch -and $Manager) {
                $SummarySearch.EmailEscalations.Values
            } elseif ($SummarySearch -and $Manager -eq $false) {
                $SummarySearch.EmailSent.Values
            }
        }
    }
}