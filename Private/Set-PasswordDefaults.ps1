function Set-PasswordDefaults {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Logging
    )
    if (-not $Logging) {
        $Logging = @{
            ShowTime   = $true
            LogFile    = ""
            TimeFormat = "yyyy-MM-dd HH:mm:ss"
        }
    }
}