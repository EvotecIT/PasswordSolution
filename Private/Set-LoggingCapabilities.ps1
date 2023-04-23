function Set-LoggingCapabilities {
    [CmdletBinding()]
    param(
        [string] $LogPath,
        [int] $LogMaximum,
        [switch] $ShowTime,
        [string] $TimeFormat
    )

    $Script:PSDefaultParameterValues = @{
        "Write-Color:LogFile"    = $LogPath
        "Write-Color:ShowTime"   = if ($PSBoundParameters.ContainsKey('ShowTime')) { $ShowTime.IsPresent } else { $null }
        "Write-Color:TimeFormat" = $TimeFormat
    }
    Remove-EmptyValue -Hashtable $Script:PSDefaultParameterValues

    if ($LogPath) {
        $FolderPath = [io.path]::GetDirectoryName($LogPath)
        if (-not (Test-Path -LiteralPath $FolderPath)) {
            $null = New-Item -Path $FolderPath -ItemType Directory -Force -WhatIf:$false
        }
        if ($LogMaximum -gt 0) {
            $CurrentLogs = Get-ChildItem -LiteralPath $FolderPath | Sort-Object -Property CreationTime -Descending | Select-Object -Skip $LogMaximum
            if ($CurrentLogs) {
                Write-Color -Text '[i] ', "Logs directory has more than ", $LogMaximum, " log files. Cleanup required..." -Color Yellow, DarkCyan, Red, DarkCyan
                foreach ($Log in $CurrentLogs) {
                    try {
                        Remove-Item -LiteralPath $Log.FullName -Confirm:$false -WhatIf:$false
                        Write-Color -Text '[+] ', "Deleted ", "$($Log.FullName)" -Color Yellow, White, Green
                    } catch {
                        Write-Color -Text '[-] ', "Couldn't delete log file $($Log.FullName). Error: ', "$($_.Exception.Message) -Color Yellow, White, Red
                    }
                }
            }
        } else {
            Write-Color -Text '[i] ', "LogMaximum is set to 0 (Unlimited). No log files will be deleted." -Color Yellow, DarkCyan
        }
    }
}