function Get-PasswordNotificationFailures {
    <#
    .SYNOPSIS
    Returns the latest unsuccessful user/rule deliveries from the notification history.
    .DESCRIPTION
    A later successful delivery resolves an earlier failure. Legacy empty statuses remain
    recoverable. The history contains delivery metadata, without saved message bodies or credentials.
    #>
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $SummarySearch
    )

    if ($null -eq $SummarySearch -or $SummarySearch.EmailSent -isnot [System.Collections.IDictionary]) {
        throw 'The search history does not contain a valid EmailSent collection.'
    }
    $Latest = @{}
    foreach ($Batch in $SummarySearch.EmailSent.Values) {
        foreach ($Record in $Batch) {
            $Key = '{0}|{1}|{2}' -f $Record.Domain, $Record.UserPrincipalName, $Record.Rule
            if (-not $Latest.ContainsKey($Key) -or $Record.StatusWhen -ge $Latest[$Key].StatusWhen) {
                $Latest[$Key] = $Record
            }
        }
    }
    $Latest.Values | Where-Object { $_.Status -ne $true }
}
