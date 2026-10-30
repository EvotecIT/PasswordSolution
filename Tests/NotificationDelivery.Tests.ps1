BeforeAll {
    $RepositoryRoot = Split-Path $PSScriptRoot -Parent
    . "$RepositoryRoot/Private/Send-PasswordEmail.ps1"
    . "$RepositoryRoot/Private/Send-PasswordUserNotifications.ps1"
    . "$RepositoryRoot/Private/Get-PasswordNotificationFailures.ps1"
    . "$RepositoryRoot/Public/Find-PasswordNotification.ps1"

    function Write-Color { param($Text, $Color) }
    function EmailBody { param($EmailBody, $Parameter) 'Synthetic notification' }
    function Add-ParametersToString { param($String, $Parameter) $String }
    function Disable-ADAccount {
        [CmdletBinding(SupportsShouldProcess)]
        param($Identity)
        throw 'Synthetic tests must not disable accounts.'
    }
    function Send-EmailMessage {
        [CmdletBinding()]
        param($To, $From, $Subject, $Body, $Attachment)
        throw 'A test must provide its synthetic mail result.'
    }
}

Describe 'Notification delivery recovery' {
    BeforeEach {
        Mock Write-Color {}
    }

    It 'Reports the failed recipient and continues to the next selected user' {
        Mock Send-EmailMessage {
            if ($To -eq 'first@example.com') {
                throw [TimeoutException]::new('Synthetic SMTP timeout')
            }
            [pscustomobject] @{
                Status = $true
                SentTo = $To
                Error = $null
                Server = 'synthetic-relay'
                Attempts = @([pscustomobject] @{ Number = 1; Server = 'synthetic-relay' })
                DeliveryUnknown = $false
            }
        }
        $Notifications = [ordered] @{}
        foreach ($Name in 'first', 'second') {
            $Notifications[$Name] = [pscustomobject] @{
                User = [pscustomobject] @{
                    UserPrincipalName = "$Name@example.com"
                    EmailAddress = "$Name@example.com"
                    SamAccountName = $Name
                    DisplayName = $Name
                    Domain = 'example.com'
                    DaysToExpire = 3
                    Enabled = $true
                    DistinguishedName = "CN=$Name,DC=example,DC=com"
                }
                Rule = [pscustomobject] @{ Name = 'expiry'; DisableDays = @(3) }
            }
        }

        $Results = @(Send-PasswordUserNofifications -Summary @{ Notify = $Notifications } `
            -UserSection @{ Enable = $true } -Logging @{} -EmailParameters @{ From = 'sender@example.com' } `
            -TemplatePreExpiry { 'Synthetic notification' } -RetryFailedOnly)

        $Results.Count | Should -Be 2
        $Results[0].Status | Should -BeFalse
        $Results[0].SentTo | Should -Be 'first@example.com'
        $Results[0].StatusError | Should -Be 'Synthetic SMTP timeout'
        $Results[1].Status | Should -BeTrue
        $Results[1].SentTo | Should -Be 'second@example.com'
        $Results[1].Server | Should -Be 'synthetic-relay'
        $Results[0].Disabled | Should -BeNullOrEmpty
        $Results[1].Disabled | Should -BeNullOrEmpty
        @($Results[1].Attempts).Count | Should -Be 1
        Should -Invoke Write-Color -ParameterFilter {
            ($Text -join '') -like '*Failed to send email to first@example.com*'
        } -Times 1 -Exactly
        Should -Invoke Send-EmailMessage -Times 2 -Exactly
    }

    It 'Returns an explicit failure if a provider returns no result' {
        Mock Send-EmailMessage { }
        $Result = Send-PasswordEmail -User ([pscustomobject] @{}) -Template { 'Synthetic notification' } `
            -EmailParameters @{ To = 'recipient@example.com'; From = 'sender@example.com'; Suppress = $true }

        $Result.Status | Should -BeFalse
        $Result.SentTo | Should -Be 'recipient@example.com'
        $Result.Error | Should -Be 'The mail provider returned no delivery result.'
    }

    It 'Selects unresolved latest failures and recovers legacy empty statuses from saved history' {
        $Records = @(
            [pscustomobject] @{ Domain = 'example.com'; UserPrincipalName = 'resolved@example.com'; Rule = 'expiry'; Status = $false; StatusWhen = '2026-01-01 09:00:00' }
            [pscustomobject] @{ Domain = 'example.com'; UserPrincipalName = 'resolved@example.com'; Rule = 'expiry'; Status = $true; StatusWhen = '2026-01-01 10:00:00' }
            [pscustomobject] @{ Domain = 'example.com'; UserPrincipalName = 'legacy@example.com'; Rule = 'expiry'; Status = $null; StatusWhen = '2026-01-01 09:00:00' }
            [pscustomobject] @{ Domain = 'example.com'; UserPrincipalName = 'failed@example.com'; Rule = 'expiry'; Status = $false; StatusWhen = '2026-01-01 10:00:00' }
        )
        $HistoryPath = Join-Path $TestDrive 'notifications.xml'
        @{ EmailSent = @{ run = $Records } } | Export-Clixml -LiteralPath $HistoryPath

        $Failures = @(Find-PasswordNotification -SearchPath $HistoryPath -FailedOnly)

        $Failures.Count | Should -Be 2
        @($Failures.UserPrincipalName | Sort-Object) | Should -Be @('failed@example.com', 'legacy@example.com')
    }

    It 'Reports unavailable recovery history instead of an empty success' {
        $HistoryPath = Join-Path $TestDrive 'missing.xml'
        { Find-PasswordNotification -SearchPath $HistoryPath -FailedOnly } | Should -Throw '*does not exist*'
        @{ Unrelated = @{} } | Export-Clixml -LiteralPath $HistoryPath
        { Find-PasswordNotification -SearchPath $HistoryPath -FailedOnly } | Should -Throw '*valid EmailSent collection*'
    }
}
