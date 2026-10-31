function New-PasswordConfigurationEmail {
    <#
    .SYNOPSIS
    Defines mail delivery settings for a PasswordSolution configuration.
    .DESCRIPTION
    Passes the selected settings to Mailozaurr. Multiple SMTP servers default to ordered
    failover. Permanent rejections and uncertain delivery are reported without automatic resend.
    .PARAMETER Server
    Hosts or endpoint URIs: smtp:// for plaintext, smtps:// for implicit TLS,
    and smtp+starttls:// for required STARTTLS. Bare hosts use Port and SecureSocketOptions.
    .PARAMETER Strategy
    Failover, RoundRobin or Random starting-server selection across a notification batch.
    .PARAMETER CircuitBreakMinutes
    Host cooldown after a transient failure. Zero disables cooldown.
    .PARAMETER RetryCount
    Additional attempts. The total budget is at least the number of configured servers.
    .PARAMETER RetryDelayMilliseconds
    Initial retry delay; Mailozaurr applies bounded exponential backoff.
    .PARAMETER Port
    Default SMTP port for bare hosts.
    .PARAMETER From
    Sender address or named sender object accepted by Mailozaurr.
    .PARAMETER ReplyTo
    Reply address for notifications.
    .PARAMETER Priority
    Message priority: Low, Normal or High.
    .PARAMETER DeliveryNotificationOption
    SMTP delivery status notification conditions.
    .PARAMETER DeliveryStatusNotificationType
    MailKit delivery status notification format.
    .PARAMETER Credential
    Credentials for the selected SMTP, Graph or SendGrid provider.
    .PARAMETER Username
    SMTP authentication username.
    .PARAMETER Password
    SMTP password, optionally protected with AsSecureString.
    .PARAMETER SecureSocketOptions
    TLS mode for bare SMTP hosts.
    .PARAMETER UseSsl
    Require STARTTLS when the TLS mode is Auto.
    .PARAMETER SkipCertificateRevocation
    Disable SMTP certificate revocation checks.
    .PARAMETER SkipCertificateValidation
    Disable SMTP certificate validation.
    .PARAMETER Timeout
    SMTP operation timeout in milliseconds.
    .PARAMETER oAuth2
    Authenticate SMTP with an OAuth access token.
    .PARAMETER RequestReadReceipt
    Request a read receipt through Graph.
    .PARAMETER RequestDeliveryReceipt
    Request a delivery receipt through Graph.
    .PARAMETER Graph
    Deliver through Mailozaurr's Graph provider.
    .PARAMETER MgGraphRequest
    Deliver through Invoke-MgGraphRequest.
    .PARAMETER AsSecureString
    Interpret Password as a protected string.
    .PARAMETER SendGrid
    Deliver through SendGrid.
    .PARAMETER SeparateTo
    Send separate recipient messages through SendGrid.
    .PARAMETER DoNotSaveToSentItems
    Disable saving Graph messages to Sent Items.
    .PARAMETER LocalDomain
    SMTP client domain setting.
    .EXAMPLE
    New-PasswordConfigurationEmail -From 'notifications@example.com' -Server 'relay-a.example.com', 'relay-b.example.com' -Port 25 -Strategy Failover
    .OUTPUTS
    A PasswordConfigurationEmail record consumed by Start-PasswordSolution.
    #>
    [cmdletBinding(DefaultParameterSetName = 'Compatibility', SupportsShouldProcess)]
    param(
        [Parameter(ParameterSetName = 'SecureString')]

        [Parameter(ParameterSetName = 'oAuth')]
        [Parameter(ParameterSetName = 'Compatibility')]
        [alias('SmtpServer')][string[]] $Server,

        [Parameter(ParameterSetName = 'SecureString')]
        [Parameter(ParameterSetName = 'oAuth')]
        [Parameter(ParameterSetName = 'Compatibility')]
        [ValidateSet('Failover', 'RoundRobin', 'Random')][string] $Strategy,

        [Parameter(ParameterSetName = 'SecureString')]
        [Parameter(ParameterSetName = 'oAuth')]
        [Parameter(ParameterSetName = 'Compatibility')]
        [ValidateRange(0, 1440)][int] $CircuitBreakMinutes,

        [Parameter(ParameterSetName = 'SecureString')]
        [Parameter(ParameterSetName = 'oAuth')]
        [Parameter(ParameterSetName = 'Compatibility')]
        [ValidateRange(0, 100)][int] $RetryCount,

        [Parameter(ParameterSetName = 'SecureString')]
        [Parameter(ParameterSetName = 'oAuth')]
        [Parameter(ParameterSetName = 'Compatibility')]
        [ValidateRange(0, 60000)][int] $RetryDelayMilliseconds,

        [Parameter(ParameterSetName = 'SecureString')]

        [Parameter(ParameterSetName = 'oAuth')]
        [Parameter(ParameterSetName = 'Compatibility')]
        [int] $Port,

        [Parameter(Mandatory, ParameterSetName = 'SecureString')]
        [Parameter(Mandatory, ParameterSetName = 'oAuth')]
        [Parameter(Mandatory, ParameterSetName = 'Graph')]
        [Parameter(Mandatory, ParameterSetName = 'MgGraphRequest')]
        [Parameter(Mandatory, ParameterSetName = 'Compatibility')]
        [Parameter(Mandatory, ParameterSetName = 'SendGrid')]
        [object] $From,

        [Parameter(ParameterSetName = 'SecureString')]
        [Parameter(ParameterSetName = 'oAuth')]
        [Parameter(ParameterSetName = 'Graph')]
        [Parameter(ParameterSetName = 'MgGraphRequest')]
        [Parameter(ParameterSetName = 'Compatibility')]
        [Parameter(ParameterSetName = 'SendGrid')]
        [string] $ReplyTo,


        [Parameter(ParameterSetName = 'SecureString')]
        [Parameter(ParameterSetName = 'oAuth')]
        [Parameter(ParameterSetName = 'Graph')]
        [Parameter(ParameterSetName = 'MgGraphRequest')]
        [Parameter(ParameterSetName = 'Compatibility')]
        [Parameter(ParameterSetName = 'SendGrid')]
        [alias('Importance')][ValidateSet('Low', 'Normal', 'High')][string] $Priority,

        [Parameter(ParameterSetName = 'SecureString')]
        [Parameter(ParameterSetName = 'oAuth')]
        [Parameter(ParameterSetName = 'Compatibility')]
        [ValidateSet('None', 'OnSuccess', 'OnFailure', 'Delay', 'Never')][string[]] $DeliveryNotificationOption,

        [Parameter(ParameterSetName = 'SecureString')]
        [Parameter(ParameterSetName = 'oAuth')]
        [Parameter(ParameterSetName = 'Compatibility')]
        [ValidateSet('Unspecified', 'Full', 'HeadersOnly')][string] $DeliveryStatusNotificationType,

        [Parameter(ParameterSetName = 'oAuth')]
        [Parameter(Mandatory, ParameterSetName = 'Graph')]
        [Parameter(ParameterSetName = 'Compatibility')]
        [Parameter(Mandatory, ParameterSetName = 'SendGrid')]
        [pscredential] $Credential,

        [Parameter(ParameterSetName = 'SecureString')]
        [string] $Username,

        [Parameter(ParameterSetName = 'SecureString')]
        [string] $Password,

        [Parameter(ParameterSetName = 'SecureString')]
        [Parameter(ParameterSetName = 'oAuth')]
        [Parameter(ParameterSetName = 'Compatibility')]
        [ValidateSet('None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable')][string] $SecureSocketOptions,

        [Parameter(ParameterSetName = 'SecureString')]
        [Parameter(ParameterSetName = 'oAuth')]
        [Parameter(ParameterSetName = 'Compatibility')]
        [switch] $UseSsl,

        [Parameter(ParameterSetName = 'SecureString')]
        [Parameter(ParameterSetName = 'oAuth')]
        [Parameter(ParameterSetName = 'Compatibility')]
        [switch] $SkipCertificateRevocation,

        [Parameter(ParameterSetName = 'SecureString')]
        [Parameter(ParameterSetName = 'oAuth')]
        [Parameter(ParameterSetName = 'Compatibility')]
        [alias('SkipCertificateValidatation')][switch] $SkipCertificateValidation,

        [Parameter(ParameterSetName = 'SecureString')]
        [Parameter(ParameterSetName = 'oAuth')]
        [Parameter(ParameterSetName = 'Compatibility')]
        [int] $Timeout,

        [Parameter(ParameterSetName = 'oAuth')]
        [alias('oAuth')][switch] $oAuth2,

        [Parameter(ParameterSetName = 'Graph')]
        [Parameter(ParameterSetName = 'MgGraphRequest')]
        [switch] $RequestReadReceipt,

        [Parameter(ParameterSetName = 'Graph')]
        [Parameter(ParameterSetName = 'MgGraphRequest')]
        [switch] $RequestDeliveryReceipt,

        [Parameter(ParameterSetName = 'Graph')]
        [Parameter(ParameterSetName = 'MgGraphRequest')]
        [switch] $Graph,

        [Parameter(ParameterSetName = 'MgGraphRequest')]
        [switch] $MgGraphRequest,

        [Parameter(ParameterSetName = 'SecureString')]
        [switch] $AsSecureString,

        [Parameter(ParameterSetName = 'SendGrid')]
        [switch] $SendGrid,

        [Parameter(ParameterSetName = 'SendGrid')]
        [switch] $SeparateTo,

        [Parameter(ParameterSetName = 'Graph')]
        [Parameter(ParameterSetName = 'MgGraphRequest')]
        [switch] $DoNotSaveToSentItems,

        [Parameter(ParameterSetName = 'SecureString')]
        [Parameter(ParameterSetName = 'oAuth')]
        [Parameter(ParameterSetName = 'Compatibility')]
        [string] $LocalDomain
    )

    $Output = [ordered] @{
        Type     = 'PasswordConfigurationEmail'
        Settings = [ordered] @{
            Server                         = if ($PSBoundParameters.ContainsKey('Server')) { $Server } else { $null }
            Strategy                       = if ($PSBoundParameters.ContainsKey('Strategy')) { $Strategy } else { $null }
            CircuitBreakMinutes            = if ($PSBoundParameters.ContainsKey('CircuitBreakMinutes')) { $CircuitBreakMinutes } else { $null }
            RetryCount                     = if ($PSBoundParameters.ContainsKey('RetryCount')) { $RetryCount } else { $null }
            RetryDelayMilliseconds         = if ($PSBoundParameters.ContainsKey('RetryDelayMilliseconds')) { $RetryDelayMilliseconds } else { $null }
            Port                           = if ($PSBoundParameters.ContainsKey('Port')) { $Port } else { $null }
            From                           = if ($PSBoundParameters.ContainsKey('From')) { $From } else { $null }
            ReplyTo                        = if ($PSBoundParameters.ContainsKey('ReplyTo')) { $ReplyTo } else { $null }
            Priority                       = if ($PSBoundParameters.ContainsKey('Priority')) { $Priority } else { $null }
            DeliveryNotificationOption     = if ($PSBoundParameters.ContainsKey('DeliveryNotificationOption')) { $DeliveryNotificationOption } else { $null }
            DeliveryStatusNotificationType = if ($PSBoundParameters.ContainsKey('DeliveryStatusNotificationType')) { $DeliveryStatusNotificationType } else { $null }
            Credential                     = if ($PSBoundParameters.ContainsKey('Credential')) { $Credential } else { $null }
            Username                       = if ($PSBoundParameters.ContainsKey('Username')) { $Username } else { $null }
            Password                       = if ($PSBoundParameters.ContainsKey('Password')) { $Password } else { $null }
            SecureSocketOptions            = if ($PSBoundParameters.ContainsKey('SecureSocketOptions')) { $SecureSocketOptions } else { $null }
            UseSsl                         = if ($PSBoundParameters.ContainsKey('UseSsl')) { $UseSsl } else { $null }
            SkipCertificateRevocation      = if ($PSBoundParameters.ContainsKey('SkipCertificateRevocation')) { $SkipCertificateRevocation } else { $null }
            SkipCertificateValidation      = if ($PSBoundParameters.ContainsKey('SkipCertificateValidatation')) { $SkipCertificateValidation } else { $null }
            Timeout                        = if ($PSBoundParameters.ContainsKey('Timeout')) { $Timeout } else { $null }
            oAuth2                         = if ($PSBoundParameters.ContainsKey('oAuth2')) { $oAuth2 } else { $null }
            RequestReadReceipt             = if ($PSBoundParameters.ContainsKey('RequestReadReceipt')) { $RequestReadReceipt } else { $null }
            RequestDeliveryReceipt         = if ($PSBoundParameters.ContainsKey('RequestDeliveryReceipt')) { $RequestDeliveryReceipt } else { $null }
            Graph                          = if ($PSBoundParameters.ContainsKey('Graph')) { $Graph } else { $null }
            MgGraphRequest                 = if ($PSBoundParameters.ContainsKey('MgGraphRequest')) { $MgGraphRequest } else { $null }
            AsSecureString                 = if ($PSBoundParameters.ContainsKey('AsSecureString')) { $AsSecureString } else { $null }
            SendGrid                       = if ($PSBoundParameters.ContainsKey('SendGrid')) { $SendGrid } else { $null }
            SeparateTo                     = if ($PSBoundParameters.ContainsKey('SeparateTo')) { $SeparateTo } else { $null }
            DoNotSaveToSentItems           = if ($PSBoundParameters.ContainsKey('DoNotSaveToSentItems')) { $DoNotSaveToSentItems } else { $null }
            WhatIf                         = $WhatIfPreference
        }
    }
    Remove-EmptyValue -Hashtable $Output.Settings
    $Output
}
