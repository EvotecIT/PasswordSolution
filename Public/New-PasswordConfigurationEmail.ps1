function New-PasswordConfigurationEmail {
    [Alias('New-PasswordQualityEmailConfiguration')]
    [cmdletBinding(DefaultParameterSetName = 'Compatibility', SupportsShouldProcess)]
    param(
        [Parameter(ParameterSetName = 'SecureString')]

        [Parameter(ParameterSetName = 'oAuth')]
        [Parameter(ParameterSetName = 'Compatibility')]
        [alias('SmtpServer')][string] $Server,

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
        [MailKit.Net.Smtp.DeliveryStatusNotificationType] $DeliveryStatusNotificationType,

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
        [MailKit.Security.SecureSocketOptions] $SecureSocketOptions,

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