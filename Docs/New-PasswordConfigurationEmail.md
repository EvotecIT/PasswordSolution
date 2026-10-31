---
external help file: PasswordSolution-help.xml
Module Name: PasswordSolution
online version: https://github.com/EvotecIT/PasswordSolution
schema: 2.0.0
---
# New-PasswordConfigurationEmail
## SYNOPSIS
Defines mail delivery settings for a PasswordSolution configuration.

## SYNTAX
### Compatibility (Default)
```powershell
New-PasswordConfigurationEmail -From <Object> [-Server <string[]>] [-Strategy <string>] [-CircuitBreakMinutes <int>] [-RetryCount <int>] [-RetryDelayMilliseconds <int>] [-Port <int>] [-ReplyTo <string>] [-Priority <string>] [-DeliveryNotificationOption <string[]>] [-DeliveryStatusNotificationType <string>] [-Credential <pscredential>] [-SecureSocketOptions <string>] [-UseSsl] [-SkipCertificateRevocation] [-SkipCertificateValidation] [-Timeout <int>] [-LocalDomain <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### oAuth
```powershell
New-PasswordConfigurationEmail -From <Object> [-Server <string[]>] [-Strategy <string>] [-CircuitBreakMinutes <int>] [-RetryCount <int>] [-RetryDelayMilliseconds <int>] [-Port <int>] [-ReplyTo <string>] [-Priority <string>] [-DeliveryNotificationOption <string[]>] [-DeliveryStatusNotificationType <string>] [-Credential <pscredential>] [-SecureSocketOptions <string>] [-UseSsl] [-SkipCertificateRevocation] [-SkipCertificateValidation] [-Timeout <int>] [-oAuth2] [-LocalDomain <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### SecureString
```powershell
New-PasswordConfigurationEmail -From <Object> [-Server <string[]>] [-Strategy <string>] [-CircuitBreakMinutes <int>] [-RetryCount <int>] [-RetryDelayMilliseconds <int>] [-Port <int>] [-ReplyTo <string>] [-Priority <string>] [-DeliveryNotificationOption <string[]>] [-DeliveryStatusNotificationType <string>] [-Username <string>] [-Password <string>] [-SecureSocketOptions <string>] [-UseSsl] [-SkipCertificateRevocation] [-SkipCertificateValidation] [-Timeout <int>] [-AsSecureString] [-LocalDomain <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### SendGrid
```powershell
New-PasswordConfigurationEmail -From <Object> -Credential <pscredential> [-ReplyTo <string>] [-Priority <string>] [-SendGrid] [-SeparateTo] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### MgGraphRequest
```powershell
New-PasswordConfigurationEmail -From <Object> [-ReplyTo <string>] [-Priority <string>] [-RequestReadReceipt] [-RequestDeliveryReceipt] [-Graph] [-MgGraphRequest] [-DoNotSaveToSentItems] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Graph
```powershell
New-PasswordConfigurationEmail -From <Object> -Credential <pscredential> [-ReplyTo <string>] [-Priority <string>] [-RequestReadReceipt] [-RequestDeliveryReceipt] [-Graph] [-DoNotSaveToSentItems] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Passes the selected settings to Mailozaurr. Multiple SMTP servers default to ordered
failover. Permanent rejections and uncertain delivery are reported without automatic resend.

## EXAMPLES

### EXAMPLE 1
```powershell
PS > New-PasswordConfigurationEmail -From 'notifications@example.com' -Server 'relay-a.example.com', 'relay-b.example.com' -Port 25 -Strategy Failover
```


## PARAMETERS

### -AsSecureString
Interpret Password as a protected string.

```yaml
Type: SwitchParameter
Parameter Sets: SecureString
Aliases: None
Possible values:

Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -CircuitBreakMinutes
Host cooldown after a transient failure. Zero disables cooldown.

```yaml
Type: Int32
Parameter Sets: Compatibility, oAuth, SecureString
Aliases: None
Possible values:

Required: False
Position: named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -Credential
Credentials for the selected SMTP, Graph or SendGrid provider.

```yaml
Type: PSCredential
Parameter Sets: Compatibility, oAuth, SendGrid, Graph
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DeliveryNotificationOption
SMTP delivery status notification conditions.

```yaml
Type: String[]
Parameter Sets: Compatibility, oAuth, SecureString
Aliases: None
Possible values: None, OnSuccess, OnFailure, Delay, Never

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DeliveryStatusNotificationType
MailKit delivery status notification format.

```yaml
Type: String
Parameter Sets: Compatibility, oAuth, SecureString
Aliases: None
Possible values: Unspecified, Full, HeadersOnly

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DoNotSaveToSentItems
Disable saving Graph messages to Sent Items.

```yaml
Type: SwitchParameter
Parameter Sets: MgGraphRequest, Graph
Aliases: None
Possible values:

Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -From
Sender address or named sender object accepted by Mailozaurr.

```yaml
Type: Object
Parameter Sets: Compatibility, oAuth, SecureString, SendGrid, MgGraphRequest, Graph
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Graph
Deliver through Mailozaurr's Graph provider.

```yaml
Type: SwitchParameter
Parameter Sets: MgGraphRequest, Graph
Aliases: None
Possible values:

Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -LocalDomain
SMTP client domain setting.

```yaml
Type: String
Parameter Sets: Compatibility, oAuth, SecureString
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MgGraphRequest
Deliver through Invoke-MgGraphRequest.

```yaml
Type: SwitchParameter
Parameter Sets: MgGraphRequest
Aliases: None
Possible values:

Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -oAuth2
Authenticate SMTP with an OAuth access token.

```yaml
Type: SwitchParameter
Parameter Sets: oAuth
Aliases: oAuth
Possible values:

Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
SMTP password, optionally protected with AsSecureString.

```yaml
Type: String
Parameter Sets: SecureString
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Port
Default SMTP port for bare hosts.

```yaml
Type: Int32
Parameter Sets: Compatibility, oAuth, SecureString
Aliases: None
Possible values:

Required: False
Position: named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -Priority
Message priority: Low, Normal or High.

```yaml
Type: String
Parameter Sets: Compatibility, oAuth, SecureString, SendGrid, MgGraphRequest, Graph
Aliases: Importance
Possible values: Low, Normal, High

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReplyTo
Reply address for notifications.

```yaml
Type: String
Parameter Sets: Compatibility, oAuth, SecureString, SendGrid, MgGraphRequest, Graph
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RequestDeliveryReceipt
Request a delivery receipt through Graph.

```yaml
Type: SwitchParameter
Parameter Sets: MgGraphRequest, Graph
Aliases: None
Possible values:

Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -RequestReadReceipt
Request a read receipt through Graph.

```yaml
Type: SwitchParameter
Parameter Sets: MgGraphRequest, Graph
Aliases: None
Possible values:

Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -RetryCount
Additional attempts. The total budget is at least the number of configured servers.

```yaml
Type: Int32
Parameter Sets: Compatibility, oAuth, SecureString
Aliases: None
Possible values:

Required: False
Position: named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -RetryDelayMilliseconds
Initial retry delay; Mailozaurr applies bounded exponential backoff.

```yaml
Type: Int32
Parameter Sets: Compatibility, oAuth, SecureString
Aliases: None
Possible values:

Required: False
Position: named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -SecureSocketOptions
TLS mode for bare SMTP hosts.

```yaml
Type: String
Parameter Sets: Compatibility, oAuth, SecureString
Aliases: None
Possible values: None, Auto, SslOnConnect, StartTls, StartTlsWhenAvailable

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SendGrid
Deliver through SendGrid.

```yaml
Type: SwitchParameter
Parameter Sets: SendGrid
Aliases: None
Possible values:

Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -SeparateTo
Send separate recipient messages through SendGrid.

```yaml
Type: SwitchParameter
Parameter Sets: SendGrid
Aliases: None
Possible values:

Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Server
Hosts or endpoint URIs: smtp:// for plaintext, smtps:// for implicit TLS,
and smtp+starttls:// for required STARTTLS. Bare hosts use Port and SecureSocketOptions.

```yaml
Type: String[]
Parameter Sets: Compatibility, oAuth, SecureString
Aliases: SmtpServer
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SkipCertificateRevocation
Disable SMTP certificate revocation checks.

```yaml
Type: SwitchParameter
Parameter Sets: Compatibility, oAuth, SecureString
Aliases: None
Possible values:

Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -SkipCertificateValidation
Disable SMTP certificate validation.

```yaml
Type: SwitchParameter
Parameter Sets: Compatibility, oAuth, SecureString
Aliases: SkipCertificateValidatation
Possible values:

Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strategy
Failover, RoundRobin or Random starting-server selection across a notification batch.

```yaml
Type: String
Parameter Sets: Compatibility, oAuth, SecureString
Aliases: None
Possible values: Failover, RoundRobin, Random

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Timeout
SMTP operation timeout in milliseconds.

```yaml
Type: Int32
Parameter Sets: Compatibility, oAuth, SecureString
Aliases: None
Possible values:

Required: False
Position: named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -Username
SMTP authentication username.

```yaml
Type: String
Parameter Sets: SecureString
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UseSsl
Require STARTTLS when the TLS mode is Auto.

```yaml
Type: SwitchParameter
Parameter Sets: Compatibility, oAuth, SecureString
Aliases: None
Possible values:

Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `A PasswordConfigurationEmail record consumed by Start-PasswordSolution.`

## RELATED LINKS

- None
