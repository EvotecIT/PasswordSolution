---
external help file: PasswordSolution-help.xml
Module Name: PasswordSolution
online version:
schema: 2.0.0
---

# New-PasswordConfigurationEmail

## SYNOPSIS
{{ Fill in the Synopsis }}

## SYNTAX

### Compatibility
```
New-PasswordConfigurationEmail [-Server <String>] [-Port <Int32>] -From <Object> [-ReplyTo <String>]
 [-Priority <String>] [-DeliveryNotificationOption <String[]>]
 [-DeliveryStatusNotificationType <DeliveryStatusNotificationType>] [-Credential <PSCredential>]
 [-SecureSocketOptions <SecureSocketOptions>] [-UseSsl] [-SkipCertificateRevocation]
 [-SkipCertificateValidatation] [-Timeout <Int32>] [-LocalDomain <String>] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

### oAuth
```
New-PasswordConfigurationEmail [-Server <String>] [-Port <Int32>] -From <Object> [-ReplyTo <String>]
 [-Priority <String>] [-DeliveryNotificationOption <String[]>]
 [-DeliveryStatusNotificationType <DeliveryStatusNotificationType>] [-Credential <PSCredential>]
 [-SecureSocketOptions <SecureSocketOptions>] [-UseSsl] [-SkipCertificateRevocation]
 [-SkipCertificateValidatation] [-Timeout <Int32>] [-oAuth2] [-LocalDomain <String>] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

### SecureString
```
New-PasswordConfigurationEmail [-Server <String>] [-Port <Int32>] -From <Object> [-ReplyTo <String>]
 [-Priority <String>] [-DeliveryNotificationOption <String[]>]
 [-DeliveryStatusNotificationType <DeliveryStatusNotificationType>] [-Username <String>] [-Password <String>]
 [-SecureSocketOptions <SecureSocketOptions>] [-UseSsl] [-SkipCertificateRevocation]
 [-SkipCertificateValidatation] [-Timeout <Int32>] [-AsSecureString] [-LocalDomain <String>] [-WhatIf]
 [-Confirm] [<CommonParameters>]
```

### SendGrid
```
New-PasswordConfigurationEmail -From <Object> [-ReplyTo <String>] [-Priority <String>]
 -Credential <PSCredential> [-SendGrid] [-SeparateTo] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### MgGraphRequest
```
New-PasswordConfigurationEmail -From <Object> [-ReplyTo <String>] [-Priority <String>] [-RequestReadReceipt]
 [-RequestDeliveryReceipt] [-Graph] [-MgGraphRequest] [-DoNotSaveToSentItems] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

### Graph
```
New-PasswordConfigurationEmail -From <Object> [-ReplyTo <String>] [-Priority <String>]
 -Credential <PSCredential> [-RequestReadReceipt] [-RequestDeliveryReceipt] [-Graph] [-DoNotSaveToSentItems]
 [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -AsSecureString
{{ Fill AsSecureString Description }}

```yaml
Type: SwitchParameter
Parameter Sets: SecureString
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Confirm
Prompts you for confirmation before running the cmdlet.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: cf

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Credential
{{ Fill Credential Description }}

```yaml
Type: PSCredential
Parameter Sets: Compatibility, oAuth
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

```yaml
Type: PSCredential
Parameter Sets: SendGrid, Graph
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DeliveryNotificationOption
{{ Fill DeliveryNotificationOption Description }}

```yaml
Type: String[]
Parameter Sets: Compatibility, oAuth, SecureString
Aliases:
Accepted values: None, OnSuccess, OnFailure, Delay, Never

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DeliveryStatusNotificationType
{{ Fill DeliveryStatusNotificationType Description }}

```yaml
Type: DeliveryStatusNotificationType
Parameter Sets: Compatibility, oAuth, SecureString
Aliases:
Accepted values: Unspecified, Full, HeadersOnly

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DoNotSaveToSentItems
{{ Fill DoNotSaveToSentItems Description }}

```yaml
Type: SwitchParameter
Parameter Sets: MgGraphRequest, Graph
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -From
{{ Fill From Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Graph
{{ Fill Graph Description }}

```yaml
Type: SwitchParameter
Parameter Sets: MgGraphRequest, Graph
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LocalDomain
{{ Fill LocalDomain Description }}

```yaml
Type: String
Parameter Sets: Compatibility, oAuth, SecureString
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MgGraphRequest
{{ Fill MgGraphRequest Description }}

```yaml
Type: SwitchParameter
Parameter Sets: MgGraphRequest
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
{{ Fill Password Description }}

```yaml
Type: String
Parameter Sets: SecureString
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Port
{{ Fill Port Description }}

```yaml
Type: Int32
Parameter Sets: Compatibility, oAuth, SecureString
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Priority
{{ Fill Priority Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases: Importance
Accepted values: Low, Normal, High

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReplyTo
{{ Fill ReplyTo Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RequestDeliveryReceipt
{{ Fill RequestDeliveryReceipt Description }}

```yaml
Type: SwitchParameter
Parameter Sets: MgGraphRequest, Graph
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RequestReadReceipt
{{ Fill RequestReadReceipt Description }}

```yaml
Type: SwitchParameter
Parameter Sets: MgGraphRequest, Graph
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SecureSocketOptions
{{ Fill SecureSocketOptions Description }}

```yaml
Type: SecureSocketOptions
Parameter Sets: Compatibility, oAuth, SecureString
Aliases:
Accepted values: None, Auto, SslOnConnect, StartTls, StartTlsWhenAvailable

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SendGrid
{{ Fill SendGrid Description }}

```yaml
Type: SwitchParameter
Parameter Sets: SendGrid
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SeparateTo
{{ Fill SeparateTo Description }}

```yaml
Type: SwitchParameter
Parameter Sets: SendGrid
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Server
{{ Fill Server Description }}

```yaml
Type: String
Parameter Sets: Compatibility, oAuth, SecureString
Aliases: SmtpServer

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SkipCertificateRevocation
{{ Fill SkipCertificateRevocation Description }}

```yaml
Type: SwitchParameter
Parameter Sets: Compatibility, oAuth, SecureString
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SkipCertificateValidatation
{{ Fill SkipCertificateValidatation Description }}

```yaml
Type: SwitchParameter
Parameter Sets: Compatibility, oAuth, SecureString
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Timeout
{{ Fill Timeout Description }}

```yaml
Type: Int32
Parameter Sets: Compatibility, oAuth, SecureString
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UseSsl
{{ Fill UseSsl Description }}

```yaml
Type: SwitchParameter
Parameter Sets: Compatibility, oAuth, SecureString
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Username
{{ Fill Username Description }}

```yaml
Type: String
Parameter Sets: SecureString
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WhatIf
Shows what would happen if the cmdlet runs. The cmdlet is not run.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: wi

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -oAuth2
{{ Fill oAuth2 Description }}

```yaml
Type: SwitchParameter
Parameter Sets: oAuth
Aliases: oAuth

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### None

## OUTPUTS

### System.Object
## NOTES

## RELATED LINKS
