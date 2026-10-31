---
external help file: PasswordSolution-help.xml
Module Name: PasswordSolution
online version: https://github.com/EvotecIT/PasswordSolution
schema: 2.0.0
---
# Start-PasswordSolution
## SYNOPSIS
Starts Password Expiry Notifications for the whole forest

## SYNTAX
### DSL (Default)
```powershell
Start-PasswordSolution [[-ConfigurationDSL] <scriptblock>] [-RetryFailedOnly] [-IncludeUnknownDelivery] [<CommonParameters>]
```

### Legacy
```powershell
Start-PasswordSolution [[-ConfigurationDSL] <scriptblock>] -EmailParameters <IDictionary> -UserSection <IDictionary> -ManagerSection <IDictionary> -SecuritySection <IDictionary> -AdminSection <IDictionary> -Rules <array> -TemplateManager <scriptblock> -TemplateManagerSubject <string> -TemplateSecurity <scriptblock> -TemplateSecuritySubject <string> -TemplateManagerNotCompliant <scriptblock> -TemplateManagerNotCompliantSubject <string> -TemplateAdmin <scriptblock> -TemplateAdminSubject <string> [-OverwriteEmailProperty <string>] [-OverwriteManagerProperty <string>] [-UsersExternalSystem <IDictionary>] [-TemplatePreExpiry <scriptblock>] [-TemplatePreExpirySubject <string>] [-TemplatePostExpiry <scriptblock>] [-TemplatePostExpirySubject <string>] [-Entra <IDictionary>] [-Logging <IDictionary>] [-HTMLReports <array>] [-SearchPath <string>] [-FilterOrganizationalUnit <string[]>] [-SearchBase <string[]>] [-RetryFailedOnly] [-IncludeUnknownDelivery] [<CommonParameters>]
```

## DESCRIPTION
Starts Password Expiry Notifications for the whole forest

## EXAMPLES

### EXAMPLE 1
```powershell
PS > Start-PasswordSolution -ConfigurationDSL $Configuration -RetryFailedOnly
Re-send the failed user notifications recorded in the configuration's SearchPath.
```


## PARAMETERS

### -AdminSection
Enables the administrator summary and controls its recipients.

```yaml
Type: IDictionary
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConfigurationDSL
Script block containing New-PasswordConfiguration commands for email, rules, templates and reports.

```yaml
Type: ScriptBlock
Parameter Sets: DSL, Legacy
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EmailParameters
Parameters for Email. Uses Mailozaurr splatting behind the scenes, so it supports all options that Mailozaurr does.

```yaml
Type: IDictionary
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Entra
Selects Microsoft Entra user discovery and its configuration instead of Active Directory discovery.

```yaml
Type: IDictionary
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: False
Position: named
Default value: [ordered] @{}
Accept pipeline input: False
Accept wildcard characters: False
```

### -FilterOrganizationalUnit
Provides a way to filter users by Organizational Unit limiting the scope of the search.
The search is performed using 'like' operator, so you can use wildcards if needed.

```yaml
Type: String[]
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HTMLReports
Report definitions controlling which HTML summaries are generated and where they are saved.

```yaml
Type: Array
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeUnknownDelivery
Include uncertain SMTP deliveries in RetryFailedOnly after checking that they were not delivered.
Without this switch they remain unresolved and are reported for manual verification.

```yaml
Type: SwitchParameter
Parameter Sets: DSL, Legacy
Aliases: None
Possible values:

Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Logging
Log file, console detail, timestamp and email date-format options.

```yaml
Type: IDictionary
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: False
Position: named
Default value: [ordered]  @{}
Accept pipeline input: False
Accept wildcard characters: False
```

### -ManagerSection
Enables manager notifications and controls manager delivery options.

```yaml
Type: IDictionary
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OverwriteEmailProperty
Property responsible for overwriting the default email field in Active Directory. Useful when the password notification has to go somewhere else than users email address.

```yaml
Type: String
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OverwriteManagerProperty
Directory property used instead of the default manager relationship.

```yaml
Type: String
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RetryFailedOnly
Re-evaluate current rules and templates for the latest failed user/rule deliveries in SearchPath.
Successful users, manager/security/admin notifications and account-disable actions are excluded.
Users who no longer match a current rule are reported and remain in the failed history.

```yaml
Type: SwitchParameter
Parameter Sets: DSL, Legacy
Aliases: None
Possible values:

Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Rules
Password expiry and account selection rules evaluated against the discovered users.

```yaml
Type: Array
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SearchBase
Provides a way to filter users by Organizational Unit limiting the scope of the search.
The search is passed to Get-ADUser cmdlet. This command should only be used when not using manager functionality.
Otherwise you won't be able to find matching managers for your users. Use FilterOrganizationalUnit instead.

```yaml
Type: String[]
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SearchPath
Path to the notification history saved after the run. Required for RetryFailedOnly.

```yaml
Type: String
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SecuritySection
Enables security escalations and controls their recipients.

```yaml
Type: IDictionary
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateAdmin
Body template for the administrator's run summary.

```yaml
Type: ScriptBlock
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateAdminSubject
Subject template for the administrator's run summary.

```yaml
Type: String
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateManager
Body template listing a manager's selected users.

```yaml
Type: ScriptBlock
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateManagerNotCompliant
Body template for users whose manager information does not meet the configured rule.

```yaml
Type: ScriptBlock
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateManagerNotCompliantSubject
Subject template for manager compliance notifications.

```yaml
Type: String
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateManagerSubject
Subject template for manager notifications.

```yaml
Type: String
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplatePostExpiry
Default notification body template for expired passwords.

```yaml
Type: ScriptBlock
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplatePostExpirySubject
Default subject template for expired passwords.

```yaml
Type: String
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplatePreExpiry
Default notification body template for passwords that have not expired.

```yaml
Type: ScriptBlock
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplatePreExpirySubject
Default subject template for passwords that have not expired.

```yaml
Type: String
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateSecurity
Body template for security escalations.

```yaml
Type: ScriptBlock
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateSecuritySubject
Subject template for security escalations.

```yaml
Type: String
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UserSection
Enables user notifications and controls default recipients and per-run send limits.

```yaml
Type: IDictionary
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UsersExternalSystem
Property responsible for overwriting the default email field in Active Directory.
Useful when the password notification has to go somewhere else than users email address.
It comes in a specific format as generated by `New-PasswordConfigurationExternalUsers`

```yaml
Type: IDictionary
Parameter Sets: Legacy
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `Run logging and configured reports. Delivery results, host attempts and final errors are saved in SearchPath.`

## RELATED LINKS

- None
