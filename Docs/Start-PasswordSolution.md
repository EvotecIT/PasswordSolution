---
external help file: PasswordSolution-help.xml
Module Name: PasswordSolution
online version:
schema: 2.0.0
---

# Start-PasswordSolution

## SYNOPSIS
Starts Password Expiry Notifications for the whole forest

## SYNTAX

### DSL (Default)
```
Start-PasswordSolution [[-ConfigurationDSL] <ScriptBlock>] [<CommonParameters>]
```

### Legacy
```
Start-PasswordSolution -EmailParameters <IDictionary> [-OverwriteEmailProperty <String>]
 [-OverwriteManagerProperty <String>] -UserSection <IDictionary> -ManagerSection <IDictionary>
 -SecuritySection <IDictionary> -AdminSection <IDictionary> -Rules <Array> [-TemplatePreExpiry <ScriptBlock>]
 [-TemplatePreExpirySubject <String>] [-TemplatePostExpiry <ScriptBlock>] [-TemplatePostExpirySubject <String>]
 -TemplateManager <ScriptBlock> -TemplateManagerSubject <String> -TemplateSecurity <ScriptBlock>
 -TemplateSecuritySubject <String> -TemplateManagerNotCompliant <ScriptBlock>
 -TemplateManagerNotCompliantSubject <String> -TemplateAdmin <ScriptBlock> -TemplateAdminSubject <String>
 [-Logging <IDictionary>] [-HTMLReports <Array>] [-SearchPath <String>] [<CommonParameters>]
```

## DESCRIPTION
Starts Password Expiry Notifications for the whole forest

## EXAMPLES

### EXAMPLE 1
```
An example
```

## PARAMETERS

### -ConfigurationDSL
Parameter description

```yaml
Type: ScriptBlock
Parameter Sets: DSL
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EmailParameters
Parameters for Email.
Uses Mailozaurr splatting behind the scenes, so it supports all options that Mailozaurr does.

```yaml
Type: IDictionary
Parameter Sets: Legacy
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OverwriteEmailProperty
Property responsible for overwriting the default email field in Active Directory.
Useful when the password notification has to go somewhere else than users email address.

```yaml
Type: String
Parameter Sets: Legacy
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OverwriteManagerProperty
{{ Fill OverwriteManagerProperty Description }}

```yaml
Type: String
Parameter Sets: Legacy
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UserSection
Parameter description

```yaml
Type: IDictionary
Parameter Sets: Legacy
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ManagerSection
Parameter description

```yaml
Type: IDictionary
Parameter Sets: Legacy
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SecuritySection
Parameter description

```yaml
Type: IDictionary
Parameter Sets: Legacy
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AdminSection
Parameter description

```yaml
Type: IDictionary
Parameter Sets: Legacy
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Rules
Parameter description

```yaml
Type: Array
Parameter Sets: Legacy
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplatePreExpiry
Parameter description

```yaml
Type: ScriptBlock
Parameter Sets: Legacy
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplatePreExpirySubject
Parameter description

```yaml
Type: String
Parameter Sets: Legacy
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplatePostExpiry
Parameter description

```yaml
Type: ScriptBlock
Parameter Sets: Legacy
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplatePostExpirySubject
Parameter description

```yaml
Type: String
Parameter Sets: Legacy
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateManager
Parameter description

```yaml
Type: ScriptBlock
Parameter Sets: Legacy
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateManagerSubject
Parameter description

```yaml
Type: String
Parameter Sets: Legacy
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateSecurity
Parameter description

```yaml
Type: ScriptBlock
Parameter Sets: Legacy
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateSecuritySubject
Parameter description

```yaml
Type: String
Parameter Sets: Legacy
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateManagerNotCompliant
Parameter description

```yaml
Type: ScriptBlock
Parameter Sets: Legacy
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateManagerNotCompliantSubject
Parameter description

```yaml
Type: String
Parameter Sets: Legacy
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateAdmin
Parameter description

```yaml
Type: ScriptBlock
Parameter Sets: Legacy
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateAdminSubject
Parameter description

```yaml
Type: String
Parameter Sets: Legacy
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Logging
Parameter description

```yaml
Type: IDictionary
Parameter Sets: Legacy
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HTMLReports
Parameter description

```yaml
Type: Array
Parameter Sets: Legacy
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SearchPath
Parameter description

```yaml
Type: String
Parameter Sets: Legacy
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES
General notes

## RELATED LINKS
