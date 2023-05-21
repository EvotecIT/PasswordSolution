---
external help file: PasswordSolution-help.xml
Module Name: PasswordSolution
online version:
schema: 2.0.0
---

# New-PasswordConfigurationRule

## SYNOPSIS
{{ Fill in the Synopsis }}

## SYNTAX

```
New-PasswordConfigurationRule [[-ReminderConfiguration] <ScriptBlock>] [-Name] <String> [-Enable]
 [-IncludeExpiring] [-IncludePasswordNeverExpires] [[-PasswordNeverExpiresDays] <Int32>]
 [[-IncludeNameProperties] <String[]>] [[-IncludeName] <String[]>] [[-ExcludeNameProperties] <String[]>]
 [[-ExcludeName] <String[]>] [[-IncludeOU] <String[]>] [[-ExcludeOU] <String[]>] [[-IncludeGroup] <String[]>]
 [[-ExcludeGroup] <String[]>] [-ReminderDays] <Array> [-ManagerReminder] [-ManagerNotCompliant]
 [[-ManagerNotCompliantDisplayName] <String>] [[-ManagerNotCompliantEmailAddress] <String>]
 [-ManagerNotCompliantDisabled] [-ManagerNotCompliantMissing] [-ManagerNotCompliantMissingEmail]
 [[-ManagerNotCompliantLastLogonDays] <Int32>] [-SecurityEscalation]
 [[-SecurityEscalationDisplayName] <String>] [[-SecurityEscalationEmailAddress] <String>]
 [[-OverwriteEmailProperty] <String>] [[-OverwriteManagerProperty] <String>]
 [[-OverwriteManagerPropertyName] <String>] [-ProcessManagersOnly] [<CommonParameters>]
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

### -ReminderConfiguration
{{ Fill ReminderConfiguration Description }}

```yaml
Type: ScriptBlock
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
{{ Fill Name Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Enable
{{ Fill Enable Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeExpiring
{{ Fill IncludeExpiring Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludePasswordNeverExpires
{{ Fill IncludePasswordNeverExpires Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PasswordNeverExpiresDays
{{ Fill PasswordNeverExpiresDays Description }}

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeNameProperties
{{ Fill IncludeNameProperties Description }}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeName
{{ Fill IncludeName Description }}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludeNameProperties
Exclude user from rule if any of the properties match the value as defined in ExcludeName

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludeName
Exclude user from rule if any of the properties match the value of Name in the properties defined in ExcludeNameProperties

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 7
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeOU
{{ Fill IncludeOU Description }}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 8
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludeOU
{{ Fill ExcludeOU Description }}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 9
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeGroup
Parameter description

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 10
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludeGroup
Parameter description

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 11
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReminderDays
Parameter description

```yaml
Type: Array
Parameter Sets: (All)
Aliases: ExpirationDays, Days

Required: True
Position: 12
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ManagerReminder
{{ Fill ManagerReminder Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ManagerNotCompliant
{{ Fill ManagerNotCompliant Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ManagerNotCompliantDisplayName
{{ Fill ManagerNotCompliantDisplayName Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 13
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ManagerNotCompliantEmailAddress
{{ Fill ManagerNotCompliantEmailAddress Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 14
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ManagerNotCompliantDisabled
{{ Fill ManagerNotCompliantDisabled Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ManagerNotCompliantMissing
{{ Fill ManagerNotCompliantMissing Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ManagerNotCompliantMissingEmail
{{ Fill ManagerNotCompliantMissingEmail Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ManagerNotCompliantLastLogonDays
{{ Fill ManagerNotCompliantLastLogonDays Description }}

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 15
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SecurityEscalation
{{ Fill SecurityEscalation Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SecurityEscalationDisplayName
{{ Fill SecurityEscalationDisplayName Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 16
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SecurityEscalationEmailAddress
{{ Fill SecurityEscalationEmailAddress Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 17
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OverwriteEmailProperty
Parameter description

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 18
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OverwriteManagerProperty
Parameter description

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 19
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OverwriteManagerPropertyName
Parameter description

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 20
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ProcessManagersOnly
This parameters is used to process users, but only managers will be notified.
Sending emails to users within the rule will be skipped completly.
This is useful if users would have email addresses, that would normally trigger an email to them.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
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
