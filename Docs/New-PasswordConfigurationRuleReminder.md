---
external help file: PasswordSolution-help.xml
Module Name: PasswordSolution
online version:
schema: 2.0.0
---

# New-PasswordConfigurationRuleReminder

## SYNOPSIS
{{ Fill in the Synopsis }}

## SYNTAX

### DayOfMonth
```
New-PasswordConfigurationRuleReminder -Type <String> -ReminderOn <Int32[]> -DayOfMonth <Int32[]>
 [-ComparisonType <String>] [<CommonParameters>]
```

### DayOfWeek
```
New-PasswordConfigurationRuleReminder -Type <String> -ReminderOn <Int32[]> -DayOfWeek <String[]>
 [-ComparisonType <String>] [<CommonParameters>]
```

### Daily
```
New-PasswordConfigurationRuleReminder -Type <String> -ReminderOn <Int32[]> [-ComparisonType <String>]
 [<CommonParameters>]
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

### -ComparisonType
{{ Fill ComparisonType Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:
Accepted values: lt, gt, eq, in

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DayOfMonth
{{ Fill DayOfMonth Description }}

```yaml
Type: Int32[]
Parameter Sets: DayOfMonth
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DayOfWeek
{{ Fill DayOfWeek Description }}

```yaml
Type: String[]
Parameter Sets: DayOfWeek
Aliases:
Accepted values: Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReminderOn
{{ Fill ReminderOn Description }}

```yaml
Type: Int32[]
Parameter Sets: (All)
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Type
{{ Fill Type Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:
Accepted values: User, Manager, ManagerNotCompliant, Security

Required: True
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
