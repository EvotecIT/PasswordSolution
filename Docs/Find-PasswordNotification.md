---
external help file: PasswordSolution-help.xml
Module Name: PasswordSolution
online version: https://github.com/EvotecIT/PasswordSolution
schema: 2.0.0
---
# Find-PasswordNotification
## SYNOPSIS
Searches thru XML logs created by Password Solution

## SYNTAX
### __AllParameterSets
```powershell
Find-PasswordNotification [-SearchPath] <string> [-Manager] [-FailedOnly] [<CommonParameters>]
```

## DESCRIPTION
Searches thru XML logs created by Password Solution

## EXAMPLES

### EXAMPLE 1
```powershell
PS > Find-PasswordNotification -SearchPath $PSScriptRoot\Search\SearchLog.xml | Format-Table
```


### EXAMPLE 2
```powershell
PS > Find-PasswordNotification -SearchPath "$PSScriptRoot\Search\SearchLog_2021-06.xml" -Manager | Format-Table
```


## PARAMETERS

### -FailedOnly
Return the latest failed user/rule deliveries, excluding failures resolved by a later successful send.
Legacy entries with empty status are included. This option cannot be combined with Manager.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Manager
Search thru manager escalations

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -SearchPath
Path to file where the XML log is located

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `Notification history records, including recipient, status and error details.`

## RELATED LINKS

- None
