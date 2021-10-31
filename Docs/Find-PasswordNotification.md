---
external help file: PasswordSolution-help.xml
Module Name: PasswordSolution
online version:
schema: 2.0.0
---

# Find-PasswordNotification

## SYNOPSIS
Searches thru XML logs created by Password Solution

## SYNTAX

```
Find-PasswordNotification [-SearchPath] <String> [-Manager] [<CommonParameters>]
```

## DESCRIPTION
Searches thru XML logs created by Password Solution

## EXAMPLES

### EXAMPLE 1
```
Find-PasswordNotification -SearchPath $PSScriptRoot\Search\SearchLog.xml | Format-Table
```

### EXAMPLE 2
```
Find-PasswordNotification -SearchPath "$PSScriptRoot\Search\SearchLog_2021-06.xml" -Manager | Format-Table
```

## PARAMETERS

### -SearchPath
Path to file where the XML log is located

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Manager
Search thru manager escalations

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

## OUTPUTS

## NOTES
General notes

## RELATED LINKS
