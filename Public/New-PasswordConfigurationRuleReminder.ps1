function New-PasswordConfigurationRuleReminder {
    [CmdletBinding(DefaultParameterSetName = 'Daily')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Daily')]
        [Parameter(Mandatory, ParameterSetName = 'DayOfWeek')]
        [Parameter(Mandatory, ParameterSetName = 'DayOfMonth')]
        [ValidateSet('Manager', 'ManagerNotCompliant', 'Security')][string] $Type,

        [Parameter(Mandatory, ParameterSetName = 'Daily')]
        [Parameter(Mandatory, ParameterSetName = 'DayOfWeek')]
        [Parameter(Mandatory, ParameterSetName = 'DayOfMonth')]
        [alias('ConditionDays', 'Days')][Array] $ExpirationDays,

        [Parameter(Mandatory, ParameterSetName = 'DayOfWeek')]
        [ValidateSet(
            'Monday',
            'Tuesday',
            'Wednesday',
            'Thursday',
            'Friday',
            'Saturday',
            'Sunday'
        )][Array] $DayOfWeek,

        [Parameter(Mandatory, ParameterSetName = 'DayOfMonth')]
        [Array] $DayOfMonth,

        [Parameter(ParameterSetName = 'Daily')]
        [Parameter(ParameterSetName = 'DayOfWeek')]
        [Parameter(ParameterSetName = 'DayOfMonth')]
        [ValidateSet('lt', 'gt', 'eq', 'in')][string] $ComparisonType = 'eq'
    )
    if ($ComparisonType -in 'eq', 'lt', 'gt') {
        if ($ExpirationDays.Count -gt 1) {
            throw "Only one number for 'ExpirationDays' can be specified when using comparison types 'eq', 'lt', and 'gt'."
        } else {
            $ExpirationDaysToUse = $ExpirationDays[0]
        }
    } else {
        $ExpirationDaysToUse = $ExpirationDays
    }

    if ($PSCmdlet.ParameterSetName -eq 'Daily') {
        $Reminders = [ordered] @{
            Type      = $Type
            Reminders = @{
                Default = [ordered] @{
                    Enable = $true
                }
            }
        }
    } elseif ($PSCmdlet.ParameterSetName -eq 'DayOfWeek') {
        $Reminders = [ordered] @{
            Type      = $Type
            Reminders = @{
                OnDay = [ordered] @{
                    Enable         = $true
                    Reminder       = $ExpirationDaysToUse
                    ComparisonType = $ComparisonType
                    Days           = $DayOfWeek
                }
            }
        }
    } elseif ($PSCmdlet.ParameterSetName -eq 'DayOfMonth') {
        $Reminders = [ordered] @{
            Type      = $Type
            Reminders = @{
                OnDayOfMonth = [ordered] @{
                    Enable         = $true
                    Reminder       = $ExpirationDaysToUse
                    ComparisonType = $ComparisonType
                    Days           = $DayOfMonth
                }
            }
        }
    }
    $Reminders
}