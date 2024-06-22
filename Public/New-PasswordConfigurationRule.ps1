function New-PasswordConfigurationRule {
    <#
    .SYNOPSIS
    Short description

    .DESCRIPTION
    Long description

    .PARAMETER ReminderConfiguration
    Parameter description

    .PARAMETER Name
    Parameter description

    .PARAMETER Enable
    Parameter description

    .PARAMETER IncludeExpiring
    Parameter description

    .PARAMETER IncludePasswordNeverExpires
    Parameter description

    .PARAMETER PasswordNeverExpiresDays
    Parameter description

    .PARAMETER IncludeName
    Include user in rule if any of the properties match the value of Name in the properties defined in IncludeNameProperties

    .PARAMETER IncludeNameProperties
    Include user in rule if any of the properties match the value as defined in IncludeName

    .PARAMETER ExcludeName
    Exclude user from rule if any of the properties match the value of Name in the properties defined in ExcludeNameProperties

    .PARAMETER ExcludeNameProperties
    Exclude user from rule if any of the properties match the value as defined in ExcludeName

    .PARAMETER IncludeOU
    Parameter description

    .PARAMETER ExcludeOU
    Parameter description

    .PARAMETER IncludeGroup
    Parameter description

    .PARAMETER ExcludeGroup
    Parameter description

    .PARAMETER ReminderDays
    Parameter description

    .PARAMETER ManagerReminder
    Parameter description

    .PARAMETER ManagerNotCompliant
    Parameter description

    .PARAMETER ManagerNotCompliantDisplayName
    Parameter description

    .PARAMETER ManagerNotCompliantEmailAddress
    Parameter description

    .PARAMETER ManagerNotCompliantDisabled
    Parameter description

    .PARAMETER ManagerNotCompliantMissing
    Parameter description

    .PARAMETER ManagerNotCompliantMissingEmail
    Parameter description

    .PARAMETER ManagerNotCompliantLastLogonDays
    Parameter description

    .PARAMETER SecurityEscalation
    Parameter description

    .PARAMETER SecurityEscalationDisplayName
    Parameter description

    .PARAMETER SecurityEscalationEmailAddress
    Parameter description

    .PARAMETER OverwriteEmailProperty
    Parameter description

    .PARAMETER OverwriteManagerProperty
    Parameter description

    .PARAMETER OverwriteEmailFromExternalUsers
    Allow to overwrite email from external users for specific rule

    .PARAMETER ProcessManagersOnly
    This parameters is used to process users, but only managers will be notified.
    Sending emails to users within the rule will be skipped completly.
    This is useful if users would have email addresses, that would normally trigger an email to them.

    .EXAMPLE
    An example

    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [scriptblock] $ReminderConfiguration,
        [parameter(Mandatory)][string] $Name,
        [switch] $Enable,
        [switch] $IncludeExpiring,
        [switch] $IncludePasswordNeverExpires,
        [nullable[int]]$PasswordNeverExpiresDays,
        [string[]] $IncludeNameProperties,
        [string[]] $IncludeName,

        [string[]] $ExcludeNameProperties,
        [string[]] $ExcludeName,

        [string[]] $IncludeOU,
        [string[]] $ExcludeOU,
        [string[]] $IncludeGroup,
        [string[]] $ExcludeGroup,

        [parameter(Mandatory)][alias('ExpirationDays', 'Days')][Array] $ReminderDays,

        [switch] $ManagerReminder,

        [switch] $ManagerNotCompliant,
        [string] $ManagerNotCompliantDisplayName,
        [string] $ManagerNotCompliantEmailAddress,

        [switch] $ManagerNotCompliantDisabled,
        [switch] $ManagerNotCompliantMissing,
        [switch]$ManagerNotCompliantMissingEmail,
        [nullable[int]] $ManagerNotCompliantLastLogonDays,

        [switch] $SecurityEscalation,
        [string] $SecurityEscalationDisplayName,
        [string] $SecurityEscalationEmailAddress,

        [string] $OverwriteEmailProperty,
        [string] $OverwriteManagerProperty,

        [switch] $ProcessManagersOnly,

        [switch] $OverwriteEmailFromExternalUsers

    )

    $Output = [ordered] @{
        Name                            = $Name
        Enable                          = $Enable.IsPresent
        IncludeExpiring                 = $IncludeExpiring.IsPresent
        IncludePasswordNeverExpires     = $IncludePasswordNeverExpires.IsPresent
        Reminders                       = $ReminderDays
        PasswordNeverExpiresDays        = $PasswordNeverExpiresDays
        IncludeNameProperties           = $IncludeNameProperties
        IncludeName                     = $IncludeName
        IncludeOU                       = $IncludeOU
        ExcludeOU                       = $ExcludeOU
        SendToManager                   = [ordered] @{}

        ProcessManagersOnly             = $ProcessManagersOnly.IsPresent

        OverwriteEmailProperty          = $OverwriteEmailProperty
        # properties to overwrite manager based on different field
        OverwriteManagerProperty        = $OverwriteManagerProperty

        OverwriteEmailFromExternalUsers = $OverwriteEmailFromExternalUsers.IsPresent
    }
    $Output.SendToManager['Manager'] = [ordered] @{
        Enable    = $false
        Reminders = [ordered] @{}
    }
    $Output.SendToManager['ManagerNotCompliant'] = [ordered] @{
        Enable        = $false
        Manager       = [ordered] @{
            DisplayName  = $ManagerNotCompliantDisplayName
            EmailAddress = $ManagerNotCompliantEmailAddress
        }
        Disabled      = $ManagerNotCompliantDisabled
        Missing       = $ManagerNotCompliantMissing
        MissingEmail  = $ManagerNotCompliantMissingEmail
        LastLogon     = if ($PSBoundParameters.ContainsKey('ManagerNotCompliantLastLogonDays')) { $true } else { $false }
        LastLogonDays = $ManagerNotCompliantLastLogonDays
        Reminders     = [ordered] @{ }
    }
    $Output.SendToManager['SecurityEscalation'] = [ordered] @{
        Enable    = $false
        Manager   = [ordered] @{
            DisplayName  = $SecurityEscalationDisplayName
            EmailAddress = $SecurityEscalationEmailAddress
        }
        Reminders = [ordered] @{}
    }
    if ($ManagerReminder) {
        $Output.SendToManager['Manager'].Enable = $true
    }
    if ($ManagerNotCompliant) {
        $Output.SendToManager['ManagerNotCompliant'].Enable = $true
    }
    if ($SecurityEscalation) {
        $Output.SendToManager['SecurityEscalation'].Enable = $true
    }
    if ($ReminderConfiguration) {
        try {
            $RemindersExecution = & $ReminderConfiguration
        } catch {
            Write-Color -Text "[e]", " Processing rule ", $Output.Name, " failed because of error: ", $_.Exception.Message -Color Yellow, White, Red
            return
        }
        foreach ($Reminder in $RemindersExecution) {
            if ($Reminder.Type -eq 'Manager') {
                foreach ($ReminderReminders in $Reminder.Reminders) {
                    $Output.SendToManager['Manager'].Reminders += $ReminderReminders
                }
            } elseif ($Reminder.Type -eq 'ManagerNotCompliant') {
                foreach ($ReminderReminders in $Reminder.Reminders) {
                    $Output.SendToManager['ManagerNotCompliant'].Reminders += $ReminderReminders
                }
            } elseif ($Reminder.Type -eq 'Security') {
                foreach ($ReminderReminders in $Reminder.Reminders) {
                    $Output.SendToManager['SecurityEscalation'].Reminders += $ReminderReminders
                }
            } else {
                # Should not happen
                throw "Invalid reminder type: $($Reminder.Type)"
            }
        }
    }

    Remove-EmptyValue -Hashtable $Output -Recursive -Rerun 2
    $Configuration = [ordered] @{
        Type     = 'PasswordConfigurationRule'
        Settings = $Output
    }
    $Configuration
}