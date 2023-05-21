function New-PasswordConfigurationRule {
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
        [string] $SecurityEscalationEmailAddress
    )

    $Output = [ordered] @{
        Name                        = $Name
        Enable                      = $Enable.IsPresent
        IncludeExpiring             = $IncludeExpiring.IsPresent
        IncludePasswordNeverExpires = $IncludePasswordNeverExpires.IsPresent
        Reminders                   = $ReminderDays
        PasswordNeverExpiresDays    = $PasswordNeverExpiresDays
        IncludeNameProperties       = $IncludeNameProperties
        IncludeName                 = $IncludeName
        IncludeOU                   = $IncludeOU
        ExcludeOU                   = $ExcludeOU
        SendToManager               = [ordered] @{}
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