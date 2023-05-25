function New-PasswordConfigurationReport {
    <#
    .SYNOPSIS
    Provides HTML report configuration for Password Notifications in Password Solution.

    .DESCRIPTION
    Provides HTML report configuration for Password Notifications in Password Solution.
    The New-PasswordConfigurationReport function generates configuration for HTML report.

    .PARAMETER Enable
    Specifies whether to enable the report generation. The default value is $false.

    .PARAMETER ShowHTML
    Specifies whether to display the report in HTML format right after it's generated in default browser. The default value is $false.

    .PARAMETER Title
    Specifies the title of the report. The default value is "Password Solution Summary".

    .PARAMETER Online
    Specifies whether to generate the report using CDN for CSS and JS scripts, or use it locally.
    It doesn't require internet connectivity during generation.
    Makes the final output 3MB smaller. The default value is $false.

    .PARAMETER DisableWarnings
    Specifies whether to disable warning messages during report generation. The default value is $false.

    .PARAMETER ShowConfiguration
    Specifies whether to display the current Password Solution configuration settings. The default value is $false.

    .PARAMETER ShowAllUsers
    Specifies whether to display information about all user accounts. The default value is $false.

    .PARAMETER ShowRules
    Specifies whether to display information from the rules. The default value is $false.

    .PARAMETER ShowUsersSent
    Specifies whether to display information about users who have received (or not) password expiry notifications. The default value is $false.

    .PARAMETER ShowManagersSent
    Specifies whether to display information about managers who have received password expiry notifications. The default value is $false.

    .PARAMETER ShowEscalationSent
    Specifies whether to display information about escalation contacts who have received password expiry notifications. The default value is $false.

    .PARAMETER ShowSkippedUsers
    Specifies whether to display information about users who were during password expiry notifications because of inability to asses their expiration date. The default value is $false.

    .PARAMETER ShowSkippedLocations
    Specifies whether to display information about locations where skipped users are located. The default value is $false.

    .PARAMETER ShowSearchUsers
    Specifies whether to display information for searching who got password expiry notifications. The default value is $false.

    .PARAMETER ShowSearchManagers
    Specifies whether to display information for searching who got password expiry notifications and for which accounts from managers. The default value is $false.

    .PARAMETER ShowSearchEscalations
    Specifies whether to display information for searching who got password escalation notifications and what's the status of that message. The default value is $false.

    .PARAMETER FilePath
    Specifies the file path for the report

    .PARAMETER AttachToEmail
    Specifies whether to attach the report to an administrative email. The default value is $false.

    .PARAMETER NestedRules
    Specifies whether to display nested password rules.
    Each rule has it's own tab with output.
    Having many rules and all other settings enabled can result in a very long list of tabs that's hard to navigate.
    This setting forces separate tab for all rules.
    The default value is $false.

    .OUTPUTS
    The function returns an ordered dictionary that contains the report settings.

    .EXAMPLE
    New-PasswordConfigurationReport -ShowHTML -Title "Password Configuration Report" -FilePath "C:\Reports\PasswordReport.html"

    .EXAMPLE
    $Date = Get-Date
    $Report = [ordered] @{
        Enable                = $true
        ShowHTML              = $true
        Title                 = "Password Solution Summary"
        Online                = $true
        DisableWarnings       = $true
        ShowConfiguration     = $true
        ShowAllUsers          = $true
        ShowRules             = $true
        ShowUsersSent         = $true
        ShowManagersSent      = $true
        ShowEscalationSent    = $true
        ShowSkippedUsers      = $true
        ShowSkippedLocations  = $true
        ShowSearchUsers       = $true
        ShowSearchManagers    = $true
        ShowSearchEscalations = $true
        NestedRules           = $false
        FilePath              = "$PSScriptRoot\Reporting\PasswordSolution_$(($Date).ToString('yyyy-MM-dd_HH_mm_ss')).html"
        AttachToEmail         = $true
    }
    New-PasswordConfigurationReport @Report

    #>
    [CmdletBinding()]
    param(
        [switch] $Enable,
        [switch] $ShowHTML,
        [string] $Title,
        [switch] $Online,
        [switch] $DisableWarnings,
        [switch] $ShowConfiguration,
        [switch] $ShowAllUsers,
        [switch] $ShowRules,
        [switch] $ShowUsersSent,
        [switch] $ShowManagersSent,
        [switch] $ShowEscalationSent,
        [switch] $ShowSkippedUsers,
        [switch] $ShowSkippedLocations,
        [switch] $ShowSearchUsers,
        [switch] $ShowSearchManagers,
        [switch] $ShowSearchEscalations ,
        [string] $FilePath,
        [switch] $AttachToEmail,
        [switch] $NestedRules
    )

    $Output = [ordered] @{
        Type     = "PasswordConfigurationReport"
        Settings = [ordered] @{
            Enable                = $Enable.IsPresent
            ShowHTML              = $ShowHTML.IsPresent
            Title                 = $Title
            Online                = $Online.IsPresent
            DisableWarnings       = $DisableWarnings.IsPresent
            ShowConfiguration     = $ShowConfiguration.IsPresent
            ShowAllUsers          = $ShowAllUsers.IsPresent
            ShowRules             = $ShowRules.IsPresent
            ShowUsersSent         = $ShowUsersSent.IsPresent
            ShowManagersSent      = $ShowManagersSent.IsPresent
            ShowEscalationSent    = $ShowEscalationSent.IsPresent
            ShowSkippedUsers      = $ShowSkippedUsers.IsPresent
            ShowSkippedLocations  = $ShowSkippedLocations.IsPresent
            ShowSearchUsers       = $ShowSearchUsers.IsPresent
            ShowSearchManagers    = $ShowSearchManagers.IsPresent
            ShowSearchEscalations = $ShowSearchEscalations.IsPresent
            FilePath              = $FilePath
            AttachToEmail         = $AttachToEmail.IsPresent
        }
    }
    $Output
}