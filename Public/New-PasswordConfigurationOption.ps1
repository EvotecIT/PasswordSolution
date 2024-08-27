function New-PasswordConfigurationOption {
    <#
    .SYNOPSIS
    Provides a way to create a PasswordConfigurationOption object.

    .DESCRIPTION
    This function provides a way to create a PasswordConfigurationOption object.
    The object is used to store configuration options for the Password Solution module.

    .PARAMETER ShowTime
    Show time in the console output. If not provided, time will not be shown.
    Time in the log file is always shown.

    .PARAMETER LogFile
    File path to the log file. If not provided, there will be no logging to file

    .PARAMETER TimeFormat
    Time format used in the logging functionality.

    .PARAMETER LogMaximum
    Maximum number of log files to keep. Default is 0 (unlimited).
    Once the number of log files exceeds the limit, the oldest log files will be deleted.

    .PARAMETER NotifyOnSkipUserManagerOnly
    Provides a way to control output to screen for SkipUserManagerOnly.

    .PARAMETER NotifyOnSecuritySend
    Provides a way to control output to screen for SecuritySend.

    .PARAMETER NotifyOnManagerSend
    Provides a way to control output to screen for ManagerSend.

    .PARAMETER NotifyOnUserSend
    Provides a way to control output to screen for UserSend.

    .PARAMETER NotifyOnUserMatchingRule
    Provides a way to control output to screen for UserMatchingRule.

    .PARAMETER NotifyOnUserDaysToExpireNull
    Provides a way to control output to screen for UserDaysToExpireNull.

    .PARAMETER NotifyOnUserMatchingRuleForManager
    Provides a way to control output to screen for UserMatchingRuleForManager.

    .PARAMETER NotifyOnUserMatchingRuleForManagerButNotCompliant
    Provides a way to control output to screen for UserMatchingRuleForManagerButNotCompliant.

    .PARAMETER SearchPath
    Path to XML file that will be used for storing search results.

    .PARAMETER EmailDateFormat
    Parameter description

    .PARAMETER EmailDateFormatUTCConversion
    Parameter description

    .PARAMETER OverwriteEmailProperty
    Parameter description

    .PARAMETER OverwriteManagerProperty
    Parameter description

    .PARAMETER FilterOrganizationalUnit
    Provides a way to filter users by Organizational Unit limiting the scope of the search.
    The search is performed using 'like' operator, so you can use wildcards if needed.

    .EXAMPLE
    $Options = @{
        # Logging to file and to screen
        ShowTime                     = $true
        LogFile                      = "$PSScriptRoot\Logs\PasswordSolution_$(($Date).ToString('yyyy-MM-dd_HH_mm_ss')).log"
        TimeFormat                   = "yyyy-MM-dd HH:mm:ss"
        LogMaximum                   = 365
        NotifyOnSkipUserManagerOnly  = $false
        NotifyOnSecuritySend         = $true
        NotifyOnManagerSend          = $true
        NotifyOnUserSend             = $true
        NotifyOnUserMatchingRule     = $false
        NotifyOnUserDaysToExpireNull = $false
        SearchPath                   = "$PSScriptRoot\Search\SearchLog_$((Get-Date).ToString('yyyy-MM')).xml"
        EmailDateFormat              = "yyyy-MM-dd"
        EmailDateFormatUTCConversion = $true
        FilterOrganizationalUnit     = @(
            "*OU=Accounts,OU=Administration,DC=ad,DC=evotec,DC=xyz"
            "*OU=Administration,DC=ad,DC=evotec,DC=xyz"
        )
    }
    New-PasswordConfigurationOption @Options

    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [switch] $ShowTime                     , #= $true
        [string] $LogFile                      , #= "$PSScriptRoot\Logs\PasswordSolution_$(($Date).ToString('yyyy-MM-dd_HH_mm_ss')).log"
        [string] $TimeFormat                   , #= "yyyy-MM-dd HH:mm:ss"
        [int] $LogMaximum                   , #= 365
        [switch] $NotifyOnSkipUserManagerOnly  , #= $false
        [switch] $NotifyOnSecuritySend         , #= $true
        [switch] $NotifyOnManagerSend          , #= $true
        [switch] $NotifyOnUserSend             , #= $true
        [switch] $NotifyOnUserMatchingRule     , #= $true
        [switch] $NotifyOnUserDaysToExpireNull , #= $true
        [switch] $NotifyOnUserMatchingRuleForManager,
        [switch] $NotifyOnUserMatchingRuleForManagerButNotCompliant,
        [string] $SearchPath,
        [string] $EmailDateFormat,
        [switch] $EmailDateFormatUTCConversion,
        [string] $OverwriteEmailProperty,
        [string] $OverwriteManagerProperty,
        [string[]] $FilterOrganizationalUnit
    )

    $Output = [ordered] @{
        Type     = "PasswordConfigurationOption"
        Settings = [ordered] @{
            ShowTime                                          = $ShowTime.IsPresent
            LogFile                                           = $LogFile
            TimeFormat                                        = $TimeFormat
            LogMaximum                                        = $LogMaximum
            NotifyOnSkipUserManagerOnly                       = $NotifyOnSkipUserManagerOnly.IsPresent
            NotifyOnSecuritySend                              = $NotifyOnSecuritySend.IsPresent
            NotifyOnManagerSend                               = $NotifyOnManagerSend.IsPresent
            NotifyOnUserSend                                  = $NotifyOnUserSend.IsPresent
            NotifyOnUserMatchingRule                          = $NotifyOnUserMatchingRule.IsPresent
            NotifyOnUserDaysToExpireNull                      = $NotifyOnUserDaysToExpireNull.IsPresent
            NotifyOnUserMatchingRuleForManager                = $NotifyOnUserMatchingRuleForManager.IsPresent
            NotifyOnUserMatchingRuleForManagerButNotCompliant = $NotifyOnUserMatchingRuleForManagerButNotCompliant.IsPresent
            SearchPath                                        = $SearchPath
            # conversion for DateExpiry/PasswordLastSet only
            EmailDateFormat                                   = $EmailDateFormat
            EmailDateFormatUTCConversion                      = $EmailDateFormatUTCConversion.IsPresent
            # email property conversion (global)
            OverwriteEmailProperty                            = $OverwriteEmailProperty
            # manager property conversion (global)
            OverwriteManagerProperty                          = $OverwriteManagerProperty
            # filtering
            FilterOrganizationalUnit                          = $FilterOrganizationalUnit
        }
    }
    Remove-EmptyValue -Hashtable $Output.Settings
    $Output
}