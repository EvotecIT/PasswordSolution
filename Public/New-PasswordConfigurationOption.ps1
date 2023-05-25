function New-PasswordConfigurationOption {
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
        [string] $SearchPath,
        [string] $EmailDateFormat,
        [switch] $EmailDateFormatUTCConversion,
        [string] $OverwriteEmailProperty,
        [string] $OverwriteManagerProperty
    )

    $Output = [ordered] @{
        Type     = "PasswordConfigurationOption"
        Settings = [ordered] @{
            ShowTime                     = $ShowTime.IsPresent
            LogFile                      = $LogFile
            TimeFormat                   = $TimeFormat
            LogMaximum                   = $LogMaximum
            NotifyOnSkipUserManagerOnly  = $NotifyOnSkipUserManagerOnly.IsPresent
            NotifyOnSecuritySend         = $NotifyOnSecuritySend.IsPresent
            NotifyOnManagerSend          = $NotifyOnManagerSend.IsPresent
            NotifyOnUserSend             = $NotifyOnUserSend.IsPresent
            NotifyOnUserMatchingRule     = $NotifyOnUserMatchingRule.IsPresent
            NotifyOnUserDaysToExpireNull = $NotifyOnUserDaysToExpireNull.IsPresent
            SearchPath                   = $SearchPath
            # conversion for DateExpiry/PasswordLastSet only
            EmailDateFormat              = $EmailDateFormat
            EmailDateFormatUTCConversion = $EmailDateFormatUTCConversion.IsPresent
            # email property conversion (global)
            OverwriteEmailProperty       = $OverwriteEmailProperty
            # manager property conversion (global)
            OverwriteManagerProperty     = $OverwriteManagerProperty
        }
    }
    Remove-EmptyValue -Hashtable $Output.Settings
    $Output
}