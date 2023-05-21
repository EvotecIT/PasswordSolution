function New-PasswordConfigurationReport {
    [CmdletBinding()]
    param(
        [switch] $Enable                , #= $true
        [switch] $ShowHTML              , #= $true
        [string] $Title                 , #= "Password Solution Summary"
        [switch] $Online                , #= $true
        [switch] $DisableWarnings       , #= $true
        [switch] $ShowConfiguration     , #= $true
        [switch] $ShowAllUsers          , #= $true
        [switch] $ShowRules             , #= $true
        [switch] $ShowUsersSent         , #= $true
        [switch] $ShowManagersSent      , #= $true
        [switch] $ShowEscalationSent    , #= $true
        [switch] $ShowSkippedUsers      , #= $true
        [switch] $ShowSkippedLocations  , #= $true
        [switch] $ShowSearchUsers       , #= $true
        [switch] $ShowSearchManagers    , #= $true
        [switch] $ShowSearchEscalations , #= $true
        [string] $FilePath              , #= "$PSScriptRoot\Reporting\PasswordSolution_$(($Date).ToString('yyyy-MM-dd_HH_mm_ss')).html"
        [switch] $AttachToEmail           #= $true
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