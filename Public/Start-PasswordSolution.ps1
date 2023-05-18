function Start-PasswordSolution {
    <#
    .SYNOPSIS
    Starts Password Expiry Notifications for the whole forest

    .DESCRIPTION
    Starts Password Expiry Notifications for the whole forest

    .PARAMETER EmailParameters
    Parameters for Email. Uses Mailozaurr splatting behind the scenes, so it supports all options that Mailozaurr does.

    .PARAMETER OverwriteEmailProperty
    Property responsible for overwriting the default email field in Active Directory. Useful when the password notification has to go somewhere else than users email address.

    .PARAMETER UserSection
    Parameter description

    .PARAMETER ManagerSection
    Parameter description

    .PARAMETER SecuritySection
    Parameter description

    .PARAMETER AdminSection
    Parameter description

    .PARAMETER Rules
    Parameter description

    .PARAMETER TemplatePreExpiry
    Parameter description

    .PARAMETER TemplatePreExpirySubject
    Parameter description

    .PARAMETER TemplatePostExpiry
    Parameter description

    .PARAMETER TemplatePostExpirySubject
    Parameter description

    .PARAMETER TemplateManager
    Parameter description

    .PARAMETER TemplateManagerSubject
    Parameter description

    .PARAMETER TemplateSecurity
    Parameter description

    .PARAMETER TemplateSecuritySubject
    Parameter description

    .PARAMETER TemplateManagerNotCompliant
    Parameter description

    .PARAMETER TemplateManagerNotCompliantSubject
    Parameter description

    .PARAMETER TemplateAdmin
    Parameter description

    .PARAMETER TemplateAdminSubject
    Parameter description

    .PARAMETER Logging
    Parameter description

    .PARAMETER HTMLReports
    Parameter description

    .PARAMETER SearchPath
    Parameter description

    .EXAMPLE
    An example

    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary] $EmailParameters,
        [string] $OverwriteEmailProperty,
        [Parameter(Mandatory)][System.Collections.IDictionary] $UserSection,
        [Parameter(Mandatory)][System.Collections.IDictionary] $ManagerSection,
        [Parameter(Mandatory)][System.Collections.IDictionary] $SecuritySection,
        [Parameter(Mandatory)][System.Collections.IDictionary] $AdminSection,
        [Parameter(Mandatory)][Array] $Rules,
        [scriptblock] $TemplatePreExpiry,
        [string] $TemplatePreExpirySubject,
        [scriptblock] $TemplatePostExpiry,
        [string] $TemplatePostExpirySubject,
        [Parameter(Mandatory)][scriptblock] $TemplateManager,
        [Parameter(Mandatory)][string] $TemplateManagerSubject,
        [Parameter(Mandatory)][scriptblock] $TemplateSecurity,
        [Parameter(Mandatory)][string] $TemplateSecuritySubject,
        [Parameter(Mandatory)][scriptblock] $TemplateManagerNotCompliant,
        [Parameter(Mandatory)][string] $TemplateManagerNotCompliantSubject,
        [Parameter(Mandatory)][scriptblock] $TemplateAdmin,
        [Parameter(Mandatory)][string] $TemplateAdminSubject,
        [Parameter()][System.Collections.IDictionary] $Logging = @{},
        [Array] $HTMLReports,
        [string] $SearchPath
    )
    $TimeStart = Start-TimeLog
    $Script:Reporting = [ordered] @{}
    $Script:Reporting['Version'] = Get-GitHubVersion -Cmdlet 'Start-PasswordSolution' -RepositoryOwner 'evotecit' -RepositoryName 'PasswordSolution'

    Write-Color -Text '[i]', "[PasswordSolution] ", 'Version', ' [Informative] ', $Script:Reporting['Version'] -Color Yellow, DarkGray, Yellow, DarkGray, Magenta

    $TodayDate = Get-Date
    $Today = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    Set-LoggingCapabilities -LogPath $Logging.LogFile -LogMaximum $Logging.LogMaximum -ShowTime:$Logging.ShowTime -TimeFormat $Logging.TimeFormat
    # since the first entry didn't go to log file, this will
    Write-Color -Text '[i]', "[PasswordSolution] ", 'Version', ' [Informative] ', $Script:Reporting['Version'] -Color Yellow, DarkGray, Yellow, DarkGray, Magenta -NoConsoleOutput

    $SummarySearch = Import-SearchInformation -SearchPath $SearchPath

    $Summary = [ordered] @{}
    $Summary['Notify'] = [ordered] @{}
    $Summary['NotifyManager'] = [ordered] @{}
    $Summary['NotifySecurity'] = [ordered] @{}
    $Summary['Rules'] = [ordered] @{}


    $AllSkipped = [ordered] @{}
    $Locations = [ordered] @{}

    Write-Color -Text "[i]", " Starting process to find expiring users" -Color Yellow, White, Green, White, Green, White, Green, White
    $CachedUsers = Find-Password -AsHashTable -OverwriteEmailProperty $OverwriteEmailProperty
    foreach ($Rule in $Rules) {
        $SplatProcessingRule = [ordered] @{
            Rule        = $Rule
            Summary     = $Summary
            CachedUsers = $CachedUsers
            AllSkipped  = $AllSkipped
            Locations   = $Locations
            Loggin      = $Logging
            TodayDate   = $TodayDate
        }
        Invoke-PasswordRuleProcessing @SplatProcessingRule
    }

    $SplatUserNotifications = [ordered] @{
        UserSection               = $UserSection
        Summary                   = $Summary
        Logging                   = $Logging
        TemplatePreExpiry         = $TemplatePreExpiry
        TemplatePreExpirySubject  = $TemplatePreExpirySubject
        TemplatePostExpiry        = $TemplatePostExpiry
        TemplatePostExpirySubject = $TemplatePostExpirySubject
        EmailParameter            = $EmailParameters
    }

    [Array] $SummaryUsersEmails = Send-PasswordUserNofifications @SplatUserNotifications

    $SplatManagerNotifications = [ordered] @{
        ManagerSection                     = $ManagerSection
        Summary                            = $Summary
        CachedUsers                        = $CachedUsers
        TemplateManager                    = $TemplateManager
        TemplateManagerSubject             = $TemplateManagerSubject
        TemplateManagerNotCompliant        = $TemplateManagerNotCompliant
        TemplateManagerNotCompliantSubject = $TemplateManagerNotCompliantSubject
        EmailParameters                    = $EmailParameters
        Loggin                             = $Logging
    }
    [Array] $SummaryManagersEmails = Send-PasswordManagerNofifications @SplatManagerNotifications

    $SplatSecurityNotifications = [ordered] @{
        SecuritySection         = $SecuritySection
        Summary                 = $Summary
        TemplateSecurity        = $TemplateSecurity
        TemplateSecuritySubject = $TemplateSecuritySubject
        Logging                 = $Logging
    }
    [Array] $SummaryEscalationEmails = Send-PasswordSecurityNotifications @SplatSecurityNotifications

    $TimeEnd = Stop-TimeLog -Time $TimeStart -Option OneLiner

    $HtmlAttachments = [System.Collections.Generic.List[string]]::new()

    foreach ($Report in $HTMLReports) {
        if ($Report.Enable) {
            $ReportSettings = @{
                Report                  = $Report
                EmailParameters         = $EmailParameters
                Logging                 = $Logging
                FilePath                = $FilePath
                SearchPath              = $SearchPath
                Rules                   = $Rules
                UserSection             = $UserSection
                ManagerSection          = $ManagerSection
                SecuritySection         = $SecuritySection
                AdminSection            = $AdminSection
                CachedUsers             = $CachedUsers
                Summary                 = $Summary
                SummaryUsersEmails      = $SummaryUsersEmails
                SummaryManagersEmails   = $SummaryManagersEmails
                SummaryEscalationEmails = $SummaryEscalationEmails
                SummarySearch           = $SummarySearch
                Locations               = $Locations
                AllSkipped              = $AllSkipped
            }
            New-HTMLReport @ReportSettings

            if ($Report.AttachToEmail) {
                if (Test-Path -LiteralPath $Report.FilePath) {
                    $HtmlAttachments.Add($Report.FilePath)
                } else {
                    Write-Color -Text "[w] HTML report ", $Report.FilePath, " does not exist! Probably a temporary path was used. " -Color DarkYellow, Red, DarkYellow
                }
            }
        }
    }

    Export-SearchInformation -SearchPath $SearchPath -SummarySearch $SummarySearch -Today $Today

    $AdminSplat = [ordered] @{
        AdminSection         = $AdminSection
        TemplateAdmin        = $TemplateAdmin
        TemplateAdminSubject = $TemplateAdminSubject
        TimeEnd              = $TimeEnd
        EmailParameters      = $EmailParameters
        HtmlAttachment       = $HtmlAttachments
    }
    Send-PasswordAdminNotifications @AdminSplat
}