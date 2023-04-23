function Show-PasswordQuality {
    <#
    .SYNOPSIS
    Creates an HTML report showing password quality for all user objects in Active Directory.

    .DESCRIPTION
    Creates an HTML report showing password quality for all user objects in Active Directory.
    This comman utilizes DSInternals PowerShell module to get the data.
    Then it uses PSWriteHTML to create nice looking report.

    .PARAMETER FilePath
    Path to the file where report will be saved.

    .PARAMETER DontShow
    If specified, report will not be opened in a browser.

    .PARAMETER Online
    If specified report will use CDN for JS and CSS files.
    If not specified, it will merge all CSS and JS files into one HTML file.
    This makes the file at least 3MB bigger, even if there is very small amount of data.
    Keep in mind that this report can be created without internet access,
    just that opening it in a browser with -Online switch will require internet access.

    .PARAMETER WeakPasswords
    List of weak passwords that should be checked for.
    Provide a list of common passwords that you want to check for, and that your users may have used.

    .PARAMETER SeparateDuplicateGroups
    If specified, report will show duplicate groups separately, one group per tab.

    .EXAMPLE
    Show-PasswordQuality -FilePath $PSScriptRoot\Reporting\PasswordQuality.html -Online -WeakPasswords "Test1", "Test2", "Test3" -Verbose

    .EXAMPLE
    Show-PasswordQuality -FilePath "C:\Support\GitHub\TheDashboard\Ignore\Reports\CustomReports\PasswordQuality_$(Get-Date -f yyyy-MM-dd_HHmmss).html" -WeakPasswords "Test1", "Test2", "Test3" #-Verbose

    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [alias('ForestName')][string] $Forest,
        [string[]] $ExcludeDomains,
        [alias('Domain', 'Domains')][string[]] $IncludeDomains,
        [System.Collections.IDictionary] $ExtendedForestInformation,
        [string] $FilePath,
        [switch] $DontShow,
        [switch] $Online,
        [string[]] $WeakPasswords,
        [switch] $SeparateDuplicateGroups,
        [switch] $PassThru,
        [switch] $AddWorldMap,
        [alias('LogFile')][string] $LogPath,
        [switch] $LogShowTime,
        [string] $LogTimeFormat = "yyyy-MM-dd HH:mm:ss"
    )
    $TimeStart = Start-TimeLog
    $Script:Reporting = [ordered] @{}
    $Script:Reporting['Version'] = Get-GitHubVersion -Cmdlet 'Show-PasswordQuality' -RepositoryOwner 'evotecit' -RepositoryName 'PasswordSolution'

    Write-Color -Text '[i]', "[PasswordSolution] ", 'Version', ' [Informative] ', $Script:Reporting['Version'] -Color Yellow, DarkGray, Yellow, DarkGray, Magenta

    Set-LoggingCapabilities -LogPath $LogPath -LogMaximum $LogMaximum -ShowTime:$LogShowTime.IsPresent -TimeFormat $TimeFormat
    # since the first entry didn't go to log file, this will
    Write-Color -InformationAction SilentlyContinue -Text '[i]', "[PasswordSolution] ", 'Version', ' [Informative] ', $Script:Reporting['Version'] -Color Yellow, DarkGray, Yellow, DarkGray, Magenta -NoConsoleOutput

    Write-Color '[i]', ' Gathering passwords data' -Color Yellow, DarkGray, Yellow, DarkGray, Magenta
    Write-Color '[i]', ' Using provided ', $WeakPasswords.Count, " weak passwords to verify against." -Color Yellow, DarkGray, Yellow, DarkGray, Magenta
    $TimeStartPasswords = Start-TimeLog
    $PasswordQuality = Find-PasswordQuality -IncludeStatistics -WeakPasswords $WeakPasswords -Forest $Forest -ExcludeDomains $ExcludeDomains -IncludeDomains $IncludeDomains -ExtendedForestInformation $ExtendedForestInformation
    if (-not $PasswordQuality) {
        # most likely DSInternals not installed
        return
    }
    $Users = $PasswordQuality.Users
    $Statistics = $PasswordQuality.Statistics
    $Countries = $PasswordQuality.StatisticsCountry
    $CountriesCodes = $PasswordQuality.StatisticsCountryCode
    $Continents = $PasswordQuality.StatisticsContinents

    $EndLogPasswords = Stop-TimeLog -Time $TimeStartPasswords -Option OneLiner

    Write-Color '[i]', ' Time to gather passwords data ', $EndLogPasswords -Color Yellow, DarkGray, Yellow, DarkGray, Magenta

    $TimeStartHTML = Start-TimeLog
    Write-Color -Text '[i] ', 'Generating HTML report...' -Color Yellow, DarkGray
    New-HTML {
        New-HTMLTabStyle -BorderRadius 0px -TextTransform capitalize -BackgroundColorActive SlateGrey
        New-HTMLSectionStyle -BorderRadius 0px -HeaderBackGroundColor Grey -RemoveShadow
        New-HTMLPanelStyle -BorderRadius 0px
        New-HTMLTableOption -DataStore JavaScript -BoolAsString -ArrayJoinString ', ' -ArrayJoin

        New-HTMLHeader {
            New-HTMLSection -Invisible {
                New-HTMLSection {
                    New-HTMLText -Text "Report generated on $(Get-Date)" -Color Blue
                } -JustifyContent flex-start -Invisible
                New-HTMLSection {
                    New-HTMLText -Text "Password Solution - $($Script:Reporting['Version'])" -Color Blue
                } -JustifyContent flex-end -Invisible
            }
        }

        Write-Color -Text '[i] ', 'Generating summary statistics' -Color Yellow, DarkGray

        New-HTMLSection {
            New-HTMLSection -Invisible {
                New-HTMLPanel -Invisible {
                    New-HTMLText -Text @(
                        "This report shows current status of an Active Directory forest $($PasswordQuality.Forest)."
                        "It focuses on the password quality of users in the following domains: "
                    ) -FontSize 12px
                    New-HTMLList {
                        foreach ($Domain in $PasswordQuality.Domains) {
                            New-HTMLListItem -Text $Domain -Color Blue
                        }
                    } -FontSize 12px
                    New-HTMLText -Text @(
                        "This report uses ", $WeakPasswords.Count, " weak passwords to check for, as provided during runtime."
                    ) -FontSize 12px -Color None, Red, None -FontWeight normal, bold, normal
                    #New-HTMLText -LineBreak
                    New-HTMLText -Text "Here's a short overview of what this report shows:" -Color None -FontSize 12px
                    #New-HTMLText -LineBreak
                    New-HTMLList {
                        foreach ($Statistic in $Statistics.Keys | Where-Object { $_ -notlike '*EnabledOnly' -and $_ -notlike '*DisabledOnly' } ) {
                            $ValueTotal = $Statistics[$Statistic]
                            if ($Statistic -eq "DuplicatePasswordGroups") {
                                $ValueEnabled = $Statistics['DuplicatePasswordUsersEnabledOnly']
                                $ValueDisabled = $Statistics['DuplicatePasswordUsersDisabledOnly']
                                New-HTMLListItem -Text @(
                                    "$($Statistic)",
                                    " property shows there are "
                                    "$ValueTotal"
                                    " groups of people with duplicate passwords."
                                ) -Color Blue, None, Salmon, None, LightSkyBlue, None -FontWeight bold, normal, bold, normal, bold, normal
                            } elseif ($Statistic -eq 'DuplicatePasswordUsers') {

                                $ValueEnabled = $Statistics['DuplicatePasswordUsersEnabledOnly']
                                $ValueDisabled = $Statistics['DuplicatePasswordUsersDisabledOnly']

                                New-HTMLListItem -Text @(
                                    "$($Statistic)",
                                    " property shows there are "
                                    "$ValueEnabled"
                                    " enabled "
                                    $ValueDisabled
                                    " accounts having duplicate passwords with other accounts."
                                ) -Color Blue, None, Salmon, None, LightSkyBlue, None -FontWeight bold, normal, bold, normal, bold, normal
                            } else {
                                $ValueEnabled = $Statistics[$Statistic + 'EnabledOnly']
                                $ValueDisabled = $Statistics[$Statistic + 'DisabledOnly']

                                New-HTMLListItem -Text @(
                                    "$($Statistic)",
                                    " property shows there are "
                                    "$ValueEnabled "
                                    "enabled accounts, and "
                                    "$ValueDisabled "
                                    "that are disabled."
                                ) -Color Blue, None, Salmon, None, LightSkyBlue, None -FontWeight bold, normal, bold, normal, bold, normal
                            }
                        }
                    } -Type Unordered -FontSize 12px

                    New-HTMLText -Text "Please review the report and make sure that you're happy with findings!" -Color Blue -FontSize 12px
                }
            }
            New-HTMLSection -Invisible {
                New-HTMLChart {
                    New-ChartBarOptions -Type barStacked
                    New-ChartAxisY -LabelMaxWidth 250 -Show -LabelAlign left
                    New-ChartLegend -LegendPosition bottom -HorizontalAlign center -Color Alizarin, LightSkyBlue -Names 'Enabled', 'Disabled'
                    foreach ($Statistic in $Statistics.Keys | Where-Object { $_ -notlike '*EnabledOnly' -and $_ -notlike '*DisabledOnly' } ) {
                        if ($Statistic -eq "DuplicatePasswordGroups") {
                            $ValueTotal = $Statistics[$Statistic]
                            New-ChartBar -Name $Statistic -Value @($ValueTotal, 0)
                        } else {
                            $ValueEnabled = $Statistics[$Statistic + 'EnabledOnly']
                            $ValueDisabled = $Statistics[$Statistic + 'DisabledOnly']
                            New-ChartBar -Name $Statistic -Value @($ValueEnabled, $ValueDisabled)
                        }
                    }
                    # # Define event
                    # New-ChartEvent -DataTableID 'NewIDtoSearchInChart' -ColumnID 0
                }
            }
        }

        $PropertiesHighlight = @(
            'ClearTextPassword'
            'LMHash'
            'EmptyPassword'
            'WeakPassword'
            #'DefaultComputerPassword'
            #'PasswordNotRequired'
            #'PasswordNeverExpires'
            'AESKeysMissing'
            'PreAuthNotRequired'
            'DESEncryptionOnly'
            'Kerberoastable'
            'DelegatableAdmins'
            'SmartCardUsersWithPassword'
            #'DuplicatePasswordGroups'
        )

        Write-Color -Text '[i] ', 'Generating users table with all information' -Color Yellow, DarkGray

        New-HTMLSection -HeaderText "Password Quality" {
            New-HTMLTable -DataTable $Users -Filtering {
                New-HTMLTableCondition -Name 'Enabled' -ComparisonType string -Operator eq -Value $true -BackgroundColor LimeGreen -FailBackgroundColor BlizzardBlue
                New-HTMLTableCondition -Name 'LastLogonDays' -ComparisonType number -Operator lt -Value 30 -BackgroundColor LimeGreen -HighlightHeaders LastLogonDays, LastLogonDate
                New-HTMLTableCondition -Name 'LastLogonDays' -ComparisonType number -Operator gt -Value 30 -BackgroundColor Orange -HighlightHeaders LastLogonDays, LastLogonDate
                New-HTMLTableCondition -Name 'LastLogonDays' -ComparisonType number -Operator gt -Value 60 -BackgroundColor Alizarin -HighlightHeaders LastLogonDays, LastLogonDate
                New-HTMLTableCondition -Name 'LastLogonDays' -ComparisonType string -Operator eq -Value '' -BackgroundColor None -HighlightHeaders LastLogonDays, LastLogonDate
                New-HTMLTableCondition -Name 'PasswordLastChangedDays' -ComparisonType number -Operator ge -Value 0 -BackgroundColor LimeGreen -HighlightHeaders PasswordLastSet, PasswordLastChangedDays
                New-HTMLTableCondition -Name 'PasswordLastChangedDays' -ComparisonType number -Operator gt -Value 300 -BackgroundColor Orange -HighlightHeaders PasswordLastSet, PasswordLastChangedDays
                New-HTMLTableCondition -Name 'PasswordLastChangedDays' -ComparisonType number -Operator gt -Value 360 -BackgroundColor Alizarin -HighlightHeaders PasswordLastSet, PasswordLastChangedDays
                New-HTMLTableCondition -Name 'PasswordNotRequired' -ComparisonType string -Operator eq -Value $false -BackgroundColor LimeGreen -FailBackgroundColor Alizarin
                New-HTMLTableCondition -Name 'PasswordExpired' -ComparisonType string -Operator eq -Value $false -BackgroundColor LimeGreen -FailBackgroundColor Alizarin -HighlightHeaders PasswordExpired, DaysToExpire, DateExpiry

                foreach ($Property in $PropertiesHighlight) {
                    New-HTMLTableCondition -Name $Property -ComparisonType string -Operator eq -Value $true -BackgroundColor Salmon -FailBackgroundColor LightGreen
                }
                New-HTMLTableCondition -Name 'DuplicatePasswordGroups' -ComparisonType string -Operator ne -Value "" -BackgroundColor Orange -FailBackgroundColor LightGreen
            } -ScrollX -ExcludeProperty 'RuleName', 'RuleOptions'

        }
        if ($SeparateDuplicateGroups) {
            Write-Color -Text '[i] ', 'Generating duplicate password groups section' -Color Yellow, DarkGray
            New-HTMLSection -HeaderText "Duplicate Password Groups" {
                $TotalDuplicateGroups = 0
                $EnabledUsersInDuplicateGroups = 0
                $DisabledUsersInDuplicateGroups = 0
                $DuplicateGroups = [ordered] @{}
                foreach ($User in $Users) {
                    if ($User.DuplicatePasswordGroups) {
                        if ($User.Enabled) {
                            $EnabledUsersInDuplicateGroups++
                        } else {
                            $DisabledUsersInDuplicateGroups++
                        }
                        if (-not $DuplicateGroups[$User.DuplicatePasswordGroups]) {
                            $DuplicateGroups[$User.DuplicatePasswordGroups] = [PSCustomObject] @{
                                GroupName             = $User.DuplicatePasswordGroups
                                UsersInGroup          = 0
                                Users                 = [System.Collections.Generic.List[string]]::new()
                                Country               = [System.Collections.Generic.List[string]]::new()
                                UsersBySamAccountName = [System.Collections.Generic.List[string]]::new()
                                UsersByUPN            = [System.Collections.Generic.List[string]]::new()
                                UsersByEmail          = [System.Collections.Generic.List[string]]::new()
                            }
                        }
                        $DuplicateGroups[$User.DuplicatePasswordGroups].Users.Add($User.DisplayName)
                        if ($User.EmailAddress) {
                            $DuplicateGroups[$User.DuplicatePasswordGroups].UsersByEmail.Add($User.EmailAddress)
                        }
                        if ($User.UserPrincipalName) {
                            $DuplicateGroups[$User.DuplicatePasswordGroups].UsersByUPN.Add($User.UserPrincipalName)
                        }
                        if ($User.SamAccountName) {
                            $DuplicateGroups[$User.DuplicatePasswordGroups].UsersBySamAccountName.Add($User.SamAccountName)
                        }
                        $DuplicateGroups[$User.DuplicatePasswordGroups].Country.Add($User.Country)
                    }
                }

                $TotalDuplicateGroups = $DuplicateGroups.Keys.Count

                foreach ($Group in $DuplicateGroups.Values) {
                    $Group.UsersInGroup = $Group.Users.Count
                    $Group.Country = $Group.Country | Select-Object -Unique
                }

                New-HTMLContainer {
                    New-HTMLSection {
                        New-HTMLPanel {
                            New-HTMLToast -TextHeader 'Total Duplicate Groups' -Text "Groups of users to review: $TotalDuplicateGroups" -BarColorLeft MayaBlue -IconSolid info-circle -IconColor MayaBlue
                        } -Invisible
                        New-HTMLPanel {
                            New-HTMLToast -TextHeader 'Enabled Users' -Text "Users with duplicate password that are enabled: $EnabledUsersInDuplicateGroups" -BarColorLeft OrangeRed -IconSolid info-circle -IconColor OrangeRed
                        } -Invisible
                        New-HTMLPanel {
                            New-HTMLToast -TextHeader 'Disabled Users' -Text "Users with duplicate password that are disabled: $DisabledUsersInDuplicateGroups" -BarColorLeft OrangePeel -IconSolid info-circle -IconColor OrangePeel
                        } -Invisible
                    } -Invisible

                    New-HTMLSection -Invisible {
                        New-HTMLTable -DataTable $DuplicateGroups.Values {
                        } -Filtering -Title "Duplicate Password Group: $DuplicateGroup" -ScrollX -ExcludeProperty 'RuleName', 'RuleOptions', 'Type', 'CountryCode'
                    }
                }
            }
        }
        if ($AddWorldMap) {
            Write-Color -Text '[i] ', 'Generating duplicate passwords map' -Color Yellow, DarkGray
            New-HTMLSection -HeaderText 'Duplicate Passwords Per Country' {
                New-HTMLTabPanel {
                    New-HTMLTab -Name 'Map showing duplicate passwords per country' {
                        New-HTMLSection -Invisible {
                            New-HTMLPanel {
                                New-HTMLMap -Map world_countries {
                                    # add the map areas
                                    # we will add unknown countries to the Greenland area
                                    foreach ($Country in $CountriesCodes['DuplicatePasswordUsers'].Keys) {
                                        if ($Country -eq 'Unknown') {
                                            New-MapArea -Area 'GL' -Value $CountriesCodes['DuplicatePasswordUsers'][$Country] -Tooltip {
                                                New-HTMLText -Text @(
                                                    'Unknown / Unavailable'
                                                    '<br>'
                                                    "Users with duplicate passwords $($CountriesCodes['DuplicatePasswordUsers'][$Country])"
                                                ) -Color Black, Black, Blue -FontWeight bold, normal, normal -SkipParagraph -FontSize 15px, 14px, 14px
                                            }
                                        } else {
                                            New-MapArea -Area $Country -Value $CountriesCodes['DuplicatePasswordUsers'][$Country] -Tooltip {
                                                New-HTMLText -Text @(
                                                    Convert-CountryCodeToCountry -CountryCode $Country
                                                    '<br>'
                                                    "Users with duplicate passwords $($CountriesCodes['DuplicatePasswordUsers'][$Country])"
                                                ) -Color Black, Black, Blue -FontWeight bold, normal, normal -SkipParagraph -FontSize 15px, 14px, 14px
                                            }
                                        }
                                    }
                                    # configure legend
                                    New-MapLegendOption -Type 'Area' -Mode horizontal
                                    New-MapLegendOption -Type 'Plot' -Mode horizontal
                                    # add legend
                                    New-MapLegendSlice -Type 'Area' -Label 'Duplicate passwords up to 5' -Min 0 -Max 5 -SliceColor 'Bisque' -StrokeWidth 0
                                    New-MapLegendSlice -Type 'Area' -Label 'Duplicate between 5 and 15' -Min 6 -Max 15 -SliceColor 'Amber' -StrokeWidth 0
                                    New-MapLegendSlice -Type 'Area' -Label 'Duplicate between 16 and 30' -Min 16 -Max 30 -SliceColor 'CarnationPink' -StrokeWidth 0
                                    New-MapLegendSlice -Type 'Area' -Label 'Duplicate between 31 and 50' -Min 31 -Max 50 -SliceColor 'OrangeRed' -StrokeWidth 0
                                    New-MapLegendSlice -Type 'Area' -Label 'Duplicate over 50' -Min 51 -Max 300 -SliceColor 'Scarlet' -StrokeWidth 0
                                } -ShowAreaLegend #-AreaTitle "Duplicate Passwords Users"
                            }
                        }
                    }
                    New-HTMLTab -Name 'Duplicate Passwords Per Country' {
                        New-HTMLTable -DataTable $Countries['DuplicatePasswordUsers'] -Filtering
                    }
                    New-HTMLTab -Name 'Duplicate Passwords Per Continent' {
                        New-HTMLTable -DataTable $Continents['DuplicatePasswordUsers'] -Filtering
                    }
                }
            }
            Write-Color -Text '[i] ', 'Generating weak password map' -Color Yellow, DarkGray
            New-HTMLSection -HeaderText 'Weak Password Per Country' {
                New-HTMLTabPanel {
                    New-HTMLTab -Name 'Map showing weak password per country' {
                        New-HTMLSection -Invisible {
                            New-HTMLPanel {
                                New-HTMLMap -Map world_countries {
                                    # add the map areas
                                    # we will add unknown countries to the Greenland area
                                    foreach ($Country in $CountriesCodes['WeakPassword'].Keys) {
                                        if ($Country -eq 'Unknown') {
                                            New-MapArea -Area 'GL' -Value $CountriesCodes['WeakPassword'][$Country] -Tooltip {
                                                New-HTMLText -Text @(
                                                    'Unknown / Unavailable'
                                                    '<br>'
                                                    "Users with weak passwords $($CountriesCodes['WeakPassword'][$Country])"
                                                ) -Color Black, Black, Blue -FontWeight bold, normal, normal -SkipParagraph -FontSize 15px, 14px, 14px
                                            }
                                        } else {
                                            New-MapArea -Area $Country -Value $CountriesCodes['WeakPassword'][$Country] -Tooltip {
                                                New-HTMLText -Text @(
                                                    Convert-CountryCodeToCountry -CountryCode $Country
                                                    '<br>'
                                                    "Users with weak passwords $($CountriesCodes['WeakPassword'][$Country])"
                                                ) -Color Black, Black, Blue -FontWeight bold, normal, normal -SkipParagraph -FontSize 15px, 14px, 14px
                                            }
                                        }
                                    }
                                    # configure legend
                                    New-MapLegendOption -Type 'Area' -Mode horizontal
                                    New-MapLegendOption -Type 'Plot' -Mode horizontal
                                    # add legend
                                    New-MapLegendSlice -Type 'Area' -Label 'Weak passwords up to 5' -Min 0 -Max 5 -SliceColor 'Bisque' -StrokeWidth 0
                                    New-MapLegendSlice -Type 'Area' -Label 'Weak between 5 and 15' -Min 6 -Max 15 -SliceColor 'Amber' -StrokeWidth 0
                                    New-MapLegendSlice -Type 'Area' -Label 'Weak between 16 and 30' -Min 16 -Max 30 -SliceColor 'CarnationPink' -StrokeWidth 0
                                    New-MapLegendSlice -Type 'Area' -Label 'Weak between 31 and 50' -Min 31 -Max 50 -SliceColor 'OrangeRed' -StrokeWidth 0
                                    New-MapLegendSlice -Type 'Area' -Label 'Weak over 50' -Min 51 -Max 300 -SliceColor 'Scarlet' -StrokeWidth 0
                                } -ShowAreaLegend #-AreaTitle "Weak Password Users"
                            }
                        }
                    }
                    New-HTMLTab -Name 'Weak Password Per Country' {
                        New-HTMLTable -DataTable $Countries['WeakPassword'] -Filtering
                    }
                    New-HTMLTab -Name 'Weak Password Per Continent' {
                        New-HTMLTable -DataTable $Continents['WeakPassword'] -Filtering
                    }
                }
            }
            if ($LogPath -and (Test-Path -LiteralPath $LogPath)) {
                $LogContent = Get-Content -Raw -LiteralPath $LogPath
                New-HTMLSection -Name 'Log' {
                    New-HTMLCodeBlock -Code $LogContent -Style generic
                }
            }
        }
    } -ShowHTML:(-not $DontShow.IsPresent) -Online:$Online.IsPresent -TitleText "Password Solution - Quality Password Check" -Author "Password Solution" -FilePath $FilePath

    $EndLogHTML = Stop-TimeLog -Time $TimeStartHTML -Option OneLiner
    $EndLog = Stop-TimeLog -Time $TimeStart -Option OneLiner
    Write-Color '[i]', ' Time to generate HTML ', $EndLogHTML -Color Yellow, DarkGray, Yellow, DarkGray, Magenta
    Write-Color '[i]', ' Time to generate ', $EndLog -Color Yellow, DarkGray, Yellow, DarkGray, Magenta
    Write-Color '[i]', "[PasswordSolution] ", 'Version', ' [Informative] ', $Script:Reporting['Version'] -Color Yellow, DarkGray, Yellow, DarkGray, Magenta

    if ($PassThru) {
        $PasswordQuality
    }
}