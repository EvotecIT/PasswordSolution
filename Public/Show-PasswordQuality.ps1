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
    Keep in mindd that this report can be created without internet access, just that using it with -Online switch will require internet access.

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
        [string] $FilePath,
        [switch] $DontShow,
        [switch] $Online,
        [string[]] $WeakPasswords,
        [switch] $SeparateDuplicateGroups
    )
    $TimeStart = Start-TimeLog
    $Script:Reporting = [ordered] @{}
    $Script:Reporting['Version'] = Get-GitHubVersion -Cmdlet 'Show-PasswordQuality' -RepositoryOwner 'evotecit' -RepositoryName 'PasswordSolution'


    Write-Color '[i]', "[PasswordSolution] ", 'Version', ' [Informative] ', $Script:Reporting['Version'] -Color Yellow, DarkGray, Yellow, DarkGray, Magenta
    Write-Color '[i]', ' Gathering passwords data' -Color Yellow, DarkGray, Yellow, DarkGray, Magenta
    $TimeStartPasswords = Start-TimeLog
    $PasswordQuality = Find-PasswordQuality -IncludeStatistics -WeakPasswords $WeakPasswords
    if (-not $PasswordQuality) {
        # most likely DSInternals not installed
        return
    }
    $Users = $PasswordQuality.Users
    $Statistics = $PasswordQuality.Statistics

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

        New-HTMLSection {
            New-HTMLSection -Invisible {
                New-HTMLPanel -Invisible {
                    New-HTMLText -Text @(
                        "This report shows current status of a forest."
                        "It focuses on the password quality that users are using across domain. "

                    ) -FontSize 12px
                    New-HTMLText -LineBreak
                    New-HTMLText -Text "Here's a short overview of what this report shows:" -Color Blue -FontSize 12px
                    New-HTMLText -LineBreak
                    New-HTMLList {
                        foreach ($Statistic in $Statistics.Keys | Where-Object { -not $_.EndsWith('Only') }) {
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
                    foreach ($Statistic in $Statistics.Keys | Where-Object { -not $_.EndsWith('Only') }) {
                        $ValueEnabled = $Statistics[$Statistic + 'EnabledOnly']
                        $ValueDisabled = $Statistics[$Statistic + 'DisabledOnly']
                        #$ValueTotal = $Statistics[$Statistic]
                        New-ChartBar -Name $Statistic -Value @($ValueEnabled, $ValueDisabled)
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
            'DefaultComputerPassword'
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
                        $DuplicateGroups[$User.DuplicatePasswordGroups] = [System.Collections.Generic.List[pscustomobject]]::new()
                    }
                    $DuplicateGroups[$User.DuplicatePasswordGroups].Add($User)
                }
            }
            $TotalDuplicateGroups = $DuplicateGroups.Keys.Count

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

                    New-HTMLTabPanel {
                        foreach ($DuplicateGroup in $DuplicateGroups.Keys | Sort-Object) {
                            New-HTMLTab -Name "$DuplicateGroup ($($DuplicateGroups[$DuplicateGroup].Count))" {
                                New-HTMLTable -DataTable $DuplicateGroups[$DuplicateGroup] {
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
                                } -Filtering -Title "Duplicate Password Group: $DuplicateGroup" -ScrollX -ExcludeProperty 'RuleName', 'RuleOptions'
                            }
                        }
                    }
                }
            }
        }
    } -ShowHTML:(-not $DontShow.IsPresent) -Online:$Online.IsPresent -TitleText "Password Solution - Quality Password Check" -Author "Password Solution" -FilePath $FilePath

    $EndLogHTML = Stop-TimeLog -Time $TimeStartHTML -Option OneLiner
    $EndLog = Stop-TimeLog -Time $TimeStart -Option OneLiner
    Write-Color '[i]', ' Time to generate HTML ', $EndLogHTML -Color Yellow, DarkGray, Yellow, DarkGray, Magenta
    Write-Color '[i]', ' Time to generate ', $EndLog -Color Yellow, DarkGray, Yellow, DarkGray, Magenta
    Write-Color '[i]', "[PasswordSolution] ", 'Version', ' [Informative] ', $Script:Reporting['Version'] -Color Yellow, DarkGray, Yellow, DarkGray, Magenta
}