function Show-PasswordQuality {
    [CmdletBinding()]
    param(
        [string] $FilePath,
        [switch] $DontShow,
        [switch] $Online,
        [string[]] $WeakPasswords
    )
    $Script:Reporting = [ordered] @{}
    $Script:Reporting['Version'] = Get-GitHubVersion -Cmdlet 'Show-PasswordQuality' -RepositoryOwner 'evotecit' -RepositoryName 'PasswordSolution'


    Write-Color '[i]', "[PasswordSolution] ", 'Version', ' [Informative] ', $Script:Reporting['Version'] -Color Yellow, DarkGray, Yellow, DarkGray, Magenta

    $PasswordQuality = Find-PasswordQuality -IncludeStatistics -WeakPasswords $WeakPasswords
    if (-not $PasswordQuality) {
        # most likely DSInternals not installed
        return
    }
    $Users = $PasswordQuality.Users
    $Statistics = $PasswordQuality.Statistics

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

        New-HTMLTable -DataTable $Users -Filtering {
            New-HTMLTableCondition -Name 'Enabled' -ComparisonType string -Operator eq -Value $true -BackgroundColor LimeGreen -FailBackgroundColor BlizzardBlue
            #New-HTMLTableCondition -Name 'LastLogonDays' -ComparisonType number -Operator gt -Value 60 -BackgroundColor Alizarin -HighlightHeaders LastLogonDays, LastLogonDate -FailBackgroundColor LimeGreen
            New-HTMLTableCondition -Name 'PasswordLastChangedDays' -ComparisonType number -Operator ge -Value 0 -BackgroundColor LimeGreen -HighlightHeaders PasswordLastSet, PasswordLastChangedDays
            New-HTMLTableCondition -Name 'PasswordLastChangedDays' -ComparisonType number -Operator gt -Value 300 -BackgroundColor Orange -HighlightHeaders PasswordLastSet, PasswordLastChangedDays
            New-HTMLTableCondition -Name 'PasswordLastChangedDays' -ComparisonType number -Operator gt -Value 360 -BackgroundColor Alizarin -HighlightHeaders PasswordLastSet, PasswordLastChangedDays
            New-HTMLTableCondition -Name 'PasswordNotRequired' -ComparisonType string -Operator eq -Value $false -BackgroundColor LimeGreen -FailBackgroundColor Alizarin
            New-HTMLTableCondition -Name 'PasswordExpired' -ComparisonType string -Operator eq -Value $false -BackgroundColor LimeGreen -FailBackgroundColor Alizarin -HighlightHeaders PasswordExpired, DaysToExpire, DateExpiry

            foreach ($Property in $PropertiesHighlight) {
                New-HTMLTableCondition -Name $Property -ComparisonType string -Operator eq -Value $true -BackgroundColor Salmon -FailBackgroundColor LightGreen
            }
            New-HTMLTableCondition -Name 'DuplicatePasswordGroups' -ComparisonType string -Operator ne -Value "" -BackgroundColor Orange -FailBackgroundColor LightGreen
        } -ScrollX -ExcludeProperty 'RuleName','RuleOptions'
    } -ShowHTML:(-not $DontShow.IsPresent) -Online:$Online.IsPresent -TitleText "Password Solution - Quality Password Check" -Author "Password Solution" -FilePath $FilePath
}