function Send-PasswordQuality {
    [CmdletBinding()]
    param (
        [scriptblock] $Configuration,
        [alias('ForestName')][string] $Forest,
        [string[]] $ExcludeDomains,
        [alias('Domain', 'Domains')][string[]] $IncludeDomains,
        [System.Collections.IDictionary] $ExtendedForestInformation,
        [string] $FilePath,
        [switch] $DontShow,
        [switch] $Online,
        [alias('KnownPasswords')][string[]] $WeakPasswords,
        [alias('KnownPasswordsFilePath')][string] $WeakPasswordsFilePath,
        [alias('KnownPasswordsHashesFile')][string] $WeakPasswordsHashesFile,
        [alias('KnownPasswordsHashesSortedFile')][string] $WeakPasswordsHashesSortedFile,
        [switch] $SeparateDuplicateGroups,
        [switch] $PassThru,
        [switch] $AddWorldMap,
        [alias('LogFile')][string] $LogPath,
        [int] $LogMaximum,
        [switch] $LogShowTime,
        [string] $LogTimeFormat = "yyyy-MM-dd HH:mm:ss",
        [System.Collections.IDictionary[]] $Replacements,
        [Array] $EmailRedirect
    )

    $TimeStart = Start-TimeLog
    $Script:Reporting = [ordered] @{}
    $Script:Reporting['Version'] = Get-GitHubVersion -Cmdlet 'Show-PasswordQuality' -RepositoryOwner 'evotecit' -RepositoryName 'PasswordSolution'

    Write-Color -Text '[i]', "[PasswordSolution] ", 'Version', ' [Informative] ', $Script:Reporting['Version'] -Color Yellow, DarkGray, Yellow, DarkGray, Magenta

    Set-LoggingCapabilities -LogPath $LogPath -LogMaximum $LogMaximum -ShowTime:$LogShowTime -TimeFormat $LogTimeFormat -ScriptPath $MyInvocation.ScriptName

    # since the first entry didn't go to log file, this will
    Write-Color -Text '[i]', "[PasswordSolution] ", 'Version', ' [Informative] ', $Script:Reporting['Version'] -Color Yellow, DarkGray, Yellow, DarkGray, Magenta -NoConsoleOutput

    $ConfigurationData = Set-PasswordQualityConfiguration -ConfigurationDSL $Configuration
    if (-not $ConfigurationData) {
        Write-Color '[e]', ' Configuration data is empty, fix errors and try again...' -Color Yellow, Red
        return
    }

    Write-Color '[i]', ' Gathering passwords data' -Color Yellow, DarkGray, Yellow, DarkGray, Magenta
    Write-Color '[i]', ' Using provided ', $WeakPasswords.Count, " weak passwords to verify against." -Color Yellow, DarkGray, Yellow, DarkGray, Magenta
    $TimeStartPasswords = Start-TimeLog
    $findPasswordQualitySplat = @{
        IncludeStatistics             = $true
        WeakPasswords                 = $WeakPasswords
        WeakPasswordsFilePath         = $WeakPasswordsFilePath
        WeakPasswordsHashesFile       = $WeakPasswordsHashesFile
        WeakPasswordsHashesSortedFile = $WeakPasswordsHashesSortedFile
        Forest                        = $Forest
        ExcludeDomains                = $ExcludeDomains
        IncludeDomains                = $IncludeDomains
        ExtendedForestInformation     = $ExtendedForestInformation
    }
    if ($Replacements) {
        $findPasswordQualitySplat['Replacements'] = $Replacements
    }
    $PasswordQuality = Find-PasswordQuality @findPasswordQualitySplat
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

    $OutputData = Send-InternalPasswordQualityEmails -Configuration $ConfigurationData -Users $Users -Statistics $Statistics -EmailRedirect $EmailRedirect

    $EndLog = Stop-TimeLog -Time $TimeStart -Option OneLiner
    Write-Color '[i]', ' Time to generate HTML ', $EndLogHTML -Color Yellow, DarkGray, Yellow, DarkGray, Magenta
    Write-Color '[i]', ' Time to generate ', $EndLog -Color Yellow, DarkGray, Yellow, DarkGray, Magenta
    Write-Color '[i]', "[PasswordSolution] ", 'Version', ' [Informative] ', $Script:Reporting['Version'] -Color Yellow, DarkGray, Yellow, DarkGray, Magenta

    if ($PassThru) {
        $OutputData.PasswordQuality = $PasswordQuality
        $OutputData
    }
}