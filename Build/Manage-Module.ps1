Clear-Host
Import-Module "PSPublishModule" -Force

Invoke-ModuleBuild -ModuleName 'PasswordSolution' {
    # Usual defaults as per standard module
    $Manifest = [ordered] @{
        # Version number of this module.
        ModuleVersion        = '1.0.X'
        # Supported PSEditions
        CompatiblePSEditions = @('Desktop', 'Core')

        PowerShellVersion    = '5.1'
        # ID used to uniquely identify this module
        GUID                 = 'c58ff818-1de6-4500-961c-a243c2043255'
        # Author of this module
        Author               = 'Przemyslaw Klys'
        # Company or vendor of this module
        CompanyName          = 'Evotec'
        # Copyright statement for this module
        Copyright            = "(c) 2011 - $((Get-Date).Year) Przemyslaw Klys @ Evotec. All rights reserved."
        # Description of the functionality provided by this module
        Description          = "This module allows the creation of password expiry emails for users, managers, administrators, and security according to defined templates. It's able to work with different rules allowing to fully customize who gets the email and when."
        # Minimum version of the Windows PowerShell engine required by this module
        Tags                 = 'password', 'passwordexpiry', 'activedirectory', 'windows'
        # A URL to the main website for this project.
        ProjectUri           = 'https://github.com/EvotecIT/PasswordSolution'

        # A URL to an icon representing this module.
        IconUri              = 'https://evotec.xyz/wp-content/uploads/2022/08/PasswordSolution.png'
    }
    New-ConfigurationManifest @Manifest

    New-ConfigurationModule -Type RequiredModule -Name 'PSSharedGoods', 'PSWriteHTML', 'PSWriteColor' -Guid Auto -Version Latest
    New-ConfigurationModule -Type RequiredModule -Name 'Mailozaurr' -Guid Auto -Version 1.0.0
    New-ConfigurationModule -Type ExternalModule -Name @(
        #"Microsoft.PowerShell.Management"
        #"Microsoft.PowerShell.Utility"
        "ActiveDirectory"
    )
    New-ConfigurationModule -Type ApprovedModule -Name 'PSSharedGoods', 'PSWriteColor', 'Connectimo', 'PSUnifi', 'PSWebToolbox', 'PSMyPassword'

    New-ConfigurationModuleSkip -IgnoreModuleName @(
        'DSInternals'
        'NetTCPIP'
        'Microsoft.WSMan.Management'
    ) #-IgnoreFunctionName 'Select-Unique', 'Compare-TwoArrays' # those functions are internal within private function

    $ConfigurationFormat = [ordered] @{
        RemoveComments                              = $true
        RemoveEmptyLines                            = $true

        PlaceOpenBraceEnable                        = $true
        PlaceOpenBraceOnSameLine                    = $true
        PlaceOpenBraceNewLineAfter                  = $true
        PlaceOpenBraceIgnoreOneLineBlock            = $false

        PlaceCloseBraceEnable                       = $true
        PlaceCloseBraceNewLineAfter                 = $false
        PlaceCloseBraceIgnoreOneLineBlock           = $true
        PlaceCloseBraceNoEmptyLineBefore            = $false

        UseConsistentIndentationEnable              = $true
        UseConsistentIndentationKind                = 'space'
        UseConsistentIndentationPipelineIndentation = 'IncreaseIndentationAfterEveryPipeline'
        UseConsistentIndentationIndentationSize     = 4

        UseConsistentWhitespaceEnable               = $true
        UseConsistentWhitespaceCheckInnerBrace      = $true
        UseConsistentWhitespaceCheckOpenBrace       = $true
        UseConsistentWhitespaceCheckOpenParen       = $true
        UseConsistentWhitespaceCheckOperator        = $true
        UseConsistentWhitespaceCheckPipe            = $true
        UseConsistentWhitespaceCheckSeparator       = $true

        AlignAssignmentStatementEnable              = $true
        AlignAssignmentStatementCheckHashtable      = $true

        UseCorrectCasingEnable                      = $true
    }
    # format PSD1 and PSM1 files when merging into a single file
    # enable formatting is not required as Configuration is provided
    New-ConfigurationFormat -ApplyTo 'OnMergePSM1', 'OnMergePSD1' -Sort None @ConfigurationFormat
    # format PSD1 and PSM1 files within the module
    # enable formatting is required to make sure that formatting is applied (with default settings)
    New-ConfigurationFormat -ApplyTo 'DefaultPSD1', 'DefaultPSM1' -EnableFormatting -Sort None
    # when creating PSD1 use special style without comments and with only required parameters
    New-ConfigurationFormat -ApplyTo 'DefaultPSD1', 'OnMergePSD1' -PSD1Style 'Minimal'
    # configuration for documentation, at the same time it enables documentation processing
    New-ConfigurationDocumentation -Enable:$false -StartClean -UpdateWhenNew -PathReadme 'Docs\Readme.md' -Path 'Docs'

    New-ConfigurationImportModule -ImportSelf

    New-ConfigurationBuild -Enable:$true -SignModule -MergeModuleOnBuild -MergeFunctionsFromApprovedModules -CertificateThumbprint '483292C9E317AA13B07BB7A96AE9D1A5ED9E7703'

    #New-ConfigurationTest -TestsPath "$PSScriptRoot\..\Tests" -Enable

    New-ConfigurationArtefact -Type Unpacked -Enable -Path "$PSScriptRoot\..\Artefacts\Unpacked" -AddRequiredModules
    New-ConfigurationArtefact -Type Packed -Enable -Path "$PSScriptRoot\..\Artefacts\Packed" -ArtefactName '<ModuleName>.v<ModuleVersion>.zip'

    # options for publishing to github/psgallery
    #New-ConfigurationPublish -Type PowerShellGallery -FilePath 'C:\Support\Important\PowerShellGalleryAPI.txt' -Enabled:$true
    #New-ConfigurationPublish -Type GitHub -FilePath 'C:\Support\Important\GitHubAPI.txt' -UserName 'EvotecIT' -Enabled:$true
} -ExitCode