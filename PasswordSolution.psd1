@{
    AliasesToExport      = @()
    Author               = 'Przemyslaw Klys'
    CmdletsToExport      = @()
    CompanyName          = 'Evotec'
    CompatiblePSEditions = @('Desktop', 'Core')
    Copyright            = '(c) 2011 - 2023 Przemyslaw Klys @ Evotec. All rights reserved.'
    Description          = 'This module allows the creation of password expiry emails for users, managers, administrators, and security according to defined templates. It''s able to work with different rules allowing to fully customize who gets the email and when.'
    FunctionsToExport    = @('Find-Password', 'Find-PasswordNotification', 'Find-PasswordQuality', 'New-PasswordConfigurationEmail', 'New-PasswordConfigurationOption', 'New-PasswordConfigurationReport', 'New-PasswordConfigurationRule', 'New-PasswordConfigurationRuleReminder', 'New-PasswordConfigurationTemplate', 'New-PasswordConfigurationType', 'Show-PasswordQuality', 'Start-PasswordSolution')
    GUID                 = 'c58ff818-1de6-4500-961c-a243c2043255'
    ModuleVersion        = '1.0.2'
    PowerShellVersion    = '5.1'
    PrivateData          = @{
        PSData = @{
            ProjectUri                 = 'https://github.com/EvotecIT/PasswordSolution'
            IconUri                    = 'https://evotec.xyz/wp-content/uploads/2022/08/PasswordSolution.png'
            Tags                       = @('password', 'passwordexpiry', 'activedirectory', 'windows')
            ExternalModuleDependencies = @('ActiveDirectory')
        }
    }
    RequiredModules      = @(@{
            ModuleName    = 'PSSharedGoods'
            ModuleVersion = '0.0.264'
            Guid          = 'ee272aa8-baaa-4edf-9f45-b6d6f7d844fe'
        }, @{
            ModuleName    = 'Mailozaurr'
            ModuleVersion = '1.0.0'
            Guid          = '2b0ea9f1-3ff1-4300-b939-106d5da608fa'
        }, @{
            ModuleName    = 'PSWriteHTML'
            ModuleVersion = '0.0.189'
            Guid          = 'a7bdf640-f5cb-4acf-9de0-365b322d245c'
        }, @{
            ModuleName    = 'PSWriteColor'
            ModuleVersion = '1.0.1'
            Guid          = '0b0ba5c5-ec85-4c2b-a718-874e55a8bc3f'
        }, 'ActiveDirectory')
    RootModule           = 'PasswordSolution.psm1'
}