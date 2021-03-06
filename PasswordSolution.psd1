@{
    AliasesToExport      = @()
    Author               = 'Przemyslaw Klys'
    CmdletsToExport      = @()
    CompanyName          = 'Evotec'
    CompatiblePSEditions = @('Desktop', 'Core')
    Copyright            = '(c) 2011 - 2021 Przemyslaw Klys @ Evotec. All rights reserved.'
    Description          = 'This module allows the creation of password expiry emails for users, managers, administrators, and security according to defined templates. It''s able to work with different rules allowing to fully customize who gets the email and when.'
    FunctionsToExport    = @('Find-Password', 'Find-PasswordNotification', 'Start-PasswordSolution')
    GUID                 = 'c58ff818-1de6-4500-961c-a243c2043255'
    ModuleVersion        = '0.0.23'
    PowerShellVersion    = '5.1'
    PrivateData          = @{
        PSData = @{
            Tags                       = @('password', 'passwordexpiry', 'activedirectory', 'windows')
            ProjectUri                 = 'https://github.com/EvotecIT/PasswordSolution'
            ExternalModuleDependencies = @('ActiveDirectory')
        }
    }
    RequiredModules      = @(@{
            ModuleVersion = '0.0.214'
            ModuleName    = 'PSSharedGoods'
            Guid          = 'ee272aa8-baaa-4edf-9f45-b6d6f7d844fe'
        }, @{
            ModuleVersion = '0.0.18'
            ModuleName    = 'Mailozaurr'
            Guid          = '2b0ea9f1-3ff1-4300-b939-106d5da608fa'
        }, @{
            ModuleVersion = '0.0.158'
            ModuleName    = 'PSWriteHTML'
            Guid          = 'a7bdf640-f5cb-4acf-9de0-365b322d245c'
        }, 'ActiveDirectory')
    RootModule           = 'PasswordSolution.psm1'
}