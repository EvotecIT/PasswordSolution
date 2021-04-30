@{
    AliasesToExport      = @()
    Author               = 'Przemyslaw Klys'
    CmdletsToExport      = @()
    CompanyName          = 'Evotec'
    CompatiblePSEditions = @('Desktop', 'Core')
    Copyright            = '(c) 2011 - 2021 Przemyslaw Klys @ Evotec. All rights reserved.'
    Description          = 'This module allows creation of password expiry emails for users, managers and administrators according to defined template.'
    FunctionsToExport    = @('Find-Password', 'Start-PasswordSolution')
    GUID                 = 'c58ff818-1de6-4500-961c-a243c2043255'
    ModuleVersion        = '0.0.1'
    PowerShellVersion    = '5.1'
    PrivateData          = @{
        PSData = @{
            Tags       = @('password', 'passwordexpiry', 'activedirectory', 'windows')
            ProjectUri = 'https://github.com/EvotecIT/PasswordSolution'
        }
    }
    RequiredModules      = @(@{
            ModuleVersion = '0.0.200'
            ModuleName    = 'PSSharedGoods'
            Guid          = 'ee272aa8-baaa-4edf-9f45-b6d6f7d844fe'
        })
    RootModule           = 'PasswordSolution.psm1'
}