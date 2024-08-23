﻿@{
    AliasesToExport      = @()
    Author               = 'Przemyslaw Klys'
    CmdletsToExport      = @()
    CompanyName          = 'Evotec'
    CompatiblePSEditions = @('Desktop', 'Core')
    Copyright            = '(c) 2011 - 2024 Przemyslaw Klys @ Evotec. All rights reserved.'
    Description          = 'This module allows the creation of password expiry emails for users, managers, administrators, and security according to defined templates. It''s able to work with different rules allowing to fully customize who gets the email and when.'
    FunctionsToExport    = @('Find-Password', 'Find-PasswordNotification', 'Find-PasswordQuality', 'New-PasswordConfigurationEmail', 'New-PasswordConfigurationExternalUsers', 'New-PasswordConfigurationOption', 'New-PasswordConfigurationReport', 'New-PasswordConfigurationRule', 'New-PasswordConfigurationRuleReminder', 'New-PasswordConfigurationTemplate', 'New-PasswordConfigurationType', 'Show-PasswordQuality', 'Start-PasswordSolution')
    GUID                 = 'c58ff818-1de6-4500-961c-a243c2043255'
    ModuleVersion        = '1.3.2'
    PowerShellVersion    = '5.1'
    PrivateData          = @{
        PSData = @{
            ExternalModuleDependencies = @('ActiveDirectory')
            IconUri                    = 'https://evotec.xyz/wp-content/uploads/2022/08/PasswordSolution.png'
            ProjectUri                 = 'https://github.com/EvotecIT/PasswordSolution'
            Tags                       = @('password', 'passwordexpiry', 'activedirectory', 'windows')
        }
    }
    RequiredModules      = @(@{
            Guid          = 'ee272aa8-baaa-4edf-9f45-b6d6f7d844fe'
            ModuleName    = 'PSSharedGoods'
            ModuleVersion = '0.0.295'
        }, @{
            Guid          = 'a7bdf640-f5cb-4acf-9de0-365b322d245c'
            ModuleName    = 'PSWriteHTML'
            ModuleVersion = '1.27.0'
        }, @{
            Guid          = '0b0ba5c5-ec85-4c2b-a718-874e55a8bc3f'
            ModuleName    = 'PSWriteColor'
            ModuleVersion = '1.0.1'
        }, @{
            Guid          = '2b0ea9f1-3ff1-4300-b939-106d5da608fa'
            ModuleName    = 'Mailozaurr'
            ModuleVersion = '1.0.0'
        }, 'ActiveDirectory')
    RootModule           = 'PasswordSolution.psm1'
}