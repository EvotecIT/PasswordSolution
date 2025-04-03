function New-PasswordConfigurationReplacement {
    <#
    .SYNOPSIS
    Password configuration replacement function for replacing properties in the password configuration

    .DESCRIPTION
    Password configuration replacement function for replacing properties in the password configuration
    This function is used to create a replacement configuration for password properties.
    It takes a property name, a type of comparison, a hash table for property replacements, and an optional overwrite property name.
    It provides ability to replace specific value on given user object, on given property with different value.
    For example user having ExtensionAttribute1 with value 'PL' can be replaced with 'Poland'.

    .PARAMETER PropertyName
    Property name to be replaced in the password configuration.
    This is the name of the property in the password configuration that will be replaced.

    .PARAMETER Type
    Type of comparison to be used for the replacement.
    This parameter specifies the type of comparison to be used when replacing the property value.

    .PARAMETER PropertyReplacementHash
    Hash table containing the property replacements to be made.
    This parameter specifies the hash table that contains the mappings of property values to be replaced.

    .PARAMETER OverwritePropertyName
    Name of the property to be overwritten in the password configuration.
    This parameter specifies the name of the property that will be overwritten in the password configuration.
    This is an optional parameter.
    If the property doesn't exists, it will be added.
    If the property exists, it will be overwritten regardless of it's value.

    .EXAMPLE
    $showPasswordQualitySplat = @{
        FilePath                = "$PSScriptRoot\Reporting\PasswordQuality_$(Get-Date -f yyyy-MM-dd_HHmmss).html"
        WeakPasswords           = "Test1", "Test2", "Test3", 'February2023!#!@ok', $Passwords | ForEach-Object { $_ }
        SeparateDuplicateGroups = $true
        PassThru                = $true
        AddWorldMap             = $true
        LogPath                 = "$PSScriptRoot\Logs\PasswordQuality_$(Get-Date -f yyyy-MM-dd_HHmmss).log"
        Online                  = $true
        LogMaximum              = 5
        Replacements            = New-PasswordConfigurationReplacement -PropertyName 'Country' -Type eq -PropertyReplacementHash @{
            'PL'      = 'Poland'
            'DE'      = 'Germany'
            'AT'      = 'Austria'
            'IT'      = 'Italy'
            'Unknown' = 'Not specified in AD'
        } -OverwritePropertyName 'AddMe'
    }

    Show-PasswordQuality @showPasswordQualitySplat -Verbose

    .EXAMPLE
    $Replacements = @(
        New-PasswordConfigurationReplacement -PropertyName 'Country' -Type eq -PropertyReplacementHash @{
            'PL'      = 'Poland'
            'DE'      = 'Germany'
            'AT'      = 'Austria'
            'IT'      = 'Italy'
            'Unknown' = 'Not specified in AD'
        } -OverwritePropertyName 'CountryCode'
    )

    $Users = Find-PasswordQuality -Replacements $Replacements
    $Users | Format-Table

    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $PropertyName,
        [ValidateSet('eq')][string] $Type = 'eq',
        [Parameter(Mandatory)][System.Collections.IDictionary] $PropertyReplacementHash,
        [string] $OverwritePropertyName
    )
    if ($PropertyReplacementHash.Count -eq 0) {
        Write-Color -Text '[-] ', "Couldn't create replacement configuration as the hash is empty. Please fix 'New-PasswordConfigurationReplacement'" -Color Yellow, White
        return
    }

    $Output = [ordered] @{
        Type     = "PasswordConfigurationReplacement"
        Settings = @{
            PropertyName            = $PropertyName
            Type                    = $Type
            PropertyReplacementHash = $PropertyReplacementHash
            OverwritePropertyName   = $OverwritePropertyName
        }
    }
    $Output
}