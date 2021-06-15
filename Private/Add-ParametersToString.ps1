function Add-ParametersToString {
    <#
    .SYNOPSIS
    Short description

    .DESCRIPTION
    Long description

    .PARAMETER String
    Parameter description

    .PARAMETER Parameter
    Parameter description

    .EXAMPLE
    $Test = 'this is a string $Test - and $Test2 AND $tEST3'

    Add-ParametersToString -String $Test -Parameter @{
        Testooo = 'sdsds'
        Test    = 'oh my god'
        Test2   = 'ole ole'
        TEST3   = '56555'
    }

    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [string] $String,
        [System.Collections.IDictionary] $Parameter
    )
    $Sorted = $Parameter.Keys | Sort-Object { $_.length } -Descending


    foreach ($Key in $Sorted) {
        $String = $String -ireplace [Regex]::Escape("`$$Key"), $Parameter[$Key]
    }
    $String
}