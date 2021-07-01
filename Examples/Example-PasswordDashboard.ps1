Import-Module .\PasswordSolution.psd1 -Force

$Passwords = Find-Password

New-HTML {
    New-TableOption -DataStore JavaScript -ArrayJoin -BoolAsString
    New-HTMLTable -DataTable $Passwords -SearchBuilder -Filtering {

    }
} -ShowHTML