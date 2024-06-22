Import-Module .\PasswordSolution.psd1 -Force

Connect-MgGraph -NoWelcome
$Users = Find-PasswordEntra -Verbose
$Users | Out-HtmlView -ScrollX -Filtering -DataStore JavaScript