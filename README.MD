<p align="center">
  <a href="https://www.powershellgallery.com/packages/PasswordSolution"><img src="https://img.shields.io/powershellgallery/v/PasswordSolution.svg"></a>
  <a href="https://www.powershellgallery.com/packages/PasswordSolution"><img src="https://img.shields.io/powershellgallery/v/PasswordSolution.svg?label=powershell%20gallery%20preview&colorB=yellow&include_prereleases"></a>
  <a href="https://github.com/EvotecIT/PasswordSolution"><img src="https://img.shields.io/github/license/EvotecIT/PasswordSolution.svg"></a>
</p>

<p align="center">
  <a href="https://www.powershellgallery.com/packages/PasswordSolution"><img src="https://img.shields.io/powershellgallery/p/PasswordSolution.svg"></a>
  <a href="https://github.com/EvotecIT/PasswordSolution"><img src="https://img.shields.io/github/languages/top/evotecit/PasswordSolution.svg"></a>
  <a href="https://github.com/EvotecIT/PasswordSolution"><img src="https://img.shields.io/github/languages/code-size/evotecit/PasswordSolution.svg"></a>
  <a href="https://www.powershellgallery.com/packages/PasswordSolution"><img src="https://img.shields.io/powershellgallery/dt/PasswordSolution.svg"></a>
</p>

<p align="center">
  <a href="https://twitter.com/PrzemyslawKlys"><img src="https://img.shields.io/twitter/follow/PrzemyslawKlys.svg?label=Twitter%20%40PrzemyslawKlys&style=social"></a>
  <a href="https://evotec.xyz/hub"><img src="https://img.shields.io/badge/Blog-evotec.xyz-2A6496.svg"></a>
  <a href="https://www.linkedin.com/in/pklys"><img src="https://img.shields.io/badge/LinkedIn-pklys-0077B5.svg?logo=LinkedIn"></a>
</p>


# PasswordSolution

**PasswordSolution** is a PowerShell module that provides Password Expiry notifications to users, managers, security and administrators. It's very configurable and was designed for enterprise use.

## Features

- Find all users with passwords that are expiring and send notifications including service accounts, admin accounts and accounts that never expire
- Find and asses password quality of users in an Active Directory forest.

## Support This Project

If you find this project helpful, please consider supporting its development.
Your sponsorship will help the maintainers dedicate more time to maintenance and new feature development for everyone.

It takes a lot of time and effort to create and maintain this project.
By becoming a sponsor, you can help ensure that it stays free and accessible to everyone who needs it.

To become a sponsor, you can choose from the following options:

 - [Become a sponsor via GitHub Sponsors :heart:](https://github.com/sponsors/PrzemyslawKlys)
 - [Become a sponsor via PayPal :heart:](https://paypal.me/PrzemyslawKlys)

Your sponsorship is completely optional and not required for using this project.
We want this project to remain open-source and available for anyone to use for free,
regardless of whether they choose to sponsor it or not.

If you work for a company that uses our .NET libraries or PowerShell Modules,
please consider asking your manager or marketing team if your company would be interested in supporting this project.
Your company's support can help us continue to maintain and improve this project for the benefit of everyone.

Thank you for considering supporting this project!

## Installing

Everyone can install this module from **PowerShellGallery** hosted by Microsoft. It's recommended way to work with the module.
Version on **PowershellGallery** is optimized for speed and signed. Using code from **GitHub** is **recommended for development**.

```powershell
Install-Module -Name PasswordSolution -AllowClobber -Force -Verbose
```

If you want to use Password Quality checks you need to install DSInternals manually.
Due to sensitive nature of the module it's not included in the package.

```powershell
Install-Module DSInternals -Force -Verbose
```

Force and AllowClobber aren't necessary, but they do skip errors in case some appear, and they do update module if newer version is available.

## Updating

```powershell
Update-Module -Name PasswordSolution
```

That's it. Whenever there's a new version, you run the command, and you can enjoy it. Remember that you may need to close, reopen PowerShell session if you have already used module before updating it.

**The essential thing** is if something works for you on production, keep using it till you test the new version on a test computer. I do changes that may not be big, but big enough that auto-update may break your code. For example, small rename to a parameter and your code stops working! Be responsible!

## Usage - Password Quality

Password Quality report uses [DSInternals](https://github.com/MichaelGrafnetter/DSInternals) module to scan Active Directory for passwords and then assesses them.
It's a very powerful tool that can be used to find weak passwords in your environment.
It's also a great tool to find out if your password policy is working as expected.
PasswordSolution wraps around DSInternals to make it easier to use and to provide a nice HTML report.

```powershell
Show-PasswordQuality -FilePath C:\Temp\PasswordQuality.html -Online -WeakPasswords "Test1", "Test2", "Test3" -Verbose -SeparateDuplicateGroups -AddWorldMap -PassThru
```

or

```powershell
$showPasswordQualitySplat = @{
    FilePath                = "$PSScriptRoot\Reporting\PasswordQuality_$(Get-Date -f yyyy-MM-dd_HHmmss).html"
    WeakPasswords           = "Test1", "Test2", "Test3", 'February2023!#!@ok', $Passwords | ForEach-Object { $_ }
    SeparateDuplicateGroups = $true
    PassThru                = $true
    AddWorldMap             = $true
    LogPath                 = "$PSScriptRoot\Logs\PasswordQuality_$(Get-Date -f yyyy-MM-dd_HHmmss).log"
    Online                  = $true
    LogMaximum              = 5
}

Show-PasswordQuality @showPasswordQualitySplat -Verbose
```

And here's what you get

![PasswordQuality](https://raw.githubusercontent.com/EvotecIT/PasswordSolution/master/Docs/Images/PasswordQuality1.png)

![PasswordQuality](https://raw.githubusercontent.com/EvotecIT/PasswordSolution/master/Docs/Images/PasswordQuality2.png)

![PasswordQuality](https://raw.githubusercontent.com/EvotecIT/PasswordSolution/master/Docs/Images/PasswordQuality3.png)

![PasswordQuality](https://raw.githubusercontent.com/EvotecIT/PasswordSolution/master/Docs/Images/PasswordQuality4.png)

![PasswordQuality](https://raw.githubusercontent.com/EvotecIT/PasswordSolution/master/Docs/Images/PasswordQuality5.png)

![PasswordQuality](https://raw.githubusercontent.com/EvotecIT/PasswordSolution/master/Docs/Images/PasswordQuality6.png)