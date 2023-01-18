<p align="center">
  <a href="https://dev.azure.com/evotecpl/PasswordSolution/_build/results?buildId=latest"><img src="https://dev.azure.com/evotecpl/PasswordSolution/_apis/build/status/EvotecIT.PasswordSolution"></a>
  <a href="https://www.powershellgallery.com/packages/PasswordSolution"><img src="https://img.shields.io/powershellgallery/v/PasswordSolution.svg"></a>
  <a href="https://www.powershellgallery.com/packages/PasswordSolution"><img src="https://img.shields.io/powershellgallery/vpre/PasswordSolution.svg?label=powershell%20gallery%20preview&colorB=yellow"></a>
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

PasswordSolution is a PowerShell module that provides Password Expiry notifications to users, managers, security and administrators. It's very configurable and was designed for enterprise use.

## Features

- Find all users with passwords that are expiring and send notifications
- Find and asses password quality of users in a forest

## Installing

Everyone can install this module from **PowerShellGallery** hosted by Microsoft. It's recommended way to work with the module.
Version on **PowershellGallery** is optimized for speed and signed. Using code from **GitHub** is **recommended for development**.

```powershell
Install-Module -Name PasswordSolution -AllowClobber -Force -Verbose
```

If you want to use Password Quality checks you need to install DSInternals manually.

```powershell
Install-Module DSInternals -Force -Verbose
```

Force and AllowClobber aren't necessary, but they do skip errors in case some appear.

## Updating

```powershell
Update-Module -Name PasswordSolution
```

That's it. Whenever there's a new version, you run the command, and you can enjoy it. Remember that you may need to close, reopen PowerShell session if you have already used module before updating it.

**The essential thing** is if something works for you on production, keep using it till you test the new version on a test computer. I do changes that may not be big, but big enough that auto-update may break your code. For example, small rename to a parameter and your code stops working! Be responsible!
