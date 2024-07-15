# PasswordSolution Release History

## 1.2.4 - 2024.07.15
- Fixes email address being wrong when using external system and using overwrite property

## 1.2.3 - 2024.06.23
- Small report improvements

## 1.2.2 - 2024.06.23
- Add reporting for replacements of emails from external sources

## 1.2.1 - 2024.06.23
- Allow using DSL and normal configuration at the same time

## 1.2.0 - 2024.06.23
- This version adds ability to allow overwritting email address from external system, or even CSV records based on prepared data
- Added `New-PasswordConfigurationExternalUsers` to allow for overwritting emails with external data in form of array of objects

## 1.1.1 - 2024.01.16
- Small improvement to error message being provided when sending email fails with summary of emails

## 1.1.0 - 2023.11.12
- Add support for weak password hashes (NTLM) in Password Quality Check
  - Add new parameters for `Find-PasswordQuality`
  - Add new parameters for `Show-PasswordQuality`

## 1.0.5 - 2023.10.18
- I can't type apparently so I fixed typo in code

## 1.0.3 - 2023.06.12
- Resolves issue with scanning **Active Directory** without exchange attributes

## 1.0.2 - 2023.06.06
- Improves `New-PasswordConfigurationRuleReminder` by allowing: `New-PasswordConfigurationRuleReminder -Type 'Manager' -ExpirationDays @(-200..-1), 0, 1, 2, 3, 7, 15, 30, 60 -ComparisonType eq` configuration

## 1.0.1 - 2023.06.01
- Fixes AdminSection not working properly

## 1.0.0 - 2023.05.25
- Improves exclusions https://github.com/EvotecIT/PasswordSolution/issues/7
- Simplifies configuration https://github.com/EvotecIT/PasswordSolution/issues/8
- Improves configuration https://github.com/EvotecIT/PasswordSolution/issues/3
- Improves configuration https://github.com/EvotecIT/PasswordSolution/issues/6
- Improve HTML report to be better in naming things
- Allow for overwriting manager field with different properties based on SamAccountName/DN
- New configuration option for Report (NestedRules), making separate tab for all rules in HTML
- Small documentation updates
- Adds defaults for templates meaning it's possible to skip their definitions although not very useful if you want personalized thing for your company users

This actually means:
- We now support DSL language for configuration (see examples)
- We now support much easier way of building configuration
- We now support overwrite email property per rule (global still works)
- We now support overwrite manager property per rule (global still works)

## 0.0.37 - 2023.04.25
- Improve `Maps` in Password Quality
- Improve colors, add column to a report

## 0.0.36 - 2023.04.24
- Added `Maps` to Password Quality
- Improved reporting in Password Quality with countries for Weak Passwords and Duplicate Groups
- Added Logs to Password Quality
- Minor improvements to reporting
- Make report much smaller

## 0.0.35 - 2023.04.19
- Small fix to report

## 0.0.34 - 2023.04.19
- Fixes reporting of quality passwords in overview, charts
- Removed tables for Duplicate Groups in favor of single table (1000 tabs don't seem to work in HTML 🤯)

## 0.0.33 - 2023.04.18
- General improvements
- Added separate duplicate groups tables to report for easier visibility
- Prefer writable DCs when quering AD

## 0.0.32 - 2023.01.18
- Removed DSInternals from required modules (made it optional / disables Password Quality Check). This is to not trigger any security alerts in your environment if you just want to use the module for password expiration and not password quality checks.

## 0.0.31 - 2023.01.17
- Small updates to logging

## 0.0.30 - 2023.01.17
- Add ability to check password quality (requires higher permissions) and based on DSInternals

## 0.0.29 - 2022.10.11
- Fixes to report **ShowSkippedUsers** to skip Contacts

## 0.0.28
- Fixes to report **All Users** to skip Contacts
- Fixes to report conditional formatting which would format unnessecary fields

## 0.0.26
- Remove Body from **EmailConfiguration** in HTML Report
- Fixes an issue with table in reports having borders (from the EmailBody)

## 0.0.25
- Fix ManagerStatus when Overwriteproperty is in use
- Add HTML report to Email to Admins

```powershell
HTMLReports                        = @(
    # Accepts a list of reports to generate. Can be multiple reprorts having different sections, or just one having it all
    [ordered] @{
        Enable                = $true
        ShowHTML              = $true
        Title                 = "Password Solution Summary"
        Online                = $true
        DisableWarnings       = $true
        ShowConfiguration     = $true
        ShowAllUsers          = $true
        ShowRules             = $true
        ShowUsersSent         = $true
        ShowManagersSent      = $true
        ShowEscalationSent    = $true
        ShowSkippedUsers      = $true
        ShowSkippedLocations  = $true
        ShowSearchUsers       = $true
        ShowSearchManagers    = $true
        ShowSearchEscalations = $true
        FilePath              = "$PSScriptRoot\Reporting\PasswordSolution_$(($Date).ToString('yyyy-MM-dd_HH_mm_ss')).html"
        AttachToEmail         = $true # new option
    }
)
```

## 0.0.24
- Bump dependencies of PSWriteHTML, Mailozaurr to newest versions
- Add support for managers of users being a Contact (with an email field)

## 0.0.23
 - Small cleanup
 - Docs update
## 0.0.22
- Fixes dashboard not showing some objects
- Added properties describing which rule was used to find user
- Added some additional logging