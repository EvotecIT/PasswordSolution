# PasswordSolution Release History

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