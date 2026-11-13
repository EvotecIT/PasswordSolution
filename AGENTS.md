# PasswordSolution

PowerShell module that sends password expiry notifications to users, managers,
administrators and security teams from configurable rules and templates, and
reports on password state and quality for Active Directory and Entra ID.
PasswordSolutionX (the C# successor) lives in a separate repository.

## Layout

- `Public/`: exported functions, one per file. `Start-PasswordSolution` is the
  notification entry point; `New-PasswordConfiguration*` build its configuration.
- `Private/`: rule processing, notification senders (`Send-Password*`), report
  and search-history helpers.
- `Tests/`: Pester 5 tests (`NotificationDelivery.Tests.ps1`).
- `Build/Manage-Module.ps1`: build and docs through PSPublishModule.
- `Docs/`: generated command help. `Examples/`: usage examples.

## Build and test

```powershell
Invoke-Pester -Path Tests                              # needs Pester 5 and Mailozaurr 3.0.6
./Build/Manage-Module.ps1 -NoSign                      # build into Artefacts/ without signing
./Build/Manage-Module.ps1 -RunMode Documentation -NoSign   # regenerate Docs/
```

Without `-NoSign` the build signs with a code-signing certificate that must be
installed locally. `-RunMode Publish` is release-only.

CI (`.github/workflows/tests.yml`) runs the tests on Windows PowerShell 5.1 and
PowerShell 7.

## Generated files

`Docs/*.md` come from the comment-based help in `Public/`. Edit the help and
rerun the Documentation mode. Don't edit `Docs/` by hand. `Artefacts/` is
build output and ignored.

## Conventions

- Keep Windows PowerShell 5.1 compatibility.
- K&R/OTBS braces, 4-space indentation, help inside the function.
- Mail delivery (retries, server selection and failover, connection reuse)
  belongs in Mailozaurr. This module calls it; don't reimplement delivery here.
- DSInternals and Microsoft Graph modules are optional (password quality and
  Entra). Don't add them as required modules.
- Commit and PR titles use Conventional Commits with Gitmoji:
  `fix(PasswordSolution): 🐛 ...`.

## Gotchas

- Validate notifications with synthetic users and a loopback or mocked mail
  server. Never run against real Active Directory, Entra or SMTP. The tests stub
  `Send-EmailMessage` and `Disable-ADAccount` to fail if reached.
