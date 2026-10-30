# PasswordSolution working rules

This repository contains the PowerShell module. PasswordSolutionX is a separate repository.
Delivery retries, server selection and connection reuse belong in Mailozaurr.
Validate notifications with synthetic users and loopback mail servers; never use real AD or SMTP.

Before adding or changing tests, read the shared authoring gate at
`$CODEX_HOME/skills/test-audit/SKILL.md#authoring-gate` and proof guidance at
`$CODEX_HOME/skills/contract-focused-tests/SKILL.md` (default CODEX_HOME: `~/.codex`).
For an explicitly requested pruning campaign, read `$CODEX_HOME/skills/test-audit/CAMPAIGN.md`.
