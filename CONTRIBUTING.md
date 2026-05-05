# Contributing to GetAutoRepliesAndAddToGroupByIdforForwardingMailboxes

Thanks for considering a contribution! This is a small project and the process is intentionally lightweight.

## How to report a bug

Open an issue using the **Bug report** template. Please include:

- The version of this project you're running (commit SHA or release tag)
- The exact command line you used (e.g. `.\GetAutoRepliesAndAddToGroupById.ps1 -GroupObjectId "..." -MailboxScope UserMailbox`)
- The full PowerShell error output (redact tenant identifiers, UPNs, group GUIDs, and any other sensitive data)
- Relevant log excerpts from the generated CSV report or the console transcript
- Your runtime environment: PowerShell version (`$PSVersionTable`), `ExchangeOnlineManagement` module version (`Get-Module ExchangeOnlineManagement -ListAvailable`), Windows version, and whether the Graph REST fallback was triggered

## How to propose a feature

Open an issue using the **Feature request** template. Describe the use case before the implementation — knowing *why* is more useful than *what* in early discussion.

## How to submit a change

1. **Fork** the repo and create a feature branch (`git checkout -b feat/short-description`).
2. **Make your change.** Keep changes focused — one logical change per PR.
3. **Test it.** See the [Testing](#testing) section below for how to verify your change locally.
4. **Lint it.** Run [PSScriptAnalyzer](https://learn.microsoft.com/powershell/utility-modules/psscriptanalyzer/overview) against any modified `.ps1` files: `Invoke-ScriptAnalyzer -Path .\GetAutoRepliesAndAddToGroupByIdforForwardedMailboxes.ps1`.
5. **Update documentation.** If your change alters user-visible behavior (parameters, output columns, supported group types, etc.), update the `README.md`.
6. **Open a PR** against `main`. Fill in the PR template.

## Testing

This script touches Exchange Online and Microsoft Graph, so testing requires a tenant you can safely make changes in. Recommended approach:

- **Use a non-production tenant** (developer tenant or test tenant) whenever possible.
- **Test against a small, dedicated test group** in Entra ID — never a production distribution list — and verify members are added/removed as expected.
- **Cover all three add-paths** when possible: a Microsoft 365 (Unified) group, a Distribution group, and a security group (which forces the Graph REST fallback).
- **Verify the CSV report** is produced and contains the expected columns (`UPN`, `AddMethod`, `Success`, `ResultMessage`, etc.) for both successful and failing rows.
- **Re-run the script** against the same group to confirm idempotent behavior — already-existing members should be reported gracefully, not as errors.
- **Test parameter variants** when relevant to your change: `-GroupObjectId`, `-MailboxScope` (each value), `-ReportPath`.

If your change affects authentication paths, also verify on a tenant where public-client device-code flows are restricted (the Graph REST fallback documents this limitation in the README).

## Coding conventions

- Follow the style of the surrounding PowerShell — verb-noun cmdlet naming, PascalCase parameter names, consistent indentation, comment-based help blocks where they exist.
- Comments should explain *why*, not *what*. The diff already shows what.
- **No new module dependencies** without strong justification. The deliberate design choice of this project is to lean on `ExchangeOnlineManagement` first and fall back to direct Graph REST calls — not to require the full `Microsoft.Graph` SDK.
- **No telemetry, ever.** This script must not phone home.

## Commit messages

Conventional Commits style is preferred but not required:

```
feat: add support for X
fix: handle empty response from Y
docs: clarify setup for Z
```

Keep the subject under 72 characters. Add a body if the change isn't obvious from the diff.

## Releases

Maintainers cut releases by tagging `vX.Y.Z` on `main`. The release workflow in `.github/workflows/release.yml` automatically bundles the script and attaches the archives to a GitHub Release.

Pre-1.0 versioning rules:

- `0.X.0` for any user-visible change
- `0.X.Y` for bug-fix-only patch releases

After 1.0, standard SemVer applies.
