---
name: Bug report
about: Report a problem with the script, configuration, or documentation
title: "[bug] "
labels: bug
assignees: ''
---

## What happened

<!-- A clear and concise description of the problem. -->

## What you expected to happen

<!-- Describe the expected behavior. -->

## Steps to reproduce

1.
2.
3.

## Environment

- Windows version:
- PowerShell version (output of `$PSVersionTable.PSVersion`):
- `ExchangeOnlineManagement` module version (`Get-Module ExchangeOnlineManagement -ListAvailable | Select Version`):
- Project version (release tag or commit SHA):
- Tenant type (production / dev / test) and approximate mailbox count:
- Was the Graph REST fallback triggered? (yes / no / not sure)

## Command used

```powershell
<!-- Paste the exact command line you ran, redacting tenant-specific GUIDs / UPNs. -->
```

## Output / logs

<details>
<summary>Console output of the failing run</summary>

```text
<paste the full output here, redacting UPNs, group GUIDs, tenant IDs, and any other sensitive data>
```

</details>

<details>
<summary>Relevant rows from the CSV report (if any)</summary>

```text
<paste here, redacting UPNs / SMTP addresses>
```

</details>

## Anything else?

<!-- Other context, screenshots, related issues, your hypothesis on the cause. -->
