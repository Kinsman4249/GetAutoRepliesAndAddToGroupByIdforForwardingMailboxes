# GetAutoRepliesAndAddToGroupByIdforForwardingMailboxes

PowerShell script to find Exchange Online mailboxes with Automatic Replies enabled and forwarding configured, then add them to a specified Entra ID group. Uses Exchange Online cmdlets first and only falls back to Microsoft Graph REST calls if required - no Graph SDK modules needed.

## Requirements

- Windows PowerShell 5.1
- [ExchangeOnlineManagement](https://www.powershellgallery.com/packages/ExchangeOnlineManagement) v3.0.0+ (installed automatically if missing)
- An account with permissions to:
  - Read mailbox configurations (Exchange Administrator or equivalent)
  - Manage group membership on the target group (Group Owner, Groups Administrator, or equivalent)
  - If Graph REST fallback is used: `User.Read.All` and `GroupMember.ReadWrite.All` delegated permissions

## Usage

### Interactive (prompts for the group GUID)

```powershell
.\GetAutoRepliesAndAddToGroupById.ps1
```

### Command line (skips the prompt)

```powershell
.\GetAutoRepliesAndAddToGroupById.ps1 -GroupObjectId "<GROUP_ID>"
```

### Scoped to a specific mailbox type

```powershell
.\GetAutoRepliesAndAddToGroupById.ps1 -MailboxScope UserMailbox
```

### Custom report path

```powershell
.\GetAutoRepliesAndAddToGroupById.ps1 -ReportPath "C:\Reports\auto_reply_report.csv"
```

## Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `GroupObjectId` | No | *(prompts)* | Object ID (GUID) of the target Entra ID group. If omitted, the script prompts interactively with GUID validation. |
| `MailboxScope` | No | `All` | Filter mailbox types. Valid values: `UserMailbox`, `SharedMailbox`, `RoomMailbox`, `EquipmentMailbox`, `All`. |
| `ReportPath` | No | `.\AutoReplies_Fwd_ToGroup_<timestamp>.csv` | Path for the CSV report output. |

## How It Works

1. **Connects to Exchange Online** using `Connect-ExchangeOnline`.
2. **Enumerates all mailboxes** of the specified type(s) and pulls forwarding properties (`ForwardingAddress`, `ForwardingSmtpAddress`, `DeliverToMailboxAndForward`).
3. **Filters to forwarding-enabled mailboxes only**, then checks each for Automatic Replies (`Enabled` or `Scheduled`) via `Get-MailboxAutoReplyConfiguration`.
4. **Determines the best method** to add members to the target group:
   - **Microsoft 365 (Unified) group** - uses `Add-UnifiedGroupLinks`
   - **Distribution / mail-enabled security group** - uses `Add-DistributionGroupMember`
   - **Security group / not resolvable via EXO** - falls back to direct Microsoft Graph REST calls using device code authentication (no Graph SDK modules required)
5. **Adds matched users** to the group. Already-existing members are handled gracefully.
6. **Exports a CSV report** with matched mailboxes, forwarding details, and the result of each group membership operation.

## Output

The CSV report includes the following columns:

| Column | Description |
|--------|-------------|
| `UPN` | User Principal Name |
| `PrimarySmtp` | Primary SMTP address |
| `DisplayName` | Mailbox display name |
| `RecipientType` | Mailbox type (UserMailbox, SharedMailbox, etc.) |
| `AutoReplyState` | `Enabled` or `Scheduled` |
| `StartTime` | Scheduled auto-reply start (if applicable) |
| `EndTime` | Scheduled auto-reply end (if applicable) |
| `ForwardingAddress` | Internal forwarding target |
| `ForwardingSmtpAddress` | External SMTP forwarding target |
| `DeliverToMailboxAndForward` | Whether mail is also delivered to the original mailbox |
| `AddMethod` | Method used (`UnifiedGroup`, `DistributionGroup`, or `GraphREST`) |
| `Success` | `True` or `False` |
| `ResultMessage` | Outcome detail (added, already member, or error message) |

## Known Limitations

- `Get-MailboxAutoReplyConfiguration` is a per-mailbox call with no bulk alternative. For large tenants the enumeration step can take time. The script pre-filters to forwarding-enabled mailboxes before checking auto-reply state to reduce the number of calls.
- If the target group uses **dynamic membership**, direct member adds will fail (membership is rule-based). The error will be captured in the CSV report.
- The Graph REST fallback uses the public client ID for Microsoft Graph CLI (`14d82eec-204b-4c2f-b7e8-296a70dab67e`). If your tenant restricts public client flows, this will not work and you will need to register your own app.

## License

[LICENSE](LICENSE)
