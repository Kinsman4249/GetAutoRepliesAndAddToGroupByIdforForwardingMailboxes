<#
.SYNOPSIS
  Find mailboxes with Automatic Replies AND forwarding enabled, then add them to a group.

.DESCRIPTION
  1) Connects to Exchange Online
  2) Collects all mailboxes that have BOTH:
       - AutoReplyState != Disabled (Enabled or Scheduled)
       - Forwarding configured (ForwardingAddress or ForwardingSmtpAddress is set)
  3) Attempts to add users to the target group without Graph first:
       - If it's a Microsoft 365 (Unified) group: Add-UnifiedGroupLinks
       - Else if it's a distribution/mail-enabled security group: Add-DistributionGroupMember
  4) If neither EXO method applies, falls back to direct Microsoft Graph REST calls
     via Invoke-RestMethod (device code auth). No Graph SDK modules required.

.PARAMETER GroupObjectId
  Optional. The Object ID (GUID) of the target group. If not provided, the script
  will prompt interactively.

.EXAMPLE
  .\GetAutoRepliesAndAddToGroupById.ps1
  # Prompts for the group GUID interactively

.EXAMPLE
  .\GetAutoRepliesAndAddToGroupById.ps1 -GroupObjectId "<GROUP_ID>"
  # Uses the provided GUID directly, no prompt

.NOTES
  PowerShell: 5.1 (Desktop)
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$GroupObjectId = "",

    [Parameter(Mandatory = $false)]
    [ValidateSet("UserMailbox","SharedMailbox","RoomMailbox","EquipmentMailbox","All")]
    [string]$MailboxScope = "All",

    [Parameter(Mandatory = $false)]
    [string]$ReportPath = (Join-Path $PWD ("AutoReplies_Fwd_ToGroup_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date)))
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Resolve Group Object ID (accept from CLI or prompt)

if (-not $GroupObjectId -or $GroupObjectId.Trim() -eq '') {
    do {
        $GroupObjectId = (Read-Host "Enter the Object ID (GUID) of the target group").Trim()
        try {
            $null = [System.Guid]::Parse($GroupObjectId)
            $validGuid = $true
        } catch {
            $validGuid = $false
            Write-Host "Invalid GUID format. Please enter a valid Object ID (e.g. <GROUP_ID>)" -ForegroundColor Red
        }
    } while (-not $validGuid)
} else {
    try {
        $null = [System.Guid]::Parse($GroupObjectId)
    } catch {
        throw "Invalid GroupObjectId parameter: '$GroupObjectId' is not a valid GUID."
    }
}

Write-Host "Target group: $GroupObjectId" -ForegroundColor Cyan

#endregion

#region Helper: TLS + Module handling

function Enable-Tls12 {
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    } catch {
        Write-Verbose "Could not set TLS 1.2: $($_.Exception.Message)"
    }
}

function Ensure-PSGallery {
    try {
        $repo = Get-PSRepository -Name "PSGallery" -ErrorAction Stop
        if ($repo.InstallationPolicy -ne "Trusted") {
            Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted -ErrorAction Stop
        }
    } catch {
        Write-Verbose "PSGallery repository check failed: $($_.Exception.Message)"
    }
}

function Ensure-Module {
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$false)][version]$MinVersion = [version]"0.0"
    )

    Enable-Tls12
    Ensure-PSGallery

    $installed = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1

    if (-not $installed) {
        Write-Host "Installing module: $Name"
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    } elseif ([version]$installed.Version -lt $MinVersion) {
        Write-Host "Upgrading module: $Name (installed $($installed.Version), need >= $MinVersion)"
        try {
            Update-Module -Name $Name -Force -ErrorAction Stop
        } catch {
            Write-Verbose "Update-Module failed for $Name, attempting Install-Module -Force: $($_.Exception.Message)"
            Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
    }

    try {
        Import-Module -Name $Name -Force -ErrorAction Stop
    } catch {
        try { Remove-Module -Name $Name -Force -ErrorAction SilentlyContinue } catch {}
        Import-Module -Name $Name -Force -ErrorAction Stop
    }
}

#endregion Helper

#region Graph REST helpers (no SDK needed)

function Get-GraphTokenDeviceCode {
    param(
        [string[]]$Scopes = @("User.Read.All","GroupMember.ReadWrite.All")
    )

    Enable-Tls12

    $clientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
    $tenantId = "organizations"
    $scopeString = ($Scopes + "offline_access") -join " "

    $deviceCodeResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/devicecode" -Body @{
        client_id = $clientId
        scope     = $scopeString
    }

    Write-Host ""
    Write-Host $deviceCodeResponse.message -ForegroundColor Yellow
    Write-Host ""

    $tokenBody = @{
        grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
        client_id   = $clientId
        device_code = $deviceCodeResponse.device_code
    }

    $timeout = (Get-Date).ToUniversalTime().AddSeconds($deviceCodeResponse.expires_in)
    $interval = $deviceCodeResponse.interval
    if ($interval -lt 5) { $interval = 5 }

    while ((Get-Date).ToUniversalTime() -lt $timeout) {
        Start-Sleep -Seconds $interval
        try {
            $tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $tokenBody -ErrorAction Stop
            Write-Host "Graph authentication successful." -ForegroundColor Green
            return $tokenResponse.access_token
        } catch {
            $errDetail = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
            if ($errDetail.error -eq "authorization_pending") {
                continue
            } elseif ($errDetail.error -eq "slow_down") {
                $interval += 5
                continue
            } else {
                throw "Device code auth failed: $($errDetail.error_description)"
            }
        }
    }
    throw "Device code authentication timed out."
}

function Invoke-GraphGet {
    param([string]$Token, [string]$Uri)
    $headers = @{ Authorization = "Bearer $Token"; "Content-Type" = "application/json" }
    Invoke-RestMethod -Method Get -Uri $Uri -Headers $headers -ErrorAction Stop
}

function Invoke-GraphPost {
    param([string]$Token, [string]$Uri, [string]$Body)
    $headers = @{ Authorization = "Bearer $Token"; "Content-Type" = "application/json" }
    Invoke-RestMethod -Method Post -Uri $Uri -Headers $headers -Body $Body -ErrorAction Stop
}

#endregion Graph REST helpers

#region Connect to Exchange Online and collect matching mailboxes

Ensure-Module -Name "ExchangeOnlineManagement" -MinVersion ([version]"3.0.0")

Write-Host "Connecting to Exchange Online..."
Connect-ExchangeOnline -ShowBanner:$false

$recipientTypeDetails = @()
switch ($MailboxScope) {
    "UserMailbox"      { $recipientTypeDetails = @("UserMailbox") }
    "SharedMailbox"    { $recipientTypeDetails = @("SharedMailbox") }
    "RoomMailbox"      { $recipientTypeDetails = @("RoomMailbox") }
    "EquipmentMailbox" { $recipientTypeDetails = @("EquipmentMailbox") }
    "All"              { $recipientTypeDetails = @("UserMailbox","SharedMailbox","RoomMailbox","EquipmentMailbox") }
}

Write-Host "Enumerating mailboxes ($MailboxScope)..."
$mailboxes = @()

$selectProps = @(
    "DisplayName",
    "UserPrincipalName",
    "PrimarySmtpAddress",
    "RecipientTypeDetails",
    "ForwardingAddress",
    "ForwardingSmtpAddress",
    "DeliverToMailboxAndForward"
)

try {
    $mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails $recipientTypeDetails -Properties ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward |
        Select-Object $selectProps
} catch {
    Write-Verbose "Get-EXOMailbox failed; falling back to Get-Mailbox: $($_.Exception.Message)"
    $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails $recipientTypeDetails |
        Select-Object $selectProps
}

$forwardingMailboxes = @($mailboxes | Where-Object {
    ($_.ForwardingAddress -and $_.ForwardingAddress.ToString().Trim() -ne '') -or
    ($_.ForwardingSmtpAddress -and $_.ForwardingSmtpAddress.ToString().Trim() -ne '')
})

Write-Host ("Total mailboxes: {0}  |  With forwarding enabled: {1}" -f $mailboxes.Count, $forwardingMailboxes.Count)
Write-Host ("Checking Automatic Replies settings for {0} forwarding-enabled mailbox(es)..." -f $forwardingMailboxes.Count)

$matchedMailboxes = New-Object System.Collections.Generic.List[object]
$idx = 0

foreach ($mbx in $forwardingMailboxes) {
    $idx++
    if ($idx % 25 -eq 0) {
        Write-Progress -Activity "Reading Automatic Replies (forwarding-enabled only)" -Status "$idx / $($forwardingMailboxes.Count)" -PercentComplete (($idx / [double]$forwardingMailboxes.Count) * 100)
    }

    $id = $mbx.UserPrincipalName
    if (-not $id -or $id.Trim() -eq '') { continue }

    try {
        $cfg = Get-MailboxAutoReplyConfiguration -Identity $id
        if ($cfg.AutoReplyState -and $cfg.AutoReplyState.ToString() -ne "Disabled") {
            $matchedMailboxes.Add([pscustomobject]@{
                UPN                       = $id
                PrimarySmtp               = $mbx.PrimarySmtpAddress
                DisplayName               = $mbx.DisplayName
                RecipientType             = $mbx.RecipientTypeDetails
                AutoReplyState            = $cfg.AutoReplyState
                StartTime                 = $cfg.StartTime
                EndTime                   = $cfg.EndTime
                ForwardingAddress         = $mbx.ForwardingAddress
                ForwardingSmtpAddress     = $mbx.ForwardingSmtpAddress
                DeliverToMailboxAndForward = $mbx.DeliverToMailboxAndForward
            })
        }
    } catch {
        Write-Verbose "Failed Get-MailboxAutoReplyConfiguration for $id : $($_.Exception.Message)"
    }
}

Write-Progress -Activity "Reading Automatic Replies (forwarding-enabled only)" -Completed

Write-Host ("Found {0} mailbox(es) with BOTH Automatic Replies and Forwarding enabled." -f $matchedMailboxes.Count)

#endregion

#region Add to group: try EXO methods first, fallback to Graph REST

$addMethod = $null

try {
    $null = Get-UnifiedGroup -Identity $GroupObjectId -ErrorAction Stop
    $addMethod = "UnifiedGroup"
    Write-Host "Target group appears to be a Microsoft 365 (Unified) group. Will use Add-UnifiedGroupLinks."
} catch {}

if (-not $addMethod) {
    try {
        $null = Get-DistributionGroup -Identity $GroupObjectId -ErrorAction Stop
        $addMethod = "DistributionGroup"
        Write-Host "Target group appears to be a Distribution/Mail-enabled Security group. Will use Add-DistributionGroupMember."
    } catch {}
}

if (-not $addMethod) {
    $addMethod = "GraphREST"
    Write-Host "Target group not manageable via EXO group cmdlets. Will use direct Graph REST calls."
}

$results = New-Object System.Collections.Generic.List[object]

function Add-WithUnifiedGroupLinks {
    param([string]$GroupId, [string]$Upn)
    try {
        if ($PSCmdlet.ShouldProcess("$Upn", "Add to Unified Group $GroupId")) {
            $null = Add-UnifiedGroupLinks -Identity $GroupId -LinkType Members -Links $Upn -ErrorAction Stop
        }
        return @{ Success = $true; Message = "Added (UnifiedGroup)" }
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match "is already a member" -or $msg -match "already exists") {
            return @{ Success = $true; Message = "Already member (UnifiedGroup)" }
        }
        return @{ Success = $false; Message = $msg }
    }
}

function Add-WithDistributionGroupMember {
    param([string]$GroupId, [string]$Upn)
    try {
        if ($PSCmdlet.ShouldProcess("$Upn", "Add to Distribution Group $GroupId")) {
            $null = Add-DistributionGroupMember -Identity $GroupId -Member $Upn -BypassSecurityGroupManagerCheck -ErrorAction Stop
        }
        return @{ Success = $true; Message = "Added (DistributionGroup)" }
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match "is already a member" -or $msg -match "already exists") {
            return @{ Success = $true; Message = "Already member (DistributionGroup)" }
        }
        return @{ Success = $false; Message = $msg }
    }
}

function Add-WithGraphREST {
    param([string]$Token, [string]$GroupId, [string]$Upn)

    try {
        $encodedUpn = [System.Uri]::EscapeDataString($Upn)
        $userUri = "https://graph.microsoft.com/v1.0/users/$encodedUpn?`$select=id"
        $user = Invoke-GraphGet -Token $Token -Uri $userUri

        $memberUri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members/`$ref"
        $body = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.id)" } | ConvertTo-Json

        if ($PSCmdlet.ShouldProcess("$Upn", "Add to Entra group $GroupId via Graph REST")) {
            $null = Invoke-GraphPost -Token $Token -Uri $memberUri -Body $body
        }
        return @{ Success = $true; Message = "Added (GraphREST)" }
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match "One or more added object references already exist" -or $msg -match "already exist") {
            return @{ Success = $true; Message = "Already member (GraphREST)" }
        }
        return @{ Success = $false; Message = $msg }
    }
}

$graphToken = $null
if ($addMethod -eq "GraphREST") {
    Write-Host "Authenticating to Microsoft Graph via device code..."
    $graphToken = Get-GraphTokenDeviceCode -Scopes @("User.Read.All","GroupMember.ReadWrite.All")
}

Write-Host ("Adding {0} user(s) to group using method: {1}" -f $matchedMailboxes.Count, $addMethod)

foreach ($entry in $matchedMailboxes) {
    $upn = $entry.UPN
    $outcome = $null

    switch ($addMethod) {
        "UnifiedGroup"      { $outcome = Add-WithUnifiedGroupLinks -GroupId $GroupObjectId -Upn $upn }
        "DistributionGroup" { $outcome = Add-WithDistributionGroupMember -GroupId $GroupObjectId -Upn $upn }
        "GraphREST"         { $outcome = Add-WithGraphREST -Token $graphToken -GroupId $GroupObjectId -Upn $upn }
    }

    $results.Add([pscustomobject]@{
        UPN                       = $entry.UPN
        PrimarySmtp               = $entry.PrimarySmtp
        DisplayName               = $entry.DisplayName
        RecipientType             = $entry.RecipientType
        AutoReplyState            = $entry.AutoReplyState
        StartTime                 = $entry.StartTime
        EndTime                   = $entry.EndTime
        ForwardingAddress         = $entry.ForwardingAddress
        ForwardingSmtpAddress     = $entry.ForwardingSmtpAddress
        DeliverToMailboxAndForward = $entry.DeliverToMailboxAndForward
        AddMethod                 = $addMethod
        Success                   = $outcome.Success
        ResultMessage             = $outcome.Message
    })
}

$results | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $ReportPath
Write-Host "Report written to: $ReportPath"

#endregion

#region Cleanup

try { Disconnect-ExchangeOnline -Confirm:$false } catch {}

#endregion
