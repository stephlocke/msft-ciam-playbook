<#
 .SYNOPSIS
  Bulk invite external (guest) users to a Microsoft Entra (Azure AD) tenant using Microsoft Graph PowerShell.

 .DESCRIPTION
  Reads a CSV of users and sends invitations via the Microsoft Graph Invitation API. Optionally:
   - Skips users that already exist (by mail)
   - Adds newly invited (or existing) guests to a specified group
   - Writes a structured log (CSV) with per‑user results

  Expected CSV columns (case‑insensitive):
	Name, InvitedUserEmailAddress
  Optional columns accepted (ignored if not present):
	InvitedUserDisplayName, Message, RedirectUrl

 .EXAMPLE
  pwsh ./inviteuser.ps1 -CsvPath .\userinvitation.csv -TenantId "contoso.onmicrosoft.com" -AddToGroupObjectId "<groupObjectId>" -SkipExisting

 .EXAMPLE
  pwsh ./inviteuser.ps1 -CsvPath .\userinvitation.csv -CustomMessage "Welcome partner" -RedirectUrl "https://myapps.microsoft.com/" -LogPath .\invites-log.csv

 .NOTES
  Required Graph delegated permissions: User.Invite.All (or User.ReadWrite.All) plus Group.ReadWrite.All if adding to a group.
  The script will request scopes on first run. Run in pwsh (PowerShell 7+) where possible.

  Reference: https://learn.microsoft.com/en-us/entra/external-id/bulk-invite-powershell
#>

param(
	[Parameter(Mandatory)][string]$CsvPath,
	[Parameter()][string]$TenantId= 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx',
	[Parameter()][string]$RedirectUrl = 'https://myapplications.microsoft.com/',
	[Parameter()][string]$CustomMessage = 'Hello. You are invited to collaborate.',
	[Parameter()][string]$LogPath = "$(Split-Path -Parent $MyInvocation.MyCommand.Path)\\invite-log-$(Get-Date -Format yyyyMMdd-HHmmss).csv",
	[Parameter()][switch]$SendInvitationMessage = $true,
	[Parameter()][string]$AddToGroupObjectId = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx',
	[Parameter()][switch]$SkipExisting,
	[Parameter()][int]$ThrottleDelaySeconds = 0
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Info($Message){ Write-Host "[INFO ] $Message" -ForegroundColor Cyan }
function Write-Warn($Message){ Write-Warning $Message }
function Write-Err ($Message){ Write-Host "[ERROR] $Message" -ForegroundColor Red }

if (-not (Test-Path $CsvPath)) {
	throw "CSV file not found: $CsvPath"
}

# Ensure Microsoft Graph module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
	Write-Info 'Installing Microsoft.Graph module (CurrentUser scope)...'
	Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph -ErrorAction Stop -Verbose

$scopes = @('User.Invite.All')
if ($AddToGroupObjectId) { $scopes += 'Group.ReadWrite.All' }

Write-Info "Connecting to Graph with scopes: $($scopes -join ', ')"
if ($TenantId) {
	Connect-MgGraph -TenantId $TenantId -Scopes $scopes | Out-Null
} else {
	Connect-MgGraph -Scopes $scopes | Out-Null
}

$profile = Get-MgContext
Write-Info "Connected. Tenant: $($profile.TenantId) Account: $($profile.Account)"

Write-Info "Importing CSV: $CsvPath"
$raw = Import-Csv -Path $CsvPath
if (-not $raw -or $raw.Count -eq 0) { throw 'CSV is empty.' }

# Normalise columns
$invitations = $raw | ForEach-Object {
	[PSCustomObject]@{
		Name                      = $_.Name
		InvitedUserEmailAddress   = $_.InvitedUserEmailAddress
		InvitedUserDisplayName    = if ($_.InvitedUserDisplayName) { $_.InvitedUserDisplayName } else { $_.Name }
		Message                   = if ($_.Message) { $_.Message } else { $CustomMessage }
		RedirectUrl               = $RedirectUrl 
	}
}

## Removed inline coalesce helper; used explicit conditional above for clarity and to avoid forward reference.

# Create message info object once per distinct message body to optimize
$messageCache = @{}
function Get-MessageInfoObject($body){
	if (-not $messageCache.ContainsKey($body)) {
		$msg = [Microsoft.Graph.PowerShell.Models.MicrosoftGraphInvitedUserMessageInfo]::new()
		$msg.CustomizedMessageBody = $body
		$messageCache[$body] = $msg
	}
	return $messageCache[$body]
}

$results = New-Object System.Collections.Generic.List[Object]
$rowNum = 0
foreach ($invite in $invitations) {
	$rowNum++
	$email = $invite.InvitedUserEmailAddress.Trim()
	if (-not $email) {
		Write-Warn "Row $rowNum missing email. Skipping."
		continue
	}

	$existingUser = $null
	if ($SkipExisting) {
		try {
			$existingUser = Get-MgUser -Filter "mail eq '$email' or userPrincipalName eq '$email'" -ConsistencyLevel eventual -CountVariable null -ErrorAction Stop | Select-Object -First 1
		} catch {
			Write-Warn "Lookup failed for $email $($_.Exception.Message)"
		}
	}

	if ($existingUser -and $existingUser.UserType -eq 'Guest') {
		Write-Info "Skipping existing guest: $email"
		$results.Add([PSCustomObject]@{ Email=$email; DisplayName=$existingUser.DisplayName; Status='SkippedExisting'; UserId=$existingUser.Id; AddedToGroup=$false; Error=$null }) | Out-Null
		if ($AddToGroupObjectId) {
			try {
				Write-Info "Adding existing guest to group $AddToGroupObjectId"
				New-MgGroupMember -GroupId $AddToGroupObjectId -DirectoryObjectId $existingUser.Id -ErrorAction Stop | Out-Null
				($results[$results.Count-1]).AddedToGroup = $true
			} catch {
				Write-Warn "Failed to add existing guest $email to group: $($_.Exception.Message)"
				($results[$results.Count-1]).Error = $_.Exception.Message
			}
		}
		continue
	}

	$messageInfo = Get-MessageInfoObject $invite.Message
	$invitedUserDisplayName = if ($invite.InvitedUserDisplayName) { $invite.InvitedUserDisplayName } else { $invite.Name }

	try {
		Write-Info "Inviting $email ..."
		$invitation = New-MgInvitation -InvitedUserEmailAddress $email `
			-InvitedUserDisplayName $invitedUserDisplayName `
			-InviteRedirectUrl $invite.RedirectUrl `
			-InvitedUserMessageInfo $messageInfo `
			-SendInvitationMessage:$SendInvitationMessage.IsPresent -ErrorAction Stop

		$userId = $invitation.InvitedUser.Id
		$groupAdded = $false
		if ($AddToGroupObjectId -and $userId) {
			try {
				Write-Info "Adding $email to group $AddToGroupObjectId"
				New-MgGroupMember -GroupId $AddToGroupObjectId -DirectoryObjectId $userId -ErrorAction Stop | Out-Null
				$groupAdded = $true
			} catch {
				Write-Warn "Failed group add for $email $($_.Exception.Message)"
			}
		}

		$results.Add([PSCustomObject]@{ Email=$email; DisplayName=$invitedUserDisplayName; Status='Invited'; UserId=$userId; AddedToGroup=$groupAdded; Error=$null }) | Out-Null
	}
	catch {
		Write-Err "Invitation failed for $email $($_.Exception.Message)"
		$results.Add([PSCustomObject]@{ Email=$email; DisplayName=$invitedUserDisplayName; Status='Error'; UserId=$null; AddedToGroup=$false; Error=$_.Exception.Message }) | Out-Null
	}

	if ($ThrottleDelaySeconds -gt 0) { Start-Sleep -Seconds $ThrottleDelaySeconds }
}

Write-Info "Writing log to $LogPath"
$results | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8

Write-Host "--- Summary ---" -ForegroundColor Green
$results | Group-Object Status | Select-Object Name,Count | Format-Table -AutoSize
Write-Host "Log file: $LogPath" -ForegroundColor Green

Disconnect-MgGraph | Out-Null

