<#
 .SYNOPSIS
  Bulk create member users in a Microsoft Entra (Azure AD) tenant using Microsoft Graph PowerShell.

 .DESCRIPTION
  Reads a CSV describing new internal (member) users and creates each account via Microsoft Graph (POST /users).
  Optional: skip if user already exists, add created (or existing) users to a group, generate passwords when not supplied,
  and produce a structured results log.

  Required CSV columns (case-insensitive):
    UserPrincipalName, DisplayName
  Optional CSV columns (if blank / missing will use defaults or be ignored):
    MailNickname, Password, GivenName, Surname, UsageLocation, JobTitle, Department,
    ForceChangePassword (Y/N), AccountEnabled (True/False), AddToGroupObjectId (per-row override)

  Generated password: If Password empty, script generates a 16 char complex temporary password.
  ForceChangePassword defaults to Y (true) when password is generated unless explicitly set to N.

 .EXAMPLE
  pwsh ./createuser.ps1 -CsvPath ./usercreation.csv -TenantId "contoso.onmicrosoft.com" -AddToGroupObjectId <groupObjectId>

 .EXAMPLE
  pwsh ./createuser.ps1 -CsvPath ./usercreation.csv -SkipExisting -DryRun

 .NOTES
  Required Graph delegated permissions: User.ReadWrite.All (+ Group.ReadWrite.All if adding to a group)
  Run in PowerShell 7+ where possible. Use -DryRun to validate CSV without creating users.
  Does not assign licenses. Ensure UsageLocation is set if you plan to assign licenses later.
#>

param(
    [Parameter(Mandatory)][string]$CsvPath,
    [Parameter()][string]$TenantId = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx',
    [Parameter()][string]$DomainSuffix = 'xxxxx.onmicrosoft.com',
    [Parameter()][string]$LogPath = "$(Split-Path -Parent $MyInvocation.MyCommand.Path)\\createuser-log-$(Get-Date -Format yyyyMMdd-HHmmss).csv",
    [Parameter()][string]$AddToGroupObjectId = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx',
    [Parameter()][switch]$SkipExisting,
    [Parameter()][switch]$DryRun,
    [Parameter()][int]$ThrottleDelaySeconds = 0
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Info($Message){ Write-Host "[INFO ] $Message" -ForegroundColor Cyan }
function Write-Warn($Message){ Write-Warning $Message }
function Write-Err ($Message){ Write-Host "[ERROR] $Message" -ForegroundColor Red }

if (-not (Test-Path $CsvPath)) { throw "CSV file not found: $CsvPath" }

# Ensure Microsoft Graph module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Info 'Installing Microsoft.Graph module (CurrentUser scope)...'
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph -ErrorAction Stop -Verbose

$scopes = @('User.ReadWrite.All')
if ($AddToGroupObjectId) { $scopes += 'Group.ReadWrite.All' }

Write-Info "Connecting to Graph with scopes: $($scopes -join ', ')"
if ($TenantId) { Connect-MgGraph -TenantId $TenantId -Scopes $scopes | Out-Null } else { Connect-MgGraph -Scopes $scopes | Out-Null }
$profile = Get-MgContext
Write-Info "Connected. Tenant: $($profile.TenantId) Account: $($profile.Account)"

Write-Info "Importing CSV: $CsvPath"
$raw = Import-Csv -Path $CsvPath
if (-not $raw -or $raw.Count -eq 0) { throw 'CSV is empty.' }

function New-TempPassword([int]$Length = 16){
    # Ensure complexity: upper, lower, digit, symbol
    $sets = @(
        (65..90),       # A-Z
        (97..122),      # a-z
        (48..57),       # 0-9
        (33,35,36,37,38,42,64,95) # symbols ! # $ % & * @ _
    )
    $chars = @()
    foreach($s in $sets){ $chars += [char]($s | Get-Random) }
    $all = ($sets | ForEach-Object { $_ })
    for($i=$chars.Count; $i -lt $Length; $i++){ $chars += [char]($all | Get-Random) }
    -join ($chars | Sort-Object { Get-Random })
}

# Normalize & enrich rows
$users = $raw | ForEach-Object {
    [PSCustomObject]@{
        UserPrincipalName = "$( [guid]::NewGuid().ToString('N') )@$DomainSuffix"
        DisplayName       = $_.DisplayName
        MailNickname      = if ($_.MailNickname) { $_.MailNickname } else { ($_.UserPrincipalName -split '@')[0] }
        Password          = $_.Password
        GivenName         = $_.GivenName
        Surname           = $_.Surname
        UsageLocation     = if ($_.UsageLocation) { $_.UsageLocation } else { 'US' }
        JobTitle          = $_.JobTitle
        Department        = $_.Department
        ForceChangePassword = $_.ForceChangePassword
        AccountEnabled    = if ($_.AccountEnabled) { [bool]::Parse($_.AccountEnabled) } else { $true }
        RowAddToGroupId   = $_.AddToGroupObjectId
    }
}

$results = New-Object System.Collections.Generic.List[Object]
$seenUpns = [System.Collections.Generic.HashSet[string]]::new()
$rowNum = 0
foreach($u in $users){
    $rowNum++
    $upn = ($u.UserPrincipalName ?? '').Trim()
    $dn  = ($u.DisplayName ?? '').Trim()
    if (-not $upn -or -not $dn){ Write-Warn "Row $rowNum missing required UserPrincipalName or DisplayName. Skipping."; continue }
    if (-not $seenUpns.Add($upn)) { Write-Warn "Duplicate UPN in CSV ($upn). Skipping duplicate."; continue }

    $existingUser = $null
    if ($SkipExisting) {
        try { $existingUser = Get-MgUser -Filter "userPrincipalName eq '$upn'" -ConsistencyLevel eventual -CountVariable null -ErrorAction Stop | Select-Object -First 1 } catch { Write-Warn "Lookup failed for $upn $($_.Exception.Message)" }
    }

    if ($existingUser) {
        Write-Info "Skipping existing user: $upn"
        $results.Add([PSCustomObject]@{ UserPrincipalName=$upn; DisplayName=$existingUser.DisplayName; Status='SkippedExisting'; UserId=$existingUser.Id; AddedToGroup=$false; TempPassword=$null; Error=$null }) | Out-Null
        $targetGroup = if ($u.RowAddToGroupId) { $u.RowAddToGroupId } elseif ($AddToGroupObjectId) { $AddToGroupObjectId } else { $null }
        if ($targetGroup) {
            try {
                Write-Info "Adding existing user to group $targetGroup"
                if (-not $DryRun) { New-MgGroupMember -GroupId $targetGroup -DirectoryObjectId $existingUser.Id -ErrorAction Stop | Out-Null }
                ($results[$results.Count-1]).AddedToGroup = $true
            } catch { Write-Warn "Failed group add for existing $upn $($_.Exception.Message)"; ($results[$results.Count-1]).Error = $_.Exception.Message }
        }
        continue
    }

    # Prepare password profile
    $tempPassword = $null
    $password = $u.Password
    if (-not $password) { $password = New-TempPassword; $tempPassword = $password }
    $forceChange = $true
    if ($u.ForceChangePassword) {
        $forceChange = ($u.ForceChangePassword -match '^(Y|Yes|True)$')
    } elseif ($u.Password) {
        # If user explicitly provided password but no flag, default to true for safety
        $forceChange = $true
    }

    $upBody = @{
        AccountEnabled = $u.AccountEnabled
        DisplayName    = $dn
        MailNickname   = $u.MailNickname
        UserPrincipalName = $upn
        PasswordProfile = @{ forceChangePasswordNextSignIn = $forceChange; password = $password }
    }
    if ($u.GivenName)   { $upBody.GivenName   = $u.GivenName }
    if ($u.Surname)     { $upBody.Surname     = $u.Surname }
    if ($u.JobTitle)    { $upBody.JobTitle    = $u.JobTitle }
    if ($u.Department)  { $upBody.Department  = $u.Department }
    if ($u.UsageLocation){ $upBody.UsageLocation = $u.UsageLocation }

    try {
        if ($DryRun) {
            Write-Info "[DryRun] Would create user $upn"
            $userId = $null
        } else {
            Write-Info "Creating user $upn ..."
            $created = New-MgUser -BodyParameter $upBody -ErrorAction Stop
            $userId = $created.Id
        }
        $groupAdded = $false
        $targetGroup = if ($u.RowAddToGroupId) { $u.RowAddToGroupId } elseif ($AddToGroupObjectId) { $AddToGroupObjectId } else { $null }
        if ($targetGroup -and -not $DryRun -and $userId) {
            try {
                Write-Info "Adding $upn to group $targetGroup"
                New-MgGroupMember -GroupId $targetGroup -DirectoryObjectId $userId -ErrorAction Stop | Out-Null
                $groupAdded = $true
            } catch { Write-Warn "Failed group add for $upn $($_.Exception.Message)" }
        } elseif ($targetGroup -and $DryRun) {
            Write-Info "[DryRun] Would add $upn to group $targetGroup"
        }
        $results.Add([PSCustomObject]@{ UserPrincipalName=$upn; DisplayName=$dn; Status= if($DryRun){'DryRun'} else {'Created'}; UserId=$userId; AddedToGroup=$groupAdded; TempPassword=$tempPassword; Error=$null }) | Out-Null
    }
    catch {
        Write-Err "Creation failed for $upn $($_.Exception.Message)"
        $results.Add([PSCustomObject]@{ UserPrincipalName=$upn; DisplayName=$dn; Status='Error'; UserId=$null; AddedToGroup=$false; TempPassword=$tempPassword; Error=$_.Exception.Message }) | Out-Null
    }

    if ($ThrottleDelaySeconds -gt 0) { Start-Sleep -Seconds $ThrottleDelaySeconds }
}

Write-Info "Writing log to $LogPath"
$results | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8

Write-Host "--- Summary ---" -ForegroundColor Green
$results | Group-Object Status | Select-Object Name,Count | Format-Table -AutoSize
Write-Host "Log file: $LogPath" -ForegroundColor Green

Disconnect-MgGraph | Out-Null
