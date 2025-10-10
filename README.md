# Micrsoft Entra External tenants runbooks

A set of runbooks for managing Microsoft Entra External tenants. These runbooks are designed to automate various operations such as creating, retrieving, listing, updating, and deleting external tenants.

## http-scripts
This project contains a set of HTTP scripts for interacting with the CIAM Tenants API. [http-scripts/README.md](http-scripts/README.md) provides detailed instructions on how to use these scripts.

## Bulk guest user invitation (PowerShell)

Script: `ps/inviteuser.ps1`

Invites external (guest) users in bulk using Microsoft Graph. Accepts a CSV containing at minimum:

```
Name,InvitedUserEmailAddress
Jane Partner,jane.partner@example.com
Dev Tester,dev.tester@example.org
```

### Prerequisites
* PowerShell 7+ (recommended) or Windows PowerShell 5.1
* Microsoft Graph PowerShell SDK (`Install-Module Microsoft.Graph -Scope CurrentUser`)
* Permissions: The signed-in account must be able to consent to (or already have) `User.Invite.All` (and `Group.ReadWrite.All` if adding to a group)

### Usage examples
Invite users and send emails:

```
pwsh ./ps/inviteuser.ps1 -CsvPath ./ps/userinvitation.csv -TenantId "contoso.onmicrosoft.com"
```

Invite users, skip existing guests, add to a group, and specify a custom message:

```
pwsh ./ps/inviteuser.ps1 -CsvPath ./ps/userinvitation.csv -TenantId "contoso.onmicrosoft.com" -AddToGroupObjectId "<groupObjectId>" -SkipExisting -CustomMessage "Welcome to Contoso partner portal" -RedirectUrl "https://myapplications.microsoft.com/"
```

Disable sending the invitation email (user created in pending acceptance state):

```
pwsh ./ps/inviteuser.ps1 -CsvPath ./ps/userinvitation.csv -SendInvitationMessage:$false
```

After execution, a log CSV is produced (path shown in script output) summarising status per email.

### CSV optional columns
You may optionally include `InvitedUserDisplayName`, `Message`, or `RedirectUrl` per row to override the script defaults.

### Verification
List guest users:

```
Get-MgUser -Filter "userType eq 'Guest'" | Select DisplayName,Mail,UserPrincipalName
```

Remove a test guest (cleanup):

```
Remove-MgUser -UserId <objectId>
```

## Bulk member user creation (PowerShell)

Script: `ps/createuser.ps1`

Creates internal (member) users in bulk from a CSV. Required columns:

```
UserPrincipalName,DisplayName
```

Optional columns: `MailNickname,Password,GivenName,Surname,UsageLocation,JobTitle,Department,ForceChangePassword,AccountEnabled,AddToGroupObjectId`

If `Password` is blank a complex temporary password is generated and users are forced to change it at next sign in (unless `ForceChangePassword` = N/False). A per‑row `AddToGroupObjectId` can override the global `-AddToGroupObjectId` parameter.

Example CSV row (see `ps/usercreation.csv.example`):

```
alex.jones@contoso.onmicrosoft.com,Alex Jones,alexj,,Alex,Jones,US,Security Analyst,Security,Y,True,
```

### Usage examples
Create users (generating passwords where not supplied):

```
pwsh ./ps/createuser.ps1 -CsvPath ./ps/usercreation.csv -TenantId "contoso.onmicrosoft.com"
```

Dry run (validate only, no creation) and show intended actions:

```
pwsh ./ps/createuser.ps1 -CsvPath ./ps/usercreation.csv -DryRun
```

Create users, skip existing, and add all to a group:

```
pwsh ./ps/createuser.ps1 -CsvPath ./ps/usercreation.csv -SkipExisting -AddToGroupObjectId <groupObjectId>
```

### Output & logging
Produces a log CSV with: `UserPrincipalName,DisplayName,Status,UserId,AddedToGroup,TempPassword,Error`.
`TempPassword` is only populated for generated passwords—store it securely and rotate after first sign‑in.

### Security notes
Avoid storing plaintext passwords in source control. Prefer leaving `Password` blank to let the script generate temporary credentials, then distribute via a secure channel.

