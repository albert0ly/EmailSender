# Microsoft 365 / Entra ID Setup for EmailSender

This guide explains how to configure an application registration and permissions so the EmailSender API can send email from a specific mailbox using Microsoft Graph with no user interaction (client credentials flow).

## Prerequisites
- Microsoft 365 tenant with Exchange Online
- Role: Global Administrator or Application Administrator (to register app + grant admin consent)

## 1. Register the application
1. Open Microsoft Entra admin center: https://entra.microsoft.com
2. Go to: Applications > App registrations > New registration
3. Name: `EmailSender` (any name)
4. Supported account types: Accounts in this organizational directory only (Single tenant)
5. Redirect URI: Not required
6. Click Register
7. Copy the following values for appsettings.json:
   - Directory (tenant) ID => `TenantId`
   - Application (client) ID => `ClientId`

## 2. Create a client secret
1. In the app registration, go to Certificates & secrets
2. Under Client secrets, click New client secret
3. Add a description (e.g., `server-secret`) and choose an expiration
4. Click Add, then copy the Secret value immediately => `ClientSecret`

## 3. Add Microsoft Graph application permissions
1. In the app registration, go to API permissions
2. Click Add a permission > Microsoft Graph > Application permissions
3. Search and add: `Mail.Send`
4. Click Grant admin consent for your tenant and confirm

Notes:
- `Mail.Send` app permission allows the app to send email as any mailbox by default. If you want to restrict which mailboxes can be accessed, configure an Exchange Online Application Access Policy.

## 4. Optional: Restrict mailbox access with Application Access Policy
To limit the app to specific mailboxes only:
- Connect to Exchange Online PowerShell
- Create a mail-enabled security group containing the allowed mailbox(es)
- Create an application access policy binding the app registration to the group

Example PowerShell (adjust IDs and names):
```
Connect-ExchangeOnline

# Create a mail-enabled security group (or use an existing one)
New-DistributionGroup -Name "EmailSenderAllowedMailboxes" -Type Security

# Add the target mailbox to the group
Add-DistributionGroupMember -Identity "EmailSenderAllowedMailboxes" -Member user@contoso.com

# Bind app registration to the group (use the app's ClientId)
New-ApplicationAccessPolicy -AppId <CLIENT_ID> -PolicyScopeGroupId EmailSenderAllowedMailboxes@contoso.com -AccessRight RestrictAccess -Description "Limit EmailSender to specific mailboxes"

# Test policy
Test-ApplicationAccessPolicy -AppId <CLIENT_ID> -Identity user@contoso.com
```
Docs: https://learn.microsoft.com/exchange/troubleshoot/administration/application-access-policy

## 5. Configure EmailSender/appsettings.json
```
"Graph": {
  "TenantId": "<YOUR_TENANT_ID>",
  "ClientId": "<YOUR_APP_CLIENT_ID>",
  "ClientSecret": "<YOUR_APP_CLIENT_SECRET>",
  "MailboxAddress": "<SENDER_MAILBOX_ADDRESS>"
}
```
- `MailboxAddress`: The mailbox to send from (UPN or user ID).

## 6. Test the API
1. Build and run the project:
```
dotnet build

dotnet run --project EmailSender
```
2. Send a request using curl or Postman to `POST https://localhost:5001/api/email/send` with multipart/form-data including recipients, subject, body, and attachments.

## Troubleshooting
- 401/403 from Graph: Verify TenantId/ClientId/ClientSecret, app permission `Mail.Send`, and admin consent granted.
- Not allowed to send as mailbox: If using access policies, ensure the mailbox is in the allowed group; otherwise check mailbox existence.
- Large attachments failing: Ensure network stability; the API uses 5MB chunks. Very large uploads can be sensitive to timeouts on proxies or gateways.
- SSL issues locally: Trust the development certificate: `dotnet dev-certs https --trust`
