# EmailSender

ASP.NET Core Web API that sends emails through Microsoft Graph using application permissions and a specific mailbox configured in appsettings.json. Supports large attachments (>100MB) via Graph upload sessions. No user interaction or auth on the API itself (for this stage).

## Project
- Solution: `EmailSender.sln`
- Project: `EmailSender`
- Endpoint: `POST /api/email/send`

### Request (multipart/form-data)
Fields:
- `To` (multiple): one or more recipient emails
- `Cc` (multiple, optional)
- `Bcc` (multiple, optional)
- `Subject`: string
- `Body`: string (HTML by default)
- `IsHtml` (optional): true/false
- `Attachments` (multiple, optional): file(s)

Example using curl:
```
curl -X POST https://localhost:5001/api/email/send \
  -k \
  -F "To=user1@contoso.com" \
  -F "To=user2@contoso.com" \
  -F "Subject=Test" \
  -F "Body=<b>Hello</b>" \
  -F "IsHtml=true" \
  -F "Attachments=@C:/path/largefile.zip"
```

## Configuration
Edit `EmailSender/appsettings.json`:
```
{
  "Graph": {
    "TenantId": "<YOUR_TENANT_ID>",
    "ClientId": "<YOUR_APP_CLIENT_ID>",
    "ClientSecret": "<YOUR_APP_CLIENT_SECRET>",
    "MailboxAddress": "<SENDER_MAILBOX_ADDRESS>"
  }
}
```
- MailboxAddress is the sender mailbox (UPN or GUID) that the app is allowed to send as.

Kestrel is configured to allow unlimited request size for large attachments. Controller and form limits are disabled as well.

## Microsoft 365 App Registration and Admin Setup

Follow these steps to register an application and grant permissions for sending mail as a specific mailbox without user interaction (client credentials).

1. Register an App in Microsoft Entra admin center
   - Go to https://entra.microsoft.com > App registrations > New registration.
   - Name: `EmailSender` (any name).
   - Supported account types: Single tenant (recommended for now).
   - Redirect URI: not required for client credentials.
   - After creation, note the `Application (client) ID` and `Directory (tenant) ID`.

2. Create a client secret
   - In your app registration: Certificates & secrets > Client secrets > New client secret.
   - Description: `server-secret` and set an expiration period.
   - Copy the secret value immediately; you cannot view it later.

3. Add application permissions
   - API permissions > Add a permission > Microsoft Graph > Application permissions.
   - Add: `Mail.Send`.
   - Click "Grant admin consent" for your tenant.

4. Allow sending as the mailbox
   For application permissions, Graph can send as the mailbox specified if the app has the right mailbox policy. Typically, for most tenants, `Mail.Send` app permission is sufficient to send as any mailbox in the tenant. If restricted, ensure:
   - The mailbox exists and is not hidden from address lists for testing.
   - Optional: Use an Exchange Online Application Access Policy to scope which mailboxes the app may access (recommended in production). See: https://learn.microsoft.com/exchange/troubleshoot/administration/application-access-policy

5. Configure appsettings.json with TenantId, ClientId, ClientSecret, MailboxAddress.

6. Network considerations
   - The API streams file uploads and uses Graph chunked upload (5MB slices) for reliability with large attachments.

## Build and Run

- Trust the dev certificate (first time only):
  - Windows: `dotnet dev-certs https --trust`
- Restore/build/run:
```
dotnet build

dotnet run --project EmailSender
```
- Swagger (development only): `https://localhost:5001/openapi/v1.json`

## Notes
- This stage does not implement authentication for the Web API itself.
- All secrets are stored in appsettings.json as requested; for production, store secrets securely (Key Vault or user-secrets) and restrict permissions with access policies.
- The API creates a draft in the configured mailbox and then uploads attachments using Graph upload sessions, then sends the message.
