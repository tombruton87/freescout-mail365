# Mail 365 — FreeScout Module

Send and receive emails via the Microsoft 365 Graph API instead of SMTP/IMAP.

Replace traditional SMTP sending and IMAP fetching with direct Microsoft Graph API calls. No SMTP relay configuration, no IMAP polling — just a single Azure AD app registration.

## Screenshots

![Admin Overview](https://github.com/user-attachments/assets/63ece9f8-e1f1-454b-b1ff-7419f96425da)
*Admin overview — fetch/send status, secret expiry, retry queue, and mailbox usage at a glance*

![Connection Settings](https://github.com/user-attachments/assets/ff1dd66e-2ba5-4e2f-87d3-a2ef66e9bd4b)
*Incoming connection settings — Azure credentials, folder selection, fetch mode, and post-fetch actions*

## Features

- **Send via Graph API** — bypasses SMTP entirely, supports attachments up to 150MB via chunked upload
- **Receive via Graph API** — fetches incoming mail from any folder (Inbox, subfolders, or custom selections)
- **Shared mailbox support** — send and receive from shared mailboxes without user credentials
- **Automatic retry queue** — transient failures (429/503) are queued and retried automatically (max 5 attempts, capped at 20 messages)
- **Client secret expiry monitoring** — detects when your Azure secret is expiring and creates a FreeScout alert ticket
- **Admin overview dashboard** — single page showing fetch/send status, secret expiry, retry queue, and mailbox usage across all configured mailboxes
- **Post-fetch actions** — mark as read or move to a specific folder after fetching
- **Fetch mode** — fetch all new messages (timestamp-based) or unread only
- **Mailbox quota monitoring** — tracks mailbox storage usage via the Graph API
- **Shared mailbox browser** — browse available Microsoft 365 mailboxes from within FreeScout
- **Send and fetch logging** — last 30 send operations and last 20 fetch operations logged per mailbox

## Requirements

- FreeScout 1.8.100 or later
- PHP 7.1+ with the `curl` extension
- An Azure AD (Entra ID) app registration

## Azure AD Setup

### 1. Register an application

1. Go to the [Azure portal](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Name: anything you like (e.g. "FreeScout Mail365")
3. Supported account types: **Accounts in this organizational directory only**
4. Redirect URI: leave blank (not needed for this module)
5. Click **Register**
6. Note the **Application (client) ID** and **Directory (tenant) ID**

### 2. Create a client secret

1. Go to **Certificates & secrets** → **New client secret**
2. Set a description and expiry (recommend 24 months)
3. Copy the secret **Value** immediately — it won't be shown again

### 3. Add API permissions

Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions** and add:

| Permission | Purpose | Required |
|---|---|---|
| `Mail.Send` | Send emails via the Graph API | Yes (for sending) |
| `Mail.ReadWrite` | Fetch and manage incoming emails | Yes (for receiving) |
| `Application.Read.All` | Detect client secret expiry dates | Optional |

Click **Grant admin consent** after adding permissions.

> **Warning**
>
> `Application.Read.All` grants read access to **all** application registrations in your Azure AD tenant, not just this one. This is a broad permission. It is only used to check the expiry date of your client secret so the module can warn you before it expires.
>
> **If you are not comfortable granting this permission, skip it.** The module will work without it — you just won't get automatic secret expiry detection. You can manually enter the expiry date in the module settings instead.

### 4. Grant mailbox access

Application permissions apply to **all** mailboxes in your tenant by default. To restrict access to specific mailboxes only, use [Application Access Policies](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access):

```powershell
New-ApplicationAccessPolicy -AppId "<client-id>" `
  -PolicyScopeGroupId "<mail-enabled-security-group>" `
  -AccessRight RestrictAccess `
  -Description "Restrict FreeScout to specific mailboxes"
```

## Installation

1. Download or clone this repository into your FreeScout `Modules/` directory:

   ```bash
   cd /path/to/freescout/Modules
   git clone https://github.com/tombruton87/freescout-mail365.git Mail365
   ```

2. Publish the module's public assets:

   ```bash
   php artisan module:publish Mail365
   ```

3. Log in as an admin and activate the module under **Manage → Modules**.

If you're running FreeScout in Docker, step 2 may happen automatically on container restart depending on your image.

## Configuration

### Incoming mail (Graph API fetch)

1. Go to **Mailbox → Connection Settings → Incoming**
2. Select **Microsoft 365 API** as the incoming protocol
3. Enter your **Tenant ID**, **Client ID**, and **Client Secret**
4. Click **Test Connection** — the module will verify authentication, check permissions, and confirm the mailbox exists in Microsoft 365
5. Select which folders to fetch from (defaults to Inbox)
6. Choose a fetch mode:
   - **All new messages** — uses a timestamp cursor, fetches everything received since the last run
   - **Unread only** — fetches only unread messages, marks them as read after import
7. Optionally configure a post-fetch action (mark as read or move to folder)

### Outgoing mail (Graph API send)

1. Go to **Mailbox → Connection Settings → Outgoing**
2. Select **Microsoft 365 API** as the outgoing method
3. Azure credentials are shared with the incoming configuration — no need to enter them again

### Shared mailboxes

Enter the shared mailbox email address in the **Shared Mailbox Email** field on the incoming settings page. The Azure app must have `Mail.Send` and `Mail.ReadWrite` permissions on that mailbox.

### Client secret expiry alerts

1. Enable **Secret expiry alerts** in the incoming settings
2. Set the warning threshold (default: 30 days)
3. The module will create a FreeScout conversation in the mailbox when the secret is within the threshold

You can also manually enter the secret expiry date if `Application.Read.All` is not granted.

## Admin overview

Access via **Manage → Mail 365**. Shows all configured mailboxes with:

- Last fetch/send status and timestamps
- Client secret expiry countdown
- Retry queue size
- Mailbox storage usage

## Scheduled commands

The module registers two commands with FreeScout's scheduler automatically:

| Command | Interval | Purpose |
|---|---|---|
| `mail365:fetch-emails` | Every minute | Fetches incoming mail from Microsoft 365 |
| `mail365:retry-sends` | Every 5 minutes | Retries failed sends (transient errors only) |

These require FreeScout's standard scheduler (`php artisan schedule:run`) to be running via cron.

The fetch command supports a `--debug=1` flag for troubleshooting:

```bash
php artisan mail365:fetch-emails --debug=1
```

## Protocol / method values

This module registers:

- Incoming protocol: `3` (Microsoft 365 API)
- Outgoing method: `4` (Microsoft 365 API)

If another module on your system already uses these values, there will be a conflict. Check your other modules before installing.

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| "Authentication failed" on test | Wrong credentials or secret expired | Verify tenant ID, client ID, and secret in Azure portal |
| "Missing permissions: Mail.Send" | API permissions not granted | Add the permission in Azure and click **Grant admin consent** |
| Emails not being fetched | Scheduler not running or mailbox inactive | Verify `php artisan schedule:run` is in your crontab |
| "Failed to download MIME" | Message too large or Graph API timeout | Check mailbox quota; the module retries automatically |
| Secret expiry not detected | `Application.Read.All` not granted | Either grant the permission or enter the expiry date manually |

## License

AGPL-3.0
