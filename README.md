[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/ryaker-outlook-mcp-badge.png)](https://mseep.ai/app/ryaker-outlook-mcp)

# Outlook MCP Server

A Model Context Protocol (MCP) server that connects Claude to Microsoft Outlook via the Microsoft Graph API. Manage emails, calendar, contacts, folders, rules, and categories directly from Claude.

Certified by MCPHub: https://mcphub.com/mcp-servers/ryaker/outlook-mcp

---

## Quick Start

1. **Install dependencies**: `npm install`
2. **Azure setup**: Register an app in Azure Portal (see [Azure Setup](#azure-app-registration) below)
3. **Configure**: Copy `.env.example` to `.env` and add your Azure credentials
4. **Add to Claude**: Update your Claude Desktop config (see [Configuration](#configuration))
5. **Start auth server**: `npm run auth-server`
6. **Authenticate**: Ask Claude to `authenticate` — follow the OAuth URL
7. **Done**: Use any of the tools below in natural language

---

## Tools Reference

### Email

| Tool | Description |
|---|---|
| `list-emails` | List recent emails from inbox or a folder |
| `search-emails` | Search emails with rich filters (see below) |
| `read-email` | Read the full content of an email |
| `send-email` | Compose and send a new email |
| `reply-email` | Reply or reply-all to an email |
| `forward-email` | Forward an email to new recipients |
| `create-draft` | Save a draft without sending |
| `delete-email` | Move an email to Deleted Items |
| `mark-as-read` | Mark email as read (`isRead: true`) or unread (`isRead: false`) |
| `flag-email` | Flag, unflag, or mark an email as complete |
| `move-email` | Move a single email to a folder |
| `move-emails` | Move multiple emails to a folder (comma-separated IDs) |
| `get-email-thread` | Retrieve all messages in a conversation thread |
| `list-attachments` | List all attachments of an email |
| `get-attachment` | Download a specific attachment (returns content) |

#### search-emails filters

All filters can be combined freely:

| Parameter | Type | Description |
|---|---|---|
| `query` | string | Free-text search (subject + body + sender) |
| `from` | string | Filter by sender name or address |
| `to` | string | Filter by recipient |
| `subject` | string | Filter by subject text |
| `body` | string | Filter by body text |
| `attachmentName` | string | Filter by attachment filename |
| `folder` | string | Folder to search in (default: all folders) |
| `hasAttachments` | boolean | Only emails with attachments |
| `unreadOnly` | boolean | Only unread emails |
| `category` | string | Filter by Outlook category name |
| `after` | string | Emails received after date (`2025-01-15` or `"7 days ago"`) |
| `before` | string | Emails received before date |
| `minSize` | number | Minimum email size in bytes |
| `maxSize` | number | Maximum email size in bytes |
| `sortOrder` | string | `asc` or `desc` (default: newest first) |
| `skip` | number | Skip N results (pagination) |
| `count` | number | Number of results (default: 10, max: 50) |
| `mailbox` | string | Shared mailbox address (optional) |

**Example prompts:**
- *"Search emails from godaddy after December 2025 with attachments"*
- *"Find all unread emails from my boss this week"*
- *"Show me emails larger than 5 MB in the last 3 months"*
- *"Search for emails about the Q4 report with PDF attachments"*

---

### Calendar

| Tool | Description |
|---|---|
| `list-events` | List upcoming calendar events |
| `search-events` | Search events by subject, attendee, or location |
| `create-event` | Create a new calendar event |
| `update-event` | Update an existing event (only changed fields) |
| `accept-event` | Accept a meeting invitation |
| `decline-event` | Decline a meeting invitation |
| `cancel-event` | Cancel an event (sends cancellation to attendees) |
| `delete-event` | Delete an event without notification |
| `get-schedule` | Check free/busy status for one or more people |
| `find-meeting-times` | Find slots when all attendees are available |

**Example prompts:**
- *"What meetings do I have next week?"*
- *"Find a 1-hour slot when Alice and Bob are both free on Thursday"*
- *"Check if John is free tomorrow between 10 and 12"*
- *"Update the project review on Friday — move it to 3pm and add the Berlin room"*

---

### Contacts

| Tool | Description |
|---|---|
| `search-people` | Find colleagues by name or email fragment (searches org + history) |
| `list-contacts` | List personal contacts alphabetically |
| `get-contact` | Get full details for a specific contact |
| `create-contact` | Create a new personal contact |
| `update-contact` | Update contact details (only changed fields) |
| `delete-contact` | Permanently delete a contact |

**Example prompts:**
- *"Find the email address of someone named 'Müller' in my company"*
- *"Add Max Mustermann with email max@example.com to my contacts"*

---

### Folders

| Tool | Description |
|---|---|
| `list-folders` | List mail folders (optionally with item counts and child folders) |
| `create-folder` | Create a new mail folder |
| `move-email` | Move a single email to a folder |
| `move-emails` | Move multiple emails to a folder |

---

### Categories

| Tool | Description |
|---|---|
| `list-categories` | List all Outlook categories with colors |
| `create-category` | Create a new category with a color |
| `add-category` | Add a category to an email |
| `remove-category` | Remove a category from an email |

---

### Rules

| Tool | Description |
|---|---|
| `list-rules` | List all inbox rules |
| `create-rule` | Create a new inbox rule |
| `edit-rule-sequence` | Change rule priority order |

---

### Shared Mailboxes

All email, calendar, and folder tools support a `mailbox` parameter to access shared mailboxes:

```
search-emails mailbox="team@company.com" from="client@example.com"
list-events mailbox="shared@company.com"
send-email mailbox="support@company.com" to="user@example.com" ...
```

| Tool | Description |
|---|---|
| `list-shared-mailboxes` | Discover shared mailboxes you have access to |

---

### Authentication

| Tool | Description |
|---|---|
| `authenticate` | Start the OAuth flow — returns a URL to open in your browser |
| `check-auth` | Check if currently authenticated |

---

## Mailbox Exploration — Typical Workflows

These workflows show how Claude can help you explore and organize your mailbox:

### Find and triage emails
```
"Show me all unread emails from the last 2 weeks"
"Find emails from external senders with large attachments (>2MB) this year"
"Search for everything related to 'invoice' or 'rechnung' in the last 6 months"
"Find all emails I haven't read from my manager"
```

### Organize emails
```
"Move all newsletters from company X to my Newsletter folder"
"Flag all emails about project Alpha that are still open"
"Mark all emails older than 3 months in the Promotions folder as read"
```

### Investigate a topic
```
"Get the full thread for this email and summarize the conversation"
"Find all emails about the contract with supplier Y — show subject, date, and sender"
"Search for emails with PDF attachments from legal@company.com this year"
```

### Calendar + Email together
```
"Find the email invite for next Monday's board meeting and show me who's attending"
"Check if all attendees for Thursday's review are available, and if not find an alternative slot"
"Create a meeting with everyone who replied to the project kickoff email"
```

---

## Azure App Registration

### 1. Register the App

1. Open [Azure Portal](https://portal.azure.com/) → **App registrations** → **New registration**
2. Name: e.g. `Outlook MCP Server`
3. Account types: *Accounts in any organizational directory and personal Microsoft accounts*
4. Redirect URI: `Web` → `http://localhost:3333/auth/callback`
5. Click **Register** and copy the **Application (client) ID**

### 2. Set Permissions

Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**

**Required:**
- `offline_access` — token refresh
- `User.Read` — basic profile
- `Mail.Read`, `Mail.ReadWrite`, `Mail.Send` — email access
- `Calendars.Read`, `Calendars.ReadWrite` — calendar access
- `Contacts.Read`, `Contacts.ReadWrite` — contacts access
- `MailboxSettings.Read`, `MailboxSettings.ReadWrite` — mailbox settings
- `People.Read` — colleague search

**Optional (shared mailboxes):**
- `Mail.Read.Shared`, `Mail.ReadWrite.Shared`, `Mail.Send.Shared`
- `Calendars.Read.Shared`, `Calendars.ReadWrite.Shared`

Click **Grant admin consent** after adding permissions.

### 3. Create a Client Secret

Go to **Certificates & secrets** → **New client secret**

> ⚠️ Copy the **Value** column — not the Secret ID. These look similar but only the Value works.

---

## Configuration

### Environment Variables (`.env`)

```bash
cp .env.example .env
```

```bash
MS_CLIENT_ID=your-application-client-id
MS_CLIENT_SECRET=your-client-secret-VALUE
USE_TEST_MODE=false
```

### Claude Desktop Config

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["/absolute/path/to/outlook-mcp/index.js"],
      "env": {
        "OUTLOOK_CLIENT_ID": "your-client-id",
        "OUTLOOK_CLIENT_SECRET": "your-client-secret-VALUE"
      }
    }
  }
}
```

---

## Authentication Flow

### Step 1 — Start the auth server
```bash
npm run auth-server
```
Starts a local server on port 3333 for the OAuth callback.

### Step 2 — Authenticate via Claude
Ask Claude to use the `authenticate` tool. It will return a URL like:
```
http://localhost:3333/auth?client_id=...
```
Open this URL, sign in with Microsoft, and grant permissions.
Tokens are saved to `~/.outlook-mcp-tokens.json` and refreshed automatically.

> The auth server only needs to run during initial authentication. Once tokens are saved, it can be stopped.

---

## Development

```bash
npm install          # Install dependencies
npm start            # Start the MCP server
npm run auth-server  # Start the OAuth server (port 3333)
npm run test-mode    # Start with mock data (no real API calls)
npm test             # Run Jest tests (316 tests)
npm run inspect      # Test with MCP Inspector
```

### Project Structure

```
index.js                    # Entry point — combines all modules
config.js                   # Centralized config (API, pagination, auth)
outlook-auth-server.js      # OAuth server (port 3333)

auth/                       # OAuth + token management
calendar/
  list.js                   # list-events (calendarView, recurring expansion)
  search.js                 # search-events (subject, attendee, location)
  create.js                 # create-event
  update.js                 # update-event (PATCH, only changed fields)
  accept.js                 # accept-event
  decline.js                # decline-event
  cancel.js                 # cancel-event
  delete.js                 # delete-event
  schedule.js               # get-schedule (free/busy)
  find-times.js             # find-meeting-times
email/
  list.js                   # list-emails
  search.js                 # search-emails (KQL/$filter/$orderby)
  read.js                   # read-email
  send.js                   # send-email
  draft.js                  # create-draft
  reply.js                  # reply-email, forward-email
  flag.js                   # flag-email
  thread.js                 # get-email-thread
  mark-as-read.js           # mark-as-read (also unread via isRead:false)
  delete.js                 # delete-email
  attachments.js            # list-attachments, get-attachment
folder/
  list.js                   # list-folders
  create.js                 # create-folder
  move.js                   # move-emails (bulk)
  move-single.js            # move-email (single)
contacts/
  search.js                 # search-people (/me/people)
  crud.js                   # list/get/create/update/delete contacts
categories/
  list.js, add.js, remove.js, create.js
rules/
  list.js, create.js, edit.js
utils/
  graph-api.js              # callGraphAPI, callGraphAPIPaginated
  mailbox-path.js           # Shared mailbox path helpers
  odata-helpers.js          # escapeODataString (only ' → '')
  folder-utils.js           # Folder resolution by name
```

### Search Architecture (email/search.js)

Three search paths selected automatically based on parameters:

- **Path 1** — Text terms present → `$search` with KQL
  Supports: `from:`, `to:`, `subject:`, `body:`, `attachment:`, `received:YYYY/MM/DD..`, `hasAttachments:true`, `isRead:false`
  Note: `$orderby` and `$skip` not compatible with `$search` → applied client-side

- **Path 2** — Filters only (no text) → `$filter` with `$orderby`
  Supports: date ranges, size, hasAttachments, unreadOnly, categories
  Full server-side pagination with `$skip`

- **Path 3** — No parameters → basic listing

---

## Troubleshooting

| Problem | Solution |
|---|---|
| `Cannot find module '@modelcontextprotocol/sdk'` | Run `npm install` |
| `EADDRINUSE: port 3333` | Run `npx kill-port 3333`, then restart auth server |
| `AADSTS7000215: Invalid client secret` | Use the secret **Value**, not the Secret ID |
| Auth URL doesn't work | Make sure `npm run auth-server` is running first |
| `Authentication required` error | Delete `~/.outlook-mcp-tokens.json` and re-authenticate |
| Server doesn't appear in Claude | Check absolute path in Claude Desktop config, restart Claude |
