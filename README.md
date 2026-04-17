# email-hostinger-remote-mcp

MCP server for a Hostinger-hosted mailbox (IMAP/SMTP). Drop-in companion to
`email-remote-mcp` (which uses Microsoft Graph for Microsoft 365 mailboxes).

Exposes the same tool shape as `email-remote-mcp`:

- `email_list_messages`
- `email_get_message`
- `email_send`
- `email_reply`
- `email_mark_read`
- `email_move`
- `email_search`

## Environment variables

| Var | Default | Description |
|---|---|---|
| `IMAP_HOST` | `imap.hostinger.com` | IMAP server |
| `IMAP_PORT` | `993` | IMAP port |
| `IMAP_SECURE` | `true` | TLS on (SSL) |
| `SMTP_HOST` | `smtp.hostinger.com` | SMTP server |
| `SMTP_PORT` | `465` | SMTP port |
| `SMTP_SECURE` | `true` | TLS on (SSL) |
| `EMAIL_USER` | — | Mailbox address (e.g. `admin@camocreations.co.uk`) |
| `EMAIL_PASS` | — | Mailbox password |
| `API_SECRET` | — | Bearer token required on `/mcp` |
| `FROM_NAME` | — | Optional display name for outgoing mail |
| `PORT` | `3000` | HTTP port |

## Run locally

```bash
npm install
EMAIL_USER=you@example.com EMAIL_PASS=... API_SECRET=local npm start
curl http://localhost:3000/health
```

## Deploy

Designed to run on Railway. Start command `npm start` (uses `tsx` to run the
TypeScript source directly — no build step).

## Folder aliases

Friendly names are mapped to common IMAP folder paths with fallback:

- `inbox` → `INBOX`
- `sent` / `sentitems` → `INBOX.Sent`, `Sent`, `Sent Items`, `Sent Messages`
- `drafts` → `INBOX.Drafts`, `Drafts`
- `trash` / `deleteditems` → `INBOX.Trash`, `Trash`, `Deleted`, `Deleted Items`
- `junk` / `junkemail` / `spam` → `INBOX.Junk`, `Junk`, `Spam`
- `archive` → `INBOX.Archive`, `Archive`, `Archives`

The resolver lists real folders on the server and picks the first existing
candidate.
