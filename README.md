# outlook-draft-queue

A command-line tool for managing email outreach from plain markdown files via the Microsoft Graph API. Drafts are version-controlled, reviewable as diffs, and pushed to Outlook in one command.

---

## What it does

Reads email blocks from a markdown file, creates drafts in your Outlook mailbox via the Graph API, and reads incoming mail back to the terminal. Nothing is ever sent automatically — every draft is reviewed and sent by hand.

```
python3 graph_queue.py login                      # one-time device-code sign-in
python3 graph_queue.py list --file emails.md      # inspect the queue
python3 graph_queue.py draft --file emails.md --all        # push all to Drafts
python3 graph_queue.py draft --file emails.md --only 3,7   # draft specific items
python3 graph_queue.py read -n 30 --detail        # read 30 inbox emails with full body
```

---

## Requirements

- macOS (written and tested on Sonoma / Sequoia)
- Python 3.8+ — ships with macOS, no extra install needed
- A personal Microsoft account (`@outlook.com`, `@hotmail.com`, `@live.com`, etc.)
- No third-party packages — pure standard library

---

## Setup

### 1. Clone

```bash
git clone https://github.com/r11up/outlook-draft-queue.git
cd outlook-draft-queue
```

### 2. Register a free Azure app (one-time, ~3 minutes)

The Graph API requires an OAuth app to identify the caller. This is a free registration — no subscription, no billing.

1. Open **[portal.azure.com](https://portal.azure.com)** and sign in with your personal Microsoft account  
   *(Use portal.azure.com, not entra.microsoft.com — that one is for organisational accounts only)*
2. Search for **App registrations** → **New registration**
3. Fill in:
   - **Name:** anything — `outlook-draft-queue` works
   - **Supported account types:** *Accounts in any organizational directory and personal Microsoft accounts*
   - **Redirect URI:** leave blank
   - Click **Register**
4. On the **Overview** page, copy the **Application (client) ID**
5. Left menu → **Authentication** → scroll to **Advanced settings** → set **Allow public client flows** to **Yes** → **Save**
6. Paste the ID into `graph_client_id.txt`:

```bash
echo "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" > graph_client_id.txt
```

The client ID is not a secret — it is fine to commit it.

### 3. Sign in

```bash
python3 graph_queue.py login
```

The terminal prints a URL and a short code. Open the URL in any browser, enter the code, sign in. Done in about 30 seconds. The token is saved to `graph_token.json` (chmod 600) and auto-refreshes for ~90 days.

---

## Email format

Emails live in plain markdown files. Each `##` section is one email:

```markdown
## 1. Prof Jane Doe (University of Example)

**Send:** Mon 9 Nov 2026 · **Deadline:** 1 Dec 2026
**To:** jane.doe@example.edu
**Attach:** CV.pdf; Research_Proposal.pdf
**Subject:** PhD enquiry — machine learning for health data

Dear Prof Doe,

[body, exactly as you want it to appear in Outlook]

---

## 2. Dr Bob Smith (Another University)

**To:** bob.smith@another.edu
**Subject:** PhD enquiry — graph neural networks

Dear Dr Smith,

[...]
```

| Field | Required | Notes |
|---|---|---|
| `**To:**` | ✅ | Recipient address |
| `**Subject:**` | ✅ | Subject line |
| `**Send:**` | No | If present, item is only drafted when its date has passed or `--all` is set |
| `**Attach:**` | No | Filename(s), semicolon-separated. Searched in the markdown file's folder. |

The `##` number prefix (e.g. `## 3.`) becomes the ID used with `--only 3`. Use any scheme you like.

A duplicate-prevention log (`graph_queue_log.json`) records every draft created. Re-running a command never creates a second copy. To re-draft an item after editing it, remove its entry from the log.

---

## Reading mail

```bash
python3 graph_queue.py read               # latest 10 inbox emails, subject + preview
python3 graph_queue.py read -n 50         # latest 50
python3 graph_queue.py read -n 20 --detail  # full body text
```

---

## Files

| File | Purpose |
|---|---|
| `graph_queue.py` | Main script — Graph API auth, drafting, reading |
| `outlook_queue.py` | Original AppleScript version. Still works on Legacy Outlook; also provides the markdown parser used by `graph_queue.py` |
| `graph_client_id.txt` | Your Azure app ID. Not a secret. |
| `graph_token.json` | OAuth token cache. In `.gitignore`. Never commit. |
| `graph_queue_log.json` | Draft log. Prevents duplicates. Safe to commit. |
| `examples/emails.md` | Minimal example showing the markdown format |

---

## Security note

The token scope is `Mail.ReadWrite` — it can create/read drafts and read mail, but **cannot send**. Sending always requires a manual action in Outlook. `graph_token.json` is chmod 600 by the script and is in `.gitignore`.

---

MIT licence.
