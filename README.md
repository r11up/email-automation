# email-automation

This is a tool I built to manage my email outreach from the command line. I write all my emails in plain markdown files, and the tool pushes them to Outlook as drafts — one command, no copy-pasting, no manual attaching. I keep it here for further use and iteration.

Works with new Outlook for Mac (2023+). Uses the Microsoft Graph API, so drafts show up instantly everywhere — on the Mac, in the browser, on the phone. Pure Python, no dependencies.

---

## Usage

```
python3 graph_queue.py login                       # sign in once
python3 graph_queue.py list --file emails.md       # see what's queued
python3 graph_queue.py draft --file emails.md --all         # push all to Drafts
python3 graph_queue.py draft --file emails.md --only 3,7    # draft specific items
python3 graph_queue.py read -n 30 --detail         # read 30 inbox emails with full body
```

Nothing is ever sent automatically. Every draft is reviewed and sent by hand from Outlook.

---

## Requirements

- macOS (Sonoma / Sequoia)
- Python 3.8+ — already on your Mac
- A personal Microsoft account (`@outlook.com`, `@hotmail.com`, `@live.com`, etc.)
- No pip installs — pure standard library

---

## Setup

### 1. Clone

```bash
git clone https://github.com/r11up/email-automation.git
cd email-automation
```

### 2. Register a free Azure app (one-time, ~3 minutes)

The Graph API needs an OAuth app to identify the caller. Free, no subscription.

1. Open **[portal.azure.com](https://portal.azure.com)** — sign in with your personal Microsoft account  
   *(not entra.microsoft.com — that's for organisational accounts)*
2. Search **App registrations** → **New registration**
3. Name: anything. Supported account types: *Accounts in any organizational directory and personal Microsoft accounts*. Redirect URI: blank. → **Register**
4. Copy the **Application (client) ID** from the Overview page
5. Left menu → **Authentication** → **Advanced settings** → **Allow public client flows** = **Yes** → **Save**
6. Put the ID in `graph_client_id.txt`:

```bash
echo "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" > graph_client_id.txt
```

### 3. Sign in

```bash
python3 graph_queue.py login
```

Prints a URL and a short code. Open the URL, enter the code, done. Token saves to `graph_token.json` and auto-refreshes.

---

## Email format

Each `##` section in a markdown file is one email:

```markdown
## 1. Recipient Name (Organisation)

**Send:** Mon 9 Nov 2026 · **Deadline:** 1 Dec 2026
**To:** recipient@domain.com
**Attach:** file.pdf
**Subject:** Subject line

Dear ...,

[body]

---

## 2. Another Recipient

**To:** another@domain.com
**Subject:** Subject line

Dear ...,

[body]
```

| Field | Required | Notes |
|---|---|---|
| `**To:**` | ✅ | Recipient address |
| `**Subject:**` | ✅ | Subject line |
| `**Send:**` | No | Draft only when date has passed, or use `--all` |
| `**Attach:**` | No | Filename(s), semicolon-separated. Looked up in the same folder as the markdown file. |

A log file (`graph_queue_log.json`) tracks every draft created — re-running never duplicates. To re-draft an edited item, remove its entry from the log.

---

## Reading mail

```bash
python3 graph_queue.py read                  # latest 10 emails
python3 graph_queue.py read -n 50            # latest 50
python3 graph_queue.py read -n 20 --detail   # full body
```

---

## Files

| File | Purpose |
|---|---|
| `graph_queue.py` | Main script — Graph API, drafting, reading |
| `outlook_queue.py` | Legacy AppleScript version + the markdown parser imported by `graph_queue.py` |
| `graph_client_id.txt` | Azure app ID. Not a secret. |
| `graph_token.json` | Token cache. In `.gitignore`. Never commit. |
| `graph_queue_log.json` | Draft log. Prevents duplicates. |

---

MIT licence.
