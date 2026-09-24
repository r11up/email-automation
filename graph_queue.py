#!/usr/bin/env python3
"""
outlook-draft-queue — Create Outlook drafts and read inbox via Microsoft Graph API.

macOS specific. Works with new Outlook for Mac (2023+) where AppleScript no longer works.
Requires Python 3.8+, no third-party packages needed.

ONE-TIME SETUP (about 3 minutes, free):
  1. Go to https://portal.azure.com  (NOT entra.microsoft.com — that's for work accounts)
     Sign in with your personal Microsoft account (@outlook.com, @hotmail.com, @live.com etc.)
  2. Search "App registrations" → New registration
  3. Name: anything (e.g. "outlook-draft-queue")
     Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
     Redirect URI: leave blank → Register
  4. Copy the Application (client) ID from the Overview page
  5. Left menu → Authentication → Advanced settings → "Allow public client flows" = Yes → Save
  6. Put the client ID in graph_client_id.txt  (or pass --client-id on the command line)

USAGE:
  python3 graph_queue.py login                     # sign in once — opens browser, 30 seconds
  python3 graph_queue.py list                      # show queued emails in the markdown file
  python3 graph_queue.py list --file "path/to/emails.md"
  python3 graph_queue.py draft --all               # create all emails as Outlook drafts
  python3 graph_queue.py draft --only 3,7          # draft only items 3 and 7
  python3 graph_queue.py draft --file "emails.md" --all
  python3 graph_queue.py read                      # read latest 10 inbox emails
  python3 graph_queue.py read -n 30                # read 30 emails
  python3 graph_queue.py read -n 10 --detail       # read with full body text

Token is cached in graph_token.json. Delete it to sign out.
Draft log is in graph_queue_log.json. Prevents re-drafting already-queued emails.

The script NEVER sends. It only creates drafts. You review and send from Outlook.
"""

import argparse, base64, json, os, stat, sys, time, urllib.parse, urllib.request, urllib.error
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from outlook_queue import parse, find_attachment, ROOT, DEFAULT_FILE   # reuse the parser

AUTH  = "https://login.microsoftonline.com/consumers/oauth2/v2.0"
GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = "offline_access Mail.ReadWrite"
TOKENS = ROOT / "graph_token.json"
IDFILE = ROOT / "graph_client_id.txt"
LOG    = ROOT / "graph_queue_log.json"


# ─── HTTP helpers ──────────────────────────────────────────────────────────────

def post(url, data, headers=None, raw=False):
    body = data if raw else urllib.parse.urlencode(data).encode()
    req = urllib.request.Request(url, data=body,
                                 headers=headers or {"Content-Type": "application/x-www-form-urlencoded"})
    try:
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read() or b"{}")
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"{e.code} {e.read().decode()[:400]}")


def get(url, headers=None):
    req = urllib.request.Request(url, headers=headers or {})
    try:
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read() or b"{}")
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"{e.code} {e.read().decode()[:400]}")


# ─── Auth ──────────────────────────────────────────────────────────────────────

def client_id(arg=None):
    """Load client ID from arg, file, or exit with a helpful message."""
    if arg:
        return arg.strip()
    if IDFILE.exists():
        return IDFILE.read_text().strip()
    sys.exit(
        "No client ID found.\n"
        f"  Create {IDFILE} with your Azure app's Application (client) ID.\n"
        "  See README.md for the 3-minute setup guide.\n"
    )


def save_tokens(t):
    t["obtained"] = time.time()
    TOKENS.write_text(json.dumps(t))
    os.chmod(TOKENS, stat.S_IRUSR | stat.S_IWUSR)   # 600 — only you can read it


def login(cid):
    """Device-code flow: print a URL + code, poll until user authenticates."""
    d = post(f"{AUTH}/devicecode", {"client_id": cid, "scope": SCOPES})
    print("\n  " + d["message"] + "\n")
    deadline = time.time() + int(d.get("expires_in", 900))
    while time.time() < deadline:
        time.sleep(int(d.get("interval", 5)))
        try:
            t = post(f"{AUTH}/token", {
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                "client_id":  cid,
                "device_code": d["device_code"],
            })
            save_tokens(t)
            print("Signed in successfully. Token saved to graph_token.json")
            return t["access_token"]
        except RuntimeError as e:
            if "authorization_pending" in str(e):
                continue
            raise
    sys.exit("Sign-in timed out. Run login again.")


def token(cid):
    """Return a valid access token, refreshing or re-logging-in as needed."""
    if not TOKENS.exists():
        return login(cid)
    t = json.loads(TOKENS.read_text())
    # Still valid (with 5-minute buffer)?
    if time.time() - t.get("obtained", 0) < int(t.get("expires_in", 3600)) - 300:
        return t["access_token"]
    # Try refresh
    try:
        n = post(f"{AUTH}/token", {
            "grant_type":    "refresh_token",
            "client_id":     cid,
            "refresh_token": t["refresh_token"],
            "scope":         SCOPES,
        })
        save_tokens(n)
        return n["access_token"]
    except RuntimeError:
        return login(cid)


# ─── Drafting ──────────────────────────────────────────────────────────────────

def create_draft(tok, item, attachment):
    """
    POST a new draft message to the Graph API, then attach files if any.

    item:       dict with keys subject, body, to  (from outlook_queue.parse())
    attachment: a Path object, a list of Path objects, or None
    returns:    the Graph message ID of the created draft
    """
    msg = {
        "subject": item["subject"],
        "body":    {"contentType": "Text", "content": item["body"]},
        "toRecipients": [{"emailAddress": {"address": item["to"]}}],
    }
    h = {"Authorization": "Bearer " + tok, "Content-Type": "application/json"}
    made = post(f"{GRAPH}/me/messages", json.dumps(msg).encode(), h, raw=True)

    # Attach files
    for one in (attachment if isinstance(attachment, list) else [attachment]):
        if not one:
            continue
        att = {
            "@odata.type":  "#microsoft.graph.fileAttachment",
            "name":          one.name,
            "contentBytes":  base64.b64encode(one.read_bytes()).decode(),
        }
        post(f"{GRAPH}/me/messages/{made['id']}/attachments",
             json.dumps(att).encode(), h, raw=True)
    return made["id"]


# ─── Reading ───────────────────────────────────────────────────────────────────

def read_messages(tok, folder="inbox", limit=10, unread_only=False):
    """Fetch messages from a mailbox folder. Returns a list of message dicts."""
    params = [
        f"$top={limit}",
        "$select=id,receivedDateTime,from,toRecipients,subject,bodyPreview,body,isRead,isDraft",
        "$orderby=receivedDateTime+desc",
    ]
    # Note: $filter and $orderby together can cause errors with some accounts.
    # If you see "Cannot use the 'filter' and 'orderby' options together", remove the orderby.
    if unread_only:
        # Use a separate params list without orderby to avoid the $filter + $orderby conflict
        params = [
            f"$top={limit}",
            "$select=id,receivedDateTime,from,toRecipients,subject,bodyPreview,body,isRead,isDraft",
            "$filter=isRead+eq+false",
        ]
    query = "&".join(params)
    url = f"{GRAPH}/me/mailFolders/{folder}/messages?{query}"
    h = {"Authorization": "Bearer " + tok, "Content-Type": "application/json"}
    res = get(url, h)
    return res.get("value", [])


def print_messages(messages, detail=False):
    """Pretty-print a list of message dicts to stdout."""
    if not messages:
        print("No messages found.")
        return
    for i, m in enumerate(messages, 1):
        sender    = m.get("from", {}).get("emailAddress", {})
        s_name    = sender.get("name", "")
        s_addr    = sender.get("address", "")
        sender_str = f"{s_name} <{s_addr}>" if s_name else s_addr
        recv      = m.get("receivedDateTime", "")[:19].replace("T", " ")
        status    = " [UNREAD]" if not m.get("isRead") else ""

        print(f"\n[{i}] {m.get('subject', '(No Subject)')}{status}")
        print(f"    From: {sender_str}")
        print(f"    Date: {recv} UTC")

        if detail:
            import re
            content = m.get("body", {}).get("content", "")
            clean   = re.sub(r"<[^<]+?>", "", content).strip()
            lines   = [line.strip() for line in clean.splitlines() if line.strip()]
            print("    Body:")
            for line in lines[:30]:
                print(f"      {line}")
            if len(lines) > 30:
                print(f"      ... ({len(lines) - 30} more lines)")
        else:
            preview = m.get("bodyPreview", "").strip()
            if preview:
                print(f"    Preview: {preview[:160]}...")


# ─── CLI ───────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(
        description="Create Outlook drafts and read inbox via Microsoft Graph API.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python3 graph_queue.py login
  python3 graph_queue.py list --file emails/phd_outreach.md
  python3 graph_queue.py draft --file emails/phd_outreach.md --all
  python3 graph_queue.py draft --file emails/phd_outreach.md --only 1,2,3
  python3 graph_queue.py read -n 20 --detail
        """,
    )
    ap.add_argument("action", choices=["login", "list", "draft", "read"],
                    help="Action to perform")
    ap.add_argument("--file",      default=str(DEFAULT_FILE),
                    help="Markdown file containing email queue (default: examples/emails.md)")
    ap.add_argument("--only",      help="Comma-separated item IDs to draft, e.g. 1,3,7")
    ap.add_argument("--all",       action="store_true",
                    help="Draft all unsent items, ignoring Send: dates")
    ap.add_argument("--client-id", help="Azure app client ID (overrides graph_client_id.txt)")
    ap.add_argument("-n", "--limit", type=int, default=10,
                    help="Number of emails to read (default: 10)")
    ap.add_argument("--folder",    default="inbox",
                    help="Mailbox folder to read from (default: inbox)")
    ap.add_argument("--unread",    action="store_true",
                    help="Show only unread emails")
    ap.add_argument("--detail",    action="store_true",
                    help="Show full message body (default: preview only)")
    a = ap.parse_args()

    cid = client_id(a.client_id)

    # ── login ──────────────────────────────────────────────────────────────────
    if a.action == "login":
        login(cid)
        return

    # ── read ───────────────────────────────────────────────────────────────────
    if a.action == "read":
        tok  = token(cid)
        msgs = read_messages(tok, folder=a.folder, limit=a.limit, unread_only=a.unread)
        print_messages(msgs, detail=a.detail)
        return

    # ── list / draft ───────────────────────────────────────────────────────────
    path = Path(a.file)
    if not path.is_absolute():
        path = ROOT / path
    if not path.exists():
        sys.exit(f"File not found: {path}\nCreate it following the format in examples/emails.md")

    items = parse(path)
    log   = json.loads(LOG.read_text()) if LOG.exists() else {}
    only  = {s.strip() for s in a.only.split(",")} if a.only else None
    key   = lambda it: f"{path.name}#{it['id']}"

    # Items that haven't been drafted yet, and pass the filter
    picked = [it for it in items
              if key(it) not in log and (only is None or it["id"] in only)]
    # Without --all, only draft items where Send: date has passed
    if only is None and not a.all:
        picked = [it for it in picked if it["send_date"]]

    if a.action == "list":
        print(f"\n{path.name}: {len(items)} email(s), {len(log)} already drafted\n")
        for it in items:
            state = "drafted" if key(it) in log else ("DUE" if it in picked else "")
            date  = it["send_date"].strftime("%d %b") if it["send_date"] else "no date"
            print(f"  #{it['id']:<4} {date:<8} {state:<8} {it['head'][:44]:<46} {it['to']}")
        return

    if not picked:
        print("Nothing to draft.")
        return

    tok = token(cid)
    print(f"Drafting {len(picked)} email(s)...\n")
    for it in picked:
        names   = [n.strip() for n in (it["attach"] or "").split(";") if n.strip()]
        att     = [find_attachment(n) for n in names]
        missing = [n for n, a in zip(names, att) if not a]
        if missing:
            print(f"  #{it['id']} SKIPPED — attachment not found: {', '.join(missing)}")
            continue
        try:
            mid = create_draft(tok, it, att)
            log[key(it)] = {
                "id":      mid,
                "to":      it["to"],
                "subject": it["subject"],
                "at":      time.strftime("%Y-%m-%d %H:%M"),
            }
            LOG.write_text(json.dumps(log, indent=1, sort_keys=True))
            print(f"  #{it['id']} drafted → {it['to']}  ({it['subject'][:50]})")
        except RuntimeError as e:
            print(f"  #{it['id']} FAILED: {e}")

    print(f"\nDone. Open Outlook to review and send the drafts.")


if __name__ == "__main__":
    main()
