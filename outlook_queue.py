#!/usr/bin/env python3
"""
outlook_queue.py — Legacy AppleScript driver for Outlook for Mac.

⚠️  WARNING: This script only works with LEGACY Outlook for Mac.
    New Outlook for Mac (2023+) has a broken AppleScript bridge that talks to
    an old disconnected database. Drafts created this way never appear in your
    real mailbox.

    Use graph_queue.py instead — it works with new Outlook for Mac.

This file is kept because:
  - graph_queue.py imports the parse() and find_attachment() functions from here
  - Some people may still run Legacy Outlook (it can be switched on in Outlook → Help → "Revert to Legacy Outlook")

MARKDOWN FORMAT (same format used by graph_queue.py):

    ## 1. Prof Jane Doe (University of Example)

    **Send:** Mon 9 Nov 2026 · **Deadline:** 1 Dec 2026
    **To:** jane.doe@example.edu
    **Attach:** My_CV.pdf
    **Subject:** PhD enquiry — machine learning for health data

    Dear Prof Doe,

    I have been reading about your work on...

USAGE (legacy only):
  python3 outlook_queue.py list                       # what is due today
  python3 outlook_queue.py draft                      # create today's emails as Outlook drafts
  python3 outlook_queue.py draft --all                # draft everything
  python3 outlook_queue.py draft --only 12,29         # draft items 12 and 29
  python3 outlook_queue.py send                       # SENDS today's emails (asks for confirmation)
  python3 outlook_queue.py send --only 29 --yes       # send without asking

  --file  path to a markdown file (default: examples/emails.md)
  --date  pretend today is this date, e.g. 2026-11-30

Draft/send log is in outlook_queue_log.json. Nothing runs twice.
"""

import argparse, datetime, json, re, subprocess, sys, tempfile
from pathlib import Path

# ─── Path configuration ────────────────────────────────────────────────────────
ROOT         = Path(__file__).resolve().parent
DEFAULT_FILE = ROOT / "emails.md"
LOG          = ROOT / "outlook_queue_log.json"

# Directories to search for attachment files (in order).
# Add your own directories here if your files are elsewhere.
CV_DIRS = [
    ROOT,   # same folder as this script
]


# ─── Markdown parser ───────────────────────────────────────────────────────────
BLOCK = re.compile(r"(?m)^## (?P<head>.+)$")
FIELD = lambda k: re.compile(r"(?m)^\*\*%s:\*\*[ ]*(?P<v>.+)$" % k)


def parse(path):
    """
    Parse a markdown file and return a list of email dicts.

    Each dict has keys:
        id          str    the number or code before the period (e.g. "1", "A3")
        head        str    the full ## heading
        to          str    recipient email address
        subject     str    email subject line
        attach      str|None  semicolon-separated attachment filenames, or None
        send_date   date|None  parsed send date, or None
        body        str    plain-text email body
    """
    text  = path.read_text()
    heads = list(BLOCK.finditer(text))
    out   = []
    for i, h in enumerate(heads):
        body = text[h.end(): heads[i + 1].start() if i + 1 < len(heads) else len(text)]
        body = re.split(r"(?m)^# ", body)[0]         # stop at any H1

        to   = FIELD("To").search(body)
        subj = FIELD("Subject").search(body)
        if not (to and subj):
            continue                                   # not an email block, skip

        att    = FIELD("Attach").search(body)
        sendm  = re.search(r"(?m)^\*\*Send:\*\*[ ]*(.+?)(?:[ ]*·|$)", body)
        msg    = body[subj.end():].strip()
        msg    = re.sub(r"(?m)^---\s*$", "", msg).strip()
        num    = re.match(r"\s*([A-E]?\d+)\.", h.group("head"))

        out.append({
            "id":        num.group(1) if num else h.group("head")[:12],
            "head":      h.group("head").strip(),
            "to":        to.group("v").strip(),
            "subject":   subj.group("v").strip(),
            "attach":    att.group("v").strip() if att else None,
            "send_date": parse_date(sendm.group(1).strip()) if sendm else None,
            "body":      msg,
        })
    return out


def parse_date(s):
    """Parse 'Mon 9 Nov 2026' or '9 Nov 2026' → datetime.date. Returns None on failure."""
    s = re.sub(r"^[A-Za-z]{3}\s+", "", s.strip())   # strip weekday
    for fmt in ("%d %b %Y", "%d %B %Y"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def find_attachment(name):
    """
    Search CV_DIRS for a file with the given name.
    Returns a Path if found, or None.
    """
    if not name:
        return None
    for d in CV_DIRS:
        p = d / name
        if p.exists():
            return p
    # Broad glob fallback
    hits = [p for d in CV_DIRS for p in d.glob("*.pdf") if p.name == name]
    return hits[0] if hits else None


# ─── Log helpers ───────────────────────────────────────────────────────────────

def load_log():
    return json.loads(LOG.read_text()) if LOG.exists() else {}


def save_log(log):
    LOG.write_text(json.dumps(log, indent=1, sort_keys=True))


def key(path, item):
    return f"{path.name}#{item['id']}"


# ─── AppleScript (legacy only) ─────────────────────────────────────────────────

def applescript(item, attachment, send):
    """
    Build an AppleScript string that creates an Outlook draft or sends it.
    Body is written to a temp file to avoid quoting issues with special characters.
    """
    tmp = Path(tempfile.mkstemp(suffix=".txt")[1])
    tmp.write_text(item["body"])
    esc = lambda s: s.replace("\\", "\\\\").replace('"', '\\"')

    lines = [
        'tell application "Microsoft Outlook"',
        f'  set bodyText to (read POSIX file "{tmp}" as «class utf8»)',
        f'  set msg to make new outgoing message with properties '
        f'{{subject:"{esc(item["subject"])}", content:bodyText}}',
        f'  tell msg to make new to recipient with properties '
        f'{{email address:{{address:"{esc(item["to"])}"}}}}',
    ]
    if attachment:
        lines.append(
            f'  tell msg to make new attachment with properties '
            f'{{file:POSIX file "{esc(str(attachment))}"}}',
        )
    lines.append("  send msg" if send else "  -- left in Drafts")
    lines += ["end tell"]
    return "\n".join(lines), tmp


def run(item, attachment, send):
    """Run the AppleScript. Returns (success: bool, error_message: str)."""
    script, tmp = applescript(item, attachment, send)
    try:
        r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
        if r.returncode != 0:
            return False, r.stderr.strip()[:300]
        return True, ""
    finally:
        tmp.unlink(missing_ok=True)


# ─── CLI ───────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(
        description="⚠️  Legacy AppleScript driver — use graph_queue.py for new Outlook."
    )
    ap.add_argument("action", choices=["list", "draft", "send"])
    ap.add_argument("--file", default=str(DEFAULT_FILE))
    ap.add_argument("--date", help="Pretend today is this date (YYYY-MM-DD)")
    ap.add_argument("--only", help="Comma-separated item IDs, e.g. 1,3,7")
    ap.add_argument("--all",  action="store_true", help="Process every item, ignoring dates")
    ap.add_argument("--yes",  action="store_true", help="Skip confirmation prompt for send")
    a = ap.parse_args()

    path  = Path(a.file)
    if not path.is_absolute():
        path = ROOT / path
    today = datetime.date.fromisoformat(a.date) if a.date else datetime.date.today()
    items = parse(path)
    log   = load_log()
    only  = {s.strip() for s in a.only.split(",")} if a.only else None

    def due(it):
        if key(path, it) in log:
            return False
        if only:
            return it["id"] in only
        if a.all:
            return True
        return it["send_date"] is not None and it["send_date"] <= today

    picked = [it for it in items if due(it)]

    if a.action == "list":
        print(f"\n{path.name}: {len(items)} emails, {len(log)} already handled\n")
        for it in items:
            state = log.get(key(path, it), {}).get("action", "")
            when  = it["send_date"].strftime("%a %d %b %Y") if it["send_date"] else "no date"
            mark  = "DUE" if it in picked else ("done:" + state if state else "")
            print(f"  #{it['id']:<4} {when:<16} {mark:<10} {it['head'][:44]:<46} {it['to']}")
        return

    if not picked:
        print("Nothing due."); return

    print(f"{a.action}: {len(picked)} email(s)")
    for it in picked:
        att  = find_attachment(it["attach"])
        flag = "" if att or not it["attach"] else "  !! attachment not found: " + it["attach"]
        print(f"  #{it['id']:<4} {it['to']:<38} {it['subject'][:46]}{flag}")

    if a.action == "send" and not a.yes:
        if input("\nSend these now? Type yes: ").strip().lower() != "yes":
            print("Stopped."); return

    for it in picked:
        att = find_attachment(it["attach"])
        if it["attach"] and not att:
            print(f"  #{it['id']} skipped, attachment missing"); continue
        ok, err = run(it, att, send=(a.action == "send"))
        print(f"  #{it['id']} {'ok' if ok else 'FAILED ' + err}")
        if ok:
            log[key(path, it)] = {
                "action":  a.action,
                "at":      datetime.datetime.now().isoformat(timespec="seconds"),
                "to":      it["to"],
                "subject": it["subject"],
            }
            save_log(log)


if __name__ == "__main__":
    main()
