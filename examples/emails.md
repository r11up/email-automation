# Email queue file — example format
#
# Each ## section is one email. The script reads this file and creates
# Outlook drafts for each one.
#
# Run:  python3 graph_queue.py list --file examples/emails.md
# Run:  python3 graph_queue.py draft --file examples/emails.md --all
#
# ─────────────────────────────────────────────────────────────────────────────

## 1. Prof Jane Doe (University of Example)

**Send:** Mon 9 Nov 2026 · **Deadline:** 1 Dec 2026
**To:** jane.doe@example.edu
**Attach:** My_CV.pdf
**Subject:** PhD enquiry — machine learning for health data

Dear Prof Doe,

I have been reading about your recent work on [topic], which closely matches my research interests in [area].

I hold a [degree] in [field] from [university], with [relevant experience/achievements]. My recent work includes [brief description].

I would be very interested in discussing the possibility of PhD study under your supervision. Would you be available for a short conversation?

Kind regards,
Your Name
your.email@example.com

---

## 2. Dr Bob Smith (Another University)

**Send:** Mon 16 Nov 2026
**To:** bob.smith@another.edu
**Subject:** PhD enquiry — graph neural networks for biomedical data

Dear Dr Smith,

[Email body here...]

Kind regards,
Your Name

---

## 3. A/Prof Carol Jones (Third University)

**To:** carol.jones@third.edu
**Attach:** My_CV.pdf; Research_Proposal.pdf
**Subject:** PhD enquiry — wearable health sensing

Dear A/Prof Jones,

[This email has no Send: date, so it will only be drafted with --all]

Kind regards,
Your Name

---
