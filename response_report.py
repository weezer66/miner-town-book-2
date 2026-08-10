#!/usr/bin/env python3
"""
response_report.py — who replied to the MINER TOWN outreach, and who didn't.

Read-only. Sends nothing. Scans your Gmail Sent folder for everyone you emailed
with the MINER TOWN subject, then checks your Inbox to see who wrote back.
Writes two CSVs you can open anytime:

    responded.csv    email, emailed_on, latest_reply_on, opted_out, has_book,
                     needs_help, conversation
    no_response.csv  email, emailed_on, conversation

The 'conversation' column is the FULL two-way thread — every message you sent
(original + manual replies + follow-ups, read from your Sent folder) and every
reply they sent (from your Inbox), timestamped and in order.
    has_book.txt     emails of readers who signaled they already have the book
                     (clicked the 'I got it on BookSprout' button or said so) —
                     execute_followups.py skips these, so they get no follow-ups.
    needs_help.txt   emails of readers who signaled they're stuck getting it
                     (clicked 'I'm having trouble' or said so) — flagged in the
                     sheet so you can reach out and help. NOT skipped.

latest_reply_on is the timestamp of their most recent reply. The "comments"
column holds a dated, one-per-line history of EVERY reply they sent (quoted
history and signatures stripped), e.g.:
    06/10/2026 - congratulations!!
    06/14/2026 - read half the book, loving it
so you can spot folks who said they'd read it and are worth a personal nudge.

  export SMTP_USER="you@gmail.com"
  export SMTP_PASS="your-16-char-app-password"
  python response_report.py

Standard library only.
"""
import csv
import email
import html
import imaplib
import os
import re
import sys
from datetime import datetime
from email.utils import getaddresses, parsedate_to_datetime

IMAP_HOST = "imap.gmail.com"
IMAP_TIMEOUT = 120
SENT_FOLDER = os.environ.get("SENT_FOLDER", "[Gmail]/Sent Mail")
SUBJECT = os.environ.get("ORIG_SUBJECT", "MINER TOWN: AWAKENING")
RESPONDED_CSV = os.environ.get("RESPONDED_CSV", "responded.csv")
NO_RESPONSE_CSV = os.environ.get("NO_RESPONSE_CSV", "no_response.csv")

OPT_OUT_WORDS = ("unsubscribe", "no thanks", "no thank you", "stop", "remove me",
                 "opt out", "opt-out", "not interested", "please stop")

# The reply buttons open threads with their OWN subject ("I got the book on
# BookSprout" / "Trouble with BookSprout"), not the MINER TOWN subject. Scanning
# Sent for this term catches YOUR replies inside those button threads.
BUTTON_SUBJECT_MATCH = "BookSprout"

# Signals that the reader already HAS the book — so we should stop following up.
# The strongest signal is the subject of the "I got it on BookSprout" button.
GOT_IT_SUBJECT_MARK = "got the book on booksprout"
GOT_IT_WORDS = ("got it", "got the book", "got my copy", "got the copy",
                "have my copy", "have the book", "downloaded", "i download",
                "preordered", "pre-ordered", "pre ordered", "purchased",
                "bought it", "bought the book", "claimed my", "claimed the",
                "i claimed", "already have it")

HAS_BOOK_FILE = os.environ.get("HAS_BOOK_FILE", "has_book.txt")

# Signals that the reader is stuck and needs help getting the book. The strongest
# signal is the subject of the "I'm having trouble" button.
TROUBLE_SUBJECT_MARK = "trouble with booksprout"
NEEDS_HELP_WORDS = ("having trouble", "trouble downloading", "trouble accessing",
                    "can't download", "cannot download", "couldn't download",
                    "won't download", "didn't work", "did not work", "doesn't work",
                    "not working", "can't log in", "cannot log in", "couldn't log in",
                    "can't login", "unable to download", "unable to access",
                    "link doesn't work", "link not working", "having issues")

NEEDS_HELP_FILE = os.environ.get("NEEDS_HELP_FILE", "needs_help.txt")


def imap_arg(s):
    """Quote an IMAP folder/search argument (imaplib won't quote spaces)."""
    return '"' + str(s).replace('\\', '\\\\').replace('"', '\\"') + '"'


def datestr(dt):
    return dt.strftime("%Y-%m-%d") if dt else ""


def stampstr(dt):
    return dt.strftime("%Y-%m-%d %H:%M") if dt else ""


# Lines at/after these markers are quoted history or a signature — drop them.
_CUT_MARKERS = re.compile(
    r"^\s*(on .+wrote:|from:\s|sent from my |-{2,}\s*original message|"
    r"_{5,}|get outlook for|^\s*>)", re.IGNORECASE)


# Inline attribution that sits on the SAME line as the reply, e.g.
# "Thanks! On Wed, Jun 10, 2026 at 8:20 PM Ankit Saxena <...> wrote:"
_INLINE_ATTR = re.compile(r"\bon\s+.{0,120}?\bwrote:.*$", re.IGNORECASE)
# Boilerplate footers worth dropping from the snippet.
_FOOTERS = re.compile(
    r"(yahoo mail: search.*$|sent from my .*$|get outlook for .*$|"
    r"reacted via gmail.*$)", re.IGNORECASE)


def clean_reply(text, limit=300):
    """Pull just the person's own words out of a reply body: drop quoted
    history, signatures, and blank noise, then collapse to a short snippet."""
    kept = []
    for line in (text or "").splitlines():
        if _CUT_MARKERS.match(line):
            break                     # everything below is quoted/sig — stop
        if line.strip().startswith(">"):
            continue
        kept.append(line.strip())
    snippet = " ".join(w for w in " ".join(kept).split() if w)
    snippet = _INLINE_ATTR.sub("", snippet)   # cut inline "On ... wrote:"
    snippet = _FOOTERS.sub("", snippet).strip()
    if len(snippet) > limit:
        snippet = snippet[:limit].rsplit(" ", 1)[0] + "..."
    return snippet


def _strip_html(h):
    """Turn an HTML email body into readable plain text (for HTML-only replies
    that carry no text/plain part — common from mobile / some webmail)."""
    h = re.sub(r"(?is)<(script|style).*?</\1>", " ", h)
    h = re.sub(r"(?is)<br\s*/?>", "\n", h)
    h = re.sub(r"(?is)</(p|div|li|tr|h[1-6])>", "\n", h)
    h = re.sub(r"(?is)<[^>]+>", " ", h)          # drop remaining tags
    return html.unescape(h)


def extract_text(msg):
    """Return the reader's message text. Prefers text/plain; falls back to the
    text/html part with tags stripped so HTML-only replies aren't lost."""
    plain, htmltext = "", ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_disposition() == "attachment":
                continue
            ct = part.get_content_type()
            if ct == "text/plain" and not plain:
                try:
                    plain = part.get_payload(decode=True).decode(errors="ignore")
                except Exception:
                    pass
            elif ct == "text/html" and not htmltext:
                try:
                    htmltext = part.get_payload(decode=True).decode(errors="ignore")
                except Exception:
                    pass
    else:
        try:
            payload = msg.get_payload(decode=True).decode(errors="ignore")
        except Exception:
            payload = str(msg.get_payload())
        if msg.get_content_type() == "text/html":
            htmltext = payload
        else:
            plain = payload
    if plain.strip():
        return plain
    if htmltext.strip():
        return _strip_html(htmltext)
    return ""


def hdr_datetime(raw):
    """Parse an email Date header into a naive datetime in LOCAL time.

    Emails carry a timezone offset (you send from -0500/-0700, readers from
    -0400 etc.). Stripping tzinfo WITHOUT converting makes messages from
    different zones sort incorrectly — a reply can appear before the message it
    answers. Convert to local time first, then drop tzinfo.
    """
    if not raw:
        return None
    try:
        d = parsedate_to_datetime(raw)
    except Exception:
        return None
    if d.tzinfo is not None:
        d = d.astimezone()          # -> this machine's local timezone
    return d.replace(tzinfo=None)


def scan_sent(M, sender):
    """Scan the Sent folder and capture EVERY message you sent to a recipient —
    the original outreach, manual replies, and follow-ups. Two passes:
      1. subject 'MINER TOWN: AWAKENING'  -> establishes recipients + emailed_on
      2. subject 'BookSprout'             -> your replies in the button threads
         ("I got the book on BookSprout" / "Trouble with BookSprout"), which
         carry a DIFFERENT subject and were previously missed. Pass 2 only
         enriches recipients already found in pass 1.
    Returns:
        recips:   {email: earliest_send_dt}              (for emailed_on)
        outbound: {email: [(dt, snippet, is_original)]}  (all your sends)
    """
    recips, outbound = {}, {}
    seen = set()
    typ, _ = M.select(imap_arg(SENT_FOLDER), readonly=True)
    if typ != "OK":
        sys.exit(f"Could not open Sent folder '{SENT_FOLDER}'.")

    def collect(search_term, primary):
        typ, data = M.search(None, 'SUBJECT', imap_arg(search_term))
        ids = data[0].split() if (typ == "OK" and data and data[0]) else []
        label = "main thread" if primary else "button threads"
        print(f"Sent {label}: {len(ids)} message(s) (reading)...", flush=True)
        n_body = 0
        for j, mid in enumerate(ids, 1):
            if mid in seen:
                continue
            seen.add(mid)
            typ, md = M.fetch(mid, '(BODY.PEEK[HEADER.FIELDS (TO CC BCC DATE SUBJECT)])')
            if typ != "OK" or not md or not md[0]:
                continue
            hdr = email.message_from_bytes(md[0][1])
            dt = hdr_datetime(hdr.get("Date"))
            subj = (hdr.get("Subject") or "").strip()
            is_original = primary and subj.lower() == SUBJECT.strip().lower()
            addrs = []
            for h in ("To", "Cc", "Bcc"):
                addrs += [a for _, a in getaddresses(hdr.get_all(h, []))]
            addrs = [a.strip().lower() for a in addrs
                     if a and "@" in a and a.strip().lower() != sender.lower()]
            # Pass 2 must not invent new recipients — only enrich known ones.
            targets = addrs if primary else [a for a in addrs if a in recips]
            if not targets:
                continue
            snip = ""
            if not is_original:                  # fetch body only for replies
                n_body += 1
                # First ~6 KB is enough for the top of YOUR reply, without
                # downloading the whole (possibly huge) quoted thread.
                typ2, md2 = M.fetch(mid, "(BODY.PEEK[]<0.6000>)")
                if typ2 == "OK" and md2 and md2[0]:
                    try:
                        snip = clean_reply(extract_text(email.message_from_bytes(md2[0][1])))
                    except Exception:
                        snip = ""
            if j % 50 == 0:
                print(f"  ...{label} {j}/{len(ids)} (bodies fetched: {n_body})", flush=True)
            for a in targets:
                if primary and (a not in recips or (dt and (recips[a] is None or dt < recips[a]))):
                    recips[a] = dt
                outbound.setdefault(a, []).append((dt, snip, is_original))

    collect(SUBJECT, primary=True)                 # the MINER TOWN thread
    collect(BUTTON_SUBJECT_MATCH, primary=False)   # the BookSprout button threads
    return recips, outbound


def fmt_conversation(outbound, inbound):
    """Merge your sends (outbound) and their replies (inbound) into one
    timestamped, chronological two-way log."""
    events = []   # (sort_dt, line)
    for dt, snip, is_original in outbound:
        body = "original outreach" if is_original else (snip or "(no plain-text content)")
        ts = stampstr(dt) if dt else "????"
        events.append((dt or datetime.max, f"{ts}  →  SENT: {body}"))
    for dt, snip in inbound:
        ts = stampstr(dt) if dt else "????"
        events.append((dt or datetime.max, f"{ts}  ←  REPLY: {snip or '(no plain-text content)'}"))
    events.sort(key=lambda e: e[0])
    return "\n".join(line for _, line in events)


def reply_status(M, addr, since_dt):
    """Return (replied, latest_reply_on, opted_out, has_book, needs_help, inbound).

    has_book   -> any reply signals they already have the book.
    needs_help -> any reply signals they're stuck getting it.
    inbound    -> [(datetime_or_None, snippet), ...] of every reply they sent.
    """
    M.select("INBOX", readonly=True)
    crit = ['FROM', addr]
    if since_dt:
        crit = ['SINCE', since_dt.strftime("%d-%b-%Y"), 'FROM', addr]
    typ, data = M.search(None, *crit)
    if typ != "OK" or not data or not data[0]:
        return False, None, False, False, False, []

    entries = []     # (datetime_or_None, snippet)
    opted = False
    has_book = False
    needs_help = False
    for mid in data[0].split():
        # Inbound replies are small (no big attachments), so a full fetch is
        # fast and reliable — only YOUR sent messages carry the book files.
        typ, md = M.fetch(mid, "(RFC822)")
        if typ != "OK" or not md or not md[0]:
            continue
        try:
            msg = email.message_from_bytes(md[0][1])
            dt = hdr_datetime(msg.get("Date"))   # timezone-safe (local time)
            subject = (msg.get("Subject") or "").lower()
            snippet = clean_reply(extract_text(msg))
            low = snippet.lower()
            # Opt-out / got-it / needs-help judged on the person's OWN words +
            # the subject, not the quoted original beneath their reply.
            if any(w in low for w in OPT_OUT_WORDS):
                opted = True
            if GOT_IT_SUBJECT_MARK in subject or any(w in low for w in GOT_IT_WORDS):
                has_book = True
            if TROUBLE_SUBJECT_MARK in subject or any(w in low for w in NEEDS_HELP_WORDS):
                needs_help = True
            entries.append((dt, snippet))
        except Exception:
            continue

    entries.sort(key=lambda e: e[0] or datetime.max)   # oldest first
    latest = max((e[0] for e in entries if e[0]), default=None)
    return True, latest, opted, has_book, needs_help, entries


def main():
    user = os.environ.get("SMTP_USER")
    pw = os.environ.get("SMTP_PASS")
    if not user or not pw:
        sys.exit("Set SMTP_USER and SMTP_PASS (Gmail App Password) first.")

    # Without a timeout a half-open socket blocks forever: a scan once hung 2h43m.
    M = imaplib.IMAP4_SSL(IMAP_HOST, timeout=IMAP_TIMEOUT)
    M.login(user, pw)

    recips, outbound = scan_sent(M, user)
    print(f"Found {len(recips)} recipient(s) for subject '{SUBJECT}'. Checking replies...", flush=True)

    responded, no_response = [], []
    has_book_emails, needs_help_emails = [], []
    for i, (addr, sent_dt) in enumerate(sorted(recips.items()), 1):
        replied, replied_on, opted, has_book, needs_help, inbound = reply_status(M, addr, sent_dt)
        conversation = fmt_conversation(outbound.get(addr, []), inbound)
        if replied:
            responded.append((addr, datestr(sent_dt), stampstr(replied_on),
                              "yes" if opted else "no", "yes" if has_book else "no",
                              "yes" if needs_help else "no", conversation))
            if has_book:
                has_book_emails.append(addr)
            if needs_help:
                needs_help_emails.append(addr)
        else:
            no_response.append((addr, datestr(sent_dt), conversation))
        if i % 25 == 0:
            print(f"  ...checked replies {i}/{len(recips)}", flush=True)

    M.logout()

    with open(RESPONDED_CSV, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["email", "emailed_on", "latest_reply_on", "opted_out",
                    "has_book", "needs_help", "conversation"])
        w.writerows(sorted(responded))
    with open(NO_RESPONSE_CSV, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["email", "emailed_on", "conversation"])
        w.writerows(sorted(no_response))
    with open(HAS_BOOK_FILE, "w", encoding="utf-8") as f:
        f.write("# Readers who signaled they already have the book — auto-skipped "
                "from follow-ups. One email per line.\n")
        for a in sorted(set(has_book_emails)):
            f.write(a + "\n")
    with open(NEEDS_HELP_FILE, "w", encoding="utf-8") as f:
        f.write("# Readers who signaled they're stuck getting the book — flagged "
                "'Needs Help' so you can assist. One email per line.\n")
        for a in sorted(set(needs_help_emails)):
            f.write(a + "\n")

    print("\n--- report ---")
    print(f"Total emailed : {len(recips)}")
    print(f"Responded     : {len(responded)}  -> {RESPONDED_CSV}")
    print(f"   of which opted out: {sum(1 for r in responded if r[3] == 'yes')}")
    print(f"   of which have the book: {len(set(has_book_emails))}  -> {HAS_BOOK_FILE}")
    print(f"   of which need help: {len(set(needs_help_emails))}  -> {NEEDS_HELP_FILE}")
    print(f"No response   : {len(no_response)}  -> {NO_RESPONSE_CSV}")


if __name__ == "__main__":
    main()
