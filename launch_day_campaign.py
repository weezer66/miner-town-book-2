#!/usr/bin/env python3
"""
launch_day_campaign.py - send launch-day emails to your existing reviewer database
in safe daily chunks.

What it does:
- Reads recipients from followups.xlsx (the Follow-ups sheet).
- Uses a template body (default templates/launch-reminder.txt/.html).
- Skips suppressed addresses and optionally rows marked Never Follow-up?=yes.
- Tracks every launch send in launch_day_sent_log.txt.
- Enforces a daily cap so you can continue across multiple days without
  duplicates or Gmail-limit spikes.

Default mode is DRY RUN. Nothing sends unless --send is provided.
"""

import argparse
import hashlib
import mimetypes
import os
import random
import smtplib
import ssl
import sys
import time
from datetime import datetime, timedelta
from email.message import EmailMessage

from openpyxl import Workbook, load_workbook

from followup_campaign import (
    BOOKSPROUT_LINK,
    ORIGINAL_SUBJECT,
    SENT_FOLDER,
    SMTP_HOST,
    SMTP_PORT,
    SMTP_TIMEOUT,
    build_reply,
    find_original,
    imap_connect,
    load_set,
)
from send_guardrail import enforce_send_run_cooldown, mark_send_run_finished

HERE = os.path.dirname(os.path.abspath(__file__))
SHEET = os.path.join(HERE, "followups.xlsx")
DEFAULT_TEXT = os.path.join(HERE, "templates", "launch-reminder.txt")
DEFAULT_HTML = os.path.join(HERE, "templates", "launch-reminder.html")
SUPPRESSION = os.path.join(HERE, "suppression.txt")
LOG_PATH = os.path.join(HERE, "launch_day_sent_log.txt")
LOG_XLSX_PATH = os.path.join(HERE, "launch_day_sent_log.xlsx")
SEND_COOLDOWN_LOCK = os.path.join(HERE, "send_cooldown_lock.json")
DEFAULT_COVER_IMAGE = r"C:\Users\Prisha\OneDrive\Documents\Miner Town\Archive\Miner_Town_Cover Page.png"
DEFAULT_COVER_GIF = r"C:\Users\Prisha\OneDrive\Documents\Miner Town\Archive\miner_town_cover_loop.gif"

# Hard safety rails to reduce Gmail lockouts during launch-day sending.
SAFE_MIN_SLEEP_SECONDS = 15.0
SAFE_MAX_LIMIT_PER_RUN = 60
SAFE_MAX_DAILY_LIMIT = 120
SAFE_MAX_PER_HOUR_GLOBAL = 20
SAFE_MAX_PER_24H_GLOBAL = 90


def ensure_log_header(path):
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        with open(path, "w", encoding="utf-8") as f:
            f.write("campaign_id,email,sent_at,mode,subject,body_hash\n")


def append_line(path, line):
    with open(path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def ensure_log_xlsx(path):
    if os.path.exists(path):
        wb = load_workbook(path)
        ws = wb["Sent Log"] if "Sent Log" in wb.sheetnames else wb.active
        ws.title = "Sent Log"
        if ws.max_row < 1:
            ws.append(["campaign_id", "email", "sent_at", "mode", "subject", "body_hash"])
        elif ws.max_row == 1 and (ws.cell(row=1, column=1).value or "") != "campaign_id":
            ws.insert_rows(1)
            ws.cell(row=1, column=1, value="campaign_id")
            ws.cell(row=1, column=2, value="email")
            ws.cell(row=1, column=3, value="sent_at")
            ws.cell(row=1, column=4, value="mode")
            ws.cell(row=1, column=5, value="subject")
            ws.cell(row=1, column=6, value="body_hash")
        return wb, ws

    wb = Workbook()
    ws = wb.active
    ws.title = "Sent Log"
    ws.append(["campaign_id", "email", "sent_at", "mode", "subject", "body_hash"])
    wb.save(path)
    return wb, ws


def sync_text_log_to_xlsx(log_txt_path, log_xlsx_path):
    """Mirror launch_day_sent_log.txt into launch_day_sent_log.xlsx safely."""
    wb, ws = ensure_log_xlsx(log_xlsx_path)

    existing = set()
    for r in range(2, ws.max_row + 1):
        key = tuple((ws.cell(row=r, column=c).value or "").strip() for c in range(1, 7))
        if any(key):
            existing.add(key)

    added = 0
    if os.path.exists(log_txt_path):
        with open(log_txt_path, encoding="utf-8", errors="ignore") as f:
            raw = f.read().replace("\\n", "\n")
        for line in raw.splitlines():
            line = line.strip()
            if not line or line.lower().startswith("campaign_id,"):
                continue
            parts = [p.strip() for p in line.split(",")]
            if len(parts) < 6:
                continue
            row = (parts[0], parts[1], parts[2], parts[3], parts[4], parts[5])
            if row in existing:
                continue
            ws.append(list(row))
            existing.add(row)
            added += 1

    wb.save(log_xlsx_path)
    wb.close()
    return added


def append_spreadsheet_log(path, campaign_id, email, sent_at, mode, subject, body_hash):
    wb, ws = ensure_log_xlsx(path)
    ws.append([campaign_id, email, sent_at, mode, subject, body_hash])
    wb.save(path)
    wb.close()


def parse_log(path):
    """Return (sent_by_campaign, sent_today_by_campaign)."""
    sent = {}
    sent_today = {}
    today = datetime.now().strftime("%Y-%m-%d")
    if not os.path.exists(path):
        return sent, sent_today

    # Backward-compatible parser: handles both real newlines and legacy logs
    # that accidentally contain literal "\\n" separators.
    with open(path, encoding="utf-8", errors="ignore") as f:
        raw = f.read().replace("\\n", "\n")

    for line in raw.splitlines():
        line = line.strip()
        if not line or line.lower().startswith("campaign_id,"):
            continue
        parts = [p.strip() for p in line.split(",")]
        if len(parts) < 3:
            continue
        cid, email, sent_at = parts[0], parts[1].lower(), parts[2]
        sent.setdefault(cid, set()).add(email)
        if sent_at[:10] == today:
            sent_today[cid] = sent_today.get(cid, 0) + 1
    return sent, sent_today


def parse_global_send_history(path):
    """Return (sent_any, latest_sent_at_by_email) across all campaigns."""
    sent_any = set()
    latest_by_email = {}
    if not os.path.exists(path):
        return sent_any, latest_by_email

    with open(path, encoding="utf-8", errors="ignore") as f:
        raw = f.read().replace("\\n", "\n")

    for line in raw.splitlines():
        line = line.strip()
        if not line or line.lower().startswith("campaign_id,"):
            continue
        parts = [p.strip() for p in line.split(",")]
        if len(parts) < 3:
            continue
        email = parts[1].lower()
        if not email:
            continue
        sent_any.add(email)
        ts = parse_sent_timestamp(parts[2])
        if ts is None:
            continue
        prev = latest_by_email.get(email)
        if prev is None or ts > prev:
            latest_by_email[email] = ts

    return sent_any, latest_by_email


def parse_sent_timestamp(value):
    value = (value or "").strip()
    if not value:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None


def recent_global_send_counts(path, now):
    """Return sends in the last hour and last 24h across all campaigns."""
    last_hour = 0
    last_24h = 0
    if not os.path.exists(path):
        return last_hour, last_24h

    with open(path, encoding="utf-8", errors="ignore") as f:
        raw = f.read().replace("\\n", "\n")

    for line in raw.splitlines():
        line = line.strip()
        if not line or line.lower().startswith("campaign_id,"):
            continue
        parts = [p.strip() for p in line.split(",")]
        if len(parts) < 3:
            continue
        ts = parse_sent_timestamp(parts[2])
        if ts is None or ts > now:
            continue
        age_seconds = (now - ts).total_seconds()
        if age_seconds <= 3600:
            last_hour += 1
        if age_seconds <= 86400:
            last_24h += 1
    return last_hour, last_24h


def body_fingerprint(text):
    norm = " ".join((text or "").split()).lower()
    return hashlib.sha1(norm.encode("utf-8")).hexdigest()[:12]


def first_name_from_email(email):
    local = (email or "").split("@")[0]
    token = local.replace(".", " ").replace("_", " ").replace("-", " ").split()
    if not token:
        return "there"
    return token[0].capitalize()


def is_gmail_throttle_error(exc):
    """Best-effort detection of Gmail throttle/quota/auth lock signals."""

    # Walk the exception chain so wrapped SMTP errors are still detected.
    chain = []
    cur = exc
    while cur is not None and cur not in chain:
        chain.append(cur)
        cur = getattr(cur, "__cause__", None) or getattr(cur, "__context__", None)

    # Known SMTP response codes that are commonly seen for temporary send
    # throttling, quota pressure, or authentication lockouts.
    risky_smtp_codes = {
        421,
        450,
        451,
        452,
        454,
        534,
        535,
    }

    for e in chain:
        if isinstance(e, smtplib.SMTPResponseException):
            code = int(getattr(e, "smtp_code", 0) or 0)
            if code in risky_smtp_codes:
                return True

    msgs = []
    for e in chain:
        msgs.append(str(e or "").lower())
        if isinstance(e, smtplib.SMTPResponseException):
            raw = getattr(e, "smtp_error", b"")
            if isinstance(raw, bytes):
                msgs.append(raw.decode("utf-8", errors="ignore").lower())

    msg = "\n".join(msgs)
    signals = (
        "too many",
        "rate",
        "quota",
        "temporarily",
        "try again later",
        "daily user sending limit exceeded",
        "too many login attempts",
        "application-specific password required",
        "please log in via your web browser",
        "unusual activity",
        "temporarily locked",
        "suspended",
        "locked",
        "invalid credentials",
        "authentication failed",
        "4.7.0",
        "4.7.14",
        "4.7.26",
        "4.7.28",
        "5.7.0",
        "5.7.1",
        "5.7.14",
    )
    return any(s in msg for s in signals)


def read_text(path, fallback=""):
    if path and os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return f.read()
    return fallback


def text_to_html(text):
    esc = (
        (text or "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace(BOOKSPROUT_LINK, f'<a href="{BOOKSPROUT_LINK}">{BOOKSPROUT_LINK}</a>')
    )
    paras = [p.strip() for p in esc.split("\\n\\n") if p.strip()]
    html_paras = []
    for p in paras:
        html_paras.append("<p>" + p.replace("\\n", "<br>") + "</p>")
    return (
        '<div style="font-family:Georgia,serif;font-size:16px;line-height:1.55;color:#222">'
        + "".join(html_paras)
        + "</div>"
    )


class SafeMap(dict):
    def __missing__(self, key):
        return "{" + key + "}"


def render_template(raw, fields):
    return (raw or "").format_map(SafeMap(fields))


def get_header_map(ws):
    cols = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v:
            cols[str(v).strip()] = c
    if "email" not in cols:
        sys.exit("followups.xlsx is missing the 'email' column.")
    return cols


def load_recipients(sheet_path, include_responded, exclude_never, suppressed):
    if not os.path.exists(sheet_path):
        sys.exit("followups.xlsx not found. Run build_control_sheet.py first.")

    try:
        wb = load_workbook(sheet_path, read_only=True, data_only=True)
    except PermissionError:
        sys.exit("Could not open followups.xlsx. Close Excel and try again.")

    ws = wb["Follow-ups"] if "Follow-ups" in wb.sheetnames else wb.active
    cols = get_header_map(ws)

    def colv(row, name):
        idx = cols.get(name)
        if idx is None:
            return ""
        value = row[idx - 1]
        return "" if value is None else str(value).strip()

    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        email = colv(row, "email").lower()
        if not email or "@" not in email:
            continue
        if email in suppressed:
            continue

        if exclude_never:
            never = colv(row, "Never Follow-up?").lower()
            if never == "yes":
                continue

        responded = colv(row, "responded").lower()
        if include_responded == "yes" and responded != "yes":
            continue
        if include_responded == "no" and responded == "yes":
            continue

        out.append(
            {
                "email": email,
                "responded": responded,
            }
        )

    wb.close()
    return out


def build_standalone(sender, to_addr, subject, text_body, html_body,
                     related_images=None):
    em = EmailMessage()
    em["From"] = sender
    em["To"] = to_addr
    em["Subject"] = subject
    em.set_content(text_body)
    em.add_alternative(html_body, subtype="html")

    if related_images:
        html_part = em.get_payload()[-1]
        for img in related_images:
            ctype = mimetypes.guess_type(img.get("name") or "image.png")[0] or "image/png"
            maintype, subtype = ctype.split("/", 1)
            html_part.add_related(
                img["bytes"],
                maintype=maintype,
                subtype=subtype,
                cid=f"<{img['cid']}>",
                filename=img.get("name") or "image.png",
                disposition="inline",
            )
    return em


def main():
    ap = argparse.ArgumentParser(
        description="Send launch-day emails in safe chunks with daily caps."
    )
    ap.add_argument("--send", action="store_true", help="Actually send messages.")
    ap.add_argument("--sheet", default=SHEET, help="Path to followups.xlsx")
    ap.add_argument("--text", default=DEFAULT_TEXT, help="Plain-text template file")
    ap.add_argument("--html", default=DEFAULT_HTML, help="HTML template file")
    ap.add_argument("--campaign-id", default=f"launch-{datetime.now():%Y-%m-%d}")
    ap.add_argument(
        "--to-address",
        default="",
        help="If set, send only to this single email address (sample/test mode).",
    )
    ap.add_argument("--limit", type=int, default=50, help="Max sends this run.")
    ap.add_argument(
        "--daily-limit",
        type=int,
        default=100,
        help="Max sends per calendar day for this campaign ID.",
    )
    ap.add_argument(
        "--global-skip-days",
        type=int,
        default=0,
        help="Skip recipients sent in any launch campaign within the last N days (0 disables).",
    )
    ap.add_argument(
        "--global-skip-ever",
        action="store_true",
        help="Skip recipients that were ever sent in any launch campaign.",
    )
    ap.add_argument(
        "--include-responded",
        choices=("all", "yes", "no"),
        default="no",
        help="all = everyone in sheet, yes = only responders, no = only non-responders",
    )
    ap.add_argument(
        "--phase-2",
        action="store_true",
        help="Allow responders only after all non-responders are exhausted for this campaign.",
    )
    ap.add_argument(
        "--include-never",
        action="store_true",
        help="Include rows marked Never Follow-up?=yes (excluded by default).",
    )
    ap.add_argument(
        "--standalone",
        action="store_true",
        default=True,
        help="Send as a new message instead of a threaded reply.",
    )
    ap.add_argument(
        "--threaded",
        action="store_true",
        help="Override default standalone mode and send as threaded replies.",
    )
    ap.add_argument(
        "--subject",
        default="My new novel, Miner Town: Awakening, is out now",
        help="Subject used only with --standalone.",
    )
    ap.add_argument(
        "--cover-image",
        default=DEFAULT_COVER_IMAGE,
        help="Local path to cover image for inline email embedding (standalone mode).",
    )
    ap.add_argument(
        "--cover-gif",
        default=DEFAULT_COVER_GIF,
        help="Optional local path to animated GIF overlay. If present, used first with static fallback.",
    )
    ap.add_argument("--orig-subject", default=ORIGINAL_SUBJECT)
    ap.add_argument("--sent-folder", default=SENT_FOLDER)
    ap.add_argument(
        "--sleep-min", type=float, default=20.0, help="Min seconds between sends."
    )
    ap.add_argument(
        "--sleep-max", type=float, default=45.0, help="Max seconds between sends."
    )
    ap.add_argument(
        "--override-guardrail",
        action="store_true",
        help="Bypass launch safety rails (not recommended; increases lockout risk).",
    )
    ap.add_argument(
        "--run-cooldown-minutes",
        type=int,
        default=10,
        help="Minimum minutes between any two send runs (default 10).",
    )
    args = ap.parse_args()

    if args.threaded:
        args.standalone = False

    if args.limit <= 0:
        sys.exit("--limit must be > 0")
    if args.daily_limit <= 0:
        sys.exit("--daily-limit must be > 0")
    if args.global_skip_days < 0:
        sys.exit("--global-skip-days must be >= 0")
    if args.sleep_min < 0 or args.sleep_max < 0 or args.sleep_max < args.sleep_min:
        sys.exit("Sleep bounds must be non-negative and sleep-max >= sleep-min")

    user = os.environ.get("SMTP_USER")
    pw = os.environ.get("SMTP_PASS")
    if not user or not pw:
        sys.exit("Set SMTP_USER and SMTP_PASS (Gmail App Password) first.")

    if args.send:
        enforce_send_run_cooldown(
            SEND_COOLDOWN_LOCK,
            args.run_cooldown_minutes,
            "launch_day_campaign.py",
        )

    text_raw = read_text(args.text)
    if not text_raw.strip():
        sys.exit(f"Template text file is empty or missing: {args.text}")

    html_raw = read_text(args.html)

    fields_base = {
        "booksprout_link": BOOKSPROUT_LINK,
        "launch_date": datetime.now().strftime("%Y-%m-%d"),
        "launch_day": datetime.now().strftime("%A"),
    }

    suppressed = load_set(SUPPRESSION)
    exclude_never = not args.include_never

    sent_by_campaign, sent_today_by_campaign = parse_log(LOG_PATH)
    sent_any_global, latest_global_by_email = parse_global_send_history(LOG_PATH)
    already = sent_by_campaign.get(args.campaign_id, set())
    sent_today = sent_today_by_campaign.get(args.campaign_id, 0)

    now = datetime.now()
    global_cutoff = (
        now - timedelta(days=args.global_skip_days) if args.global_skip_days > 0 else None
    )
    sent_last_hour, sent_last_24h = recent_global_send_counts(LOG_PATH, now)

    def blocked_by_global_guardrail(email):
        if args.to_address:
            return False
        if args.global_skip_ever and email in sent_any_global:
            return True
        if global_cutoff is not None:
            last_sent = latest_global_by_email.get(email)
            if last_sent is not None and last_sent >= global_cutoff:
                return True
        return False

    if args.send and not args.override_guardrail and not args.to_address:
        if args.sleep_min < SAFE_MIN_SLEEP_SECONDS or args.sleep_max < SAFE_MIN_SLEEP_SECONDS:
            sys.exit(
                "Guardrail blocked send: sleep-min and sleep-max must both be "
                f">= {SAFE_MIN_SLEEP_SECONDS:.0f}s. "
                "Use slower pacing or pass --override-guardrail."
            )
        if args.limit > SAFE_MAX_LIMIT_PER_RUN:
            sys.exit(
                "Guardrail blocked send: --limit is too high "
                f"({args.limit} > {SAFE_MAX_LIMIT_PER_RUN}). "
                "Reduce limit or pass --override-guardrail."
            )
        if args.daily_limit > SAFE_MAX_DAILY_LIMIT:
            sys.exit(
                "Guardrail blocked send: --daily-limit is too high "
                f"({args.daily_limit} > {SAFE_MAX_DAILY_LIMIT}). "
                "Reduce daily-limit or pass --override-guardrail."
            )
        if sent_last_hour >= SAFE_MAX_PER_HOUR_GLOBAL:
            sys.exit(
                "Guardrail blocked send: recent global volume is already high "
                f"({sent_last_hour} sent in the last hour). "
                "Wait before sending again, or pass --override-guardrail."
            )
        if sent_last_24h >= SAFE_MAX_PER_24H_GLOBAL:
            sys.exit(
                "Guardrail blocked send: recent global volume is already high "
                f"({sent_last_24h} sent in the last 24 hours). "
                "Wait until tomorrow, or pass --override-guardrail."
            )

    if args.to_address:
        recipients = [{"email": args.to_address.strip().lower(), "responded": "test"}]
    else:
        recipients = load_recipients(args.sheet, args.include_responded, exclude_never, suppressed)

    # Phase 2 gate: responders can only start once non-responders are exhausted
    # for this campaign (i.e., no pending non-responders left unsent).
    if args.phase_2 and not args.to_address:
        if args.include_responded != "yes":
            sys.exit("--phase-2 requires --include-responded yes.")
        non_resp = load_recipients(args.sheet, "no", exclude_never, suppressed)
        pending_non_resp = [
            r
            for r in non_resp
            if r["email"] not in already and not blocked_by_global_guardrail(r["email"])
        ]
        if pending_non_resp:
            print("Phase 2 is locked: non-responders are not exhausted yet.")
            print(f"Pending non-responders remaining for campaign: {len(pending_non_resp)}")
            print("Finish Phase 1 first, then re-run Phase 2.")
            return

    pending = []
    skipped_global = 0
    for r in recipients:
        email = r["email"]
        if email in already:
            continue
        if blocked_by_global_guardrail(email):
            skipped_global += 1
            continue
        pending.append(r)

    remaining_today = max(0, args.daily_limit - sent_today)
    run_cap = min(args.limit, remaining_today, len(pending))

    print(f"Campaign ID: {args.campaign_id}")
    print(f"Mode: {'SEND' if args.send else 'DRY RUN'}")
    print(f"Recipients in sheet (after filters): {len(recipients)}")
    print(f"Already sent for campaign: {len(already)}")
    if args.global_skip_ever:
        print("Global de-dupe: skip-ever enabled")
    elif args.global_skip_days > 0:
        print(f"Global de-dupe: skip last {args.global_skip_days} day(s) enabled")
    else:
        print("Global de-dupe: disabled")
    if args.global_skip_ever or args.global_skip_days > 0:
        print(f"Skipped by global de-dupe: {skipped_global}")
    print(f"Pending: {len(pending)}")
    print(f"Sent today for campaign: {sent_today} / {args.daily_limit}")
    print(f"Global recent sends: {sent_last_hour} last hour, {sent_last_24h} last 24h")
    print(
        "Guardrails: "
        f"min-delay={SAFE_MIN_SLEEP_SECONDS:.0f}s "
        f"run-limit<={SAFE_MAX_LIMIT_PER_RUN} "
        f"daily-limit<={SAFE_MAX_DAILY_LIMIT} "
        f"hourly<={SAFE_MAX_PER_HOUR_GLOBAL} "
        f"24h<={SAFE_MAX_PER_24H_GLOBAL}"
        + (" (OVERRIDDEN)" if args.override_guardrail else "")
    )
    print(f"This run cap: {run_cap}\n")

    if not pending:
        if skipped_global and (args.global_skip_ever or args.global_skip_days > 0):
            print("Nothing pending because global de-dupe guardrail filtered remaining recipients.")
        print("Nothing pending for this campaign.")
        return
    if remaining_today <= 0:
        print("Daily cap reached for this campaign. Re-run tomorrow.")
        return

    preview = pending[: min(run_cap, 20)]
    for r in preview:
        print(f"[queued] {r['email']} (responded={r['responded'] or 'unknown'})")
    if run_cap > len(preview):
        print(f"... plus {run_cap - len(preview)} more queued for this run")

    if not args.send:
        print("\nDry run only. Re-run with --send to deliver.")
        return

    M = None
    smtp = None
    try:
        if not args.standalone:
            M = imap_connect(user, pw)
        smtp = smtplib.SMTP_SSL(
            SMTP_HOST,
            SMTP_PORT,
            timeout=SMTP_TIMEOUT,
            context=ssl.create_default_context(),
        )
        smtp.login(user, pw)
    except Exception as e:
        if M is not None:
            try:
                M.logout()
            except Exception:
                pass
        sys.exit(f"Could not connect to Gmail: {e}")

    ensure_log_header(LOG_PATH)
    try:
        synced = sync_text_log_to_xlsx(LOG_PATH, LOG_XLSX_PATH)
        if synced:
            print(f"[guardrail] Synced {synced} launch log row(s) into launch_day_sent_log.xlsx")
    except Exception as e:
        sys.exit(f"Could not initialize launch_day_sent_log.xlsx: {e}")

    sent_now = 0
    errors = 0
    throttle_blocked = False

    static_cover_bytes = None
    static_cover_name = None
    static_cover_cid = "launch-cover-static"
    gif_cover_bytes = None
    gif_cover_name = None
    gif_cover_cid = "launch-cover-gif"
    cover_block = ""

    if args.cover_image and os.path.exists(args.cover_image):
        with open(args.cover_image, "rb") as f:
            static_cover_bytes = f.read()
        static_cover_name = os.path.basename(args.cover_image)
    elif args.cover_image:
        print(f"[warn] Cover image not found, sending without inline cover: {args.cover_image}")

    if args.cover_gif and os.path.exists(args.cover_gif):
        with open(args.cover_gif, "rb") as f:
            gif_cover_bytes = f.read()
        gif_cover_name = os.path.basename(args.cover_gif)
    elif args.cover_gif:
        print(f"[warn] Cover GIF not found, using static fallback only: {args.cover_gif}")

    use_gif = False
    use_static = False

    if gif_cover_bytes:
        # Email-safe approach: use GIF directly as the inline <img> src.
        # Most clients that don't animate GIFs will still show frame 1.
        use_gif = True
        cover_block = (
            '<p style="text-align:center;margin:0 0 14px 0;">'
            '<a href="https://ankitsaxenabooks.netlify.app/" target="_blank">'
            f'<img src="cid:{gif_cover_cid}" alt="Miner Town: Awakening cover" '
            'style="display:block;width:100%;max-width:360px;height:auto;border:0;margin:0 auto;">'
            '</a></p>'
        )
    elif static_cover_bytes:
        use_static = True
        cover_block = (
            '<p style="text-align:center;margin:0 0 14px 0;">'
            '<a href="https://ankitsaxenabooks.netlify.app/" target="_blank">'
            f'<img src="cid:{static_cover_cid}" alt="Miner Town: Awakening cover" '
            'style="display:block;width:100%;max-width:360px;height:auto;border:0;margin:0 auto;">'
            '</a></p>'
        )

    for r in pending[:run_cap]:
        email = r["email"]
        fields = dict(fields_base)
        fields["first_name"] = first_name_from_email(email)
        fields["cover_image_block"] = cover_block

        text_body = render_template(text_raw, fields)
        html_tmpl = html_raw if html_raw.strip() else text_to_html(text_raw)
        html_body = render_template(html_tmpl, fields)

        body_hash = body_fingerprint(text_body)
        mode = "standalone" if args.standalone else "subject-threaded"

        try:
            if args.standalone:
                related_images = []
                if use_static and static_cover_bytes:
                    related_images.append({
                        "bytes": static_cover_bytes,
                        "name": static_cover_name,
                        "cid": static_cover_cid,
                    })
                if use_gif and gif_cover_bytes:
                    related_images.append({
                        "bytes": gif_cover_bytes,
                        "name": gif_cover_name,
                        "cid": gif_cover_cid,
                    })
                msg = build_standalone(
                    user,
                    email,
                    args.subject,
                    text_body,
                    html_body,
                    related_images=related_images,
                )
            else:
                orig = find_original(M, email, args.sent_folder, args.orig_subject)
                mode = "threaded" if (orig and orig.get("message_id")) else "subject-threaded"
                msg = build_reply(
                    user,
                    email,
                    text_body,
                    html_body,
                    {},
                    orig,
                    args.orig_subject,
                    preformatted=True,
                )

            smtp.send_message(msg)
            stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
            subject = args.subject if args.standalone else f"Re: {args.orig_subject}"
            append_line(
                LOG_PATH,
                f"{args.campaign_id},{email},{stamp},{mode},{subject},{body_hash}",
            )
            append_spreadsheet_log(
                LOG_XLSX_PATH,
                args.campaign_id,
                email,
                stamp,
                mode,
                subject,
                body_hash,
            )
            sent_now += 1
            print(f"[sent] {email} ({mode})")
            time.sleep(random.uniform(args.sleep_min, args.sleep_max))
        except Exception as e:
            errors += 1
            print(f"[ERROR] {email}: {e}")
            if is_gmail_throttle_error(e):
                throttle_blocked = True
                print(
                    "[guardrail] Gmail throttle/auth signal detected; "
                    "stopping this run to avoid account lock risk."
                )
                break

    if smtp is not None:
        try:
            smtp.quit()
        except Exception:
            pass
    if M is not None:
        try:
            M.logout()
        except Exception:
            pass

    print("\n--- launch summary ---")
    print(f"Sent now: {sent_now}")
    print(f"Errors:   {errors}")
    print(f"Remaining for campaign: {max(0, len(pending) - sent_now)}")
    if args.send:
        mark_send_run_finished(SEND_COOLDOWN_LOCK, sent_now)
    if sent_now >= remaining_today:
        print("Daily cap reached for this campaign. Continue tomorrow.")
    if throttle_blocked:
        sys.exit(
            "Guardrail stopped send: Gmail returned throttle/auth/quota signal. "
            "Wait before retrying."
        )


if __name__ == "__main__":
    main()
