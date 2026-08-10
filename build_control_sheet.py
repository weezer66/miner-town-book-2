#!/usr/bin/env python3
"""
build_control_sheet.py — turn the response report into a follow-up control sheet.

Reads responded.csv + no_response.csv and writes followups.xlsx with two tabs:

  'Follow-ups' tab — one row per recipient. The last column, 'Conversation Log',
    is a timestamped two-way history of your sends (→) and their replies (←).
    Each row also has:
    Follow-up?       No (default) / Yes
    Template         which template to send (names come from the Templates tab)
    Custom Message   OR type your own note here for that person; if filled, it is
                     sent instead of the template (you can use {booksprout_link}).

  'Templates' tab — the full text of each template, editable. Edits here change
    the template for EVERYONE who uses it; the send step reads from this tab.

After you mark your choices and save, run the "Execute Follow-ups" launcher.

Auto-queue: any recipient with AUTO_FOLLOWUP_DAYS (7) of no activity — no reply
from them and no message from you — has Follow-up? flipped to Yes automatically
(highlighted yellow), so the next execute run sends another nudge. The Template
follows a FIXED sequence by how many follow-ups they've already had:
AUTO_TEMPLATE_SEQUENCE = 1st -> gentle-nudge, 2nd -> still-interested. Once the
still-interested nudge is actually sent, execute_followups.py marks the person
Never Follow-up? = yes (stop chasing), so they won't be auto-queued again. Stop-guardrails are honored: people
who have the book, opted out / said not interested, are suppressed, or are
marked 'Never Follow-up?' = yes are never auto-queued.

Re-running this builder is SAFE: it preserves your Follow-up?/Template/Custom
Message/Status choices AND your edited template text. New recipients and new
templates/<name>.txt files are added with defaults.

Needs: openpyxl  (pip install openpyxl)
"""
import csv
import glob
import os
import re
import sys
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

HERE = os.path.dirname(os.path.abspath(__file__))
RESPONDED = os.path.join(HERE, "responded.csv")
NO_RESPONSE = os.path.join(HERE, "no_response.csv")
TEMPLATES_DIR = os.path.join(HERE, "templates")
OUT = os.path.join(HERE, "followups.xlsx")

# Column order in the sheet. 'Conversation Log' is intentionally LAST.
HEADERS = ["email", "responded", "Has Book?", "Needs Help?", "emailed_on",
           "latest_reply_on", "Latest Message", "Never Follow-up?", "Follow-up?",
           "Template", "Custom Message", "Allow Repeat?", "Follow-ups Sent",
           "Follow-up History", "Status", "Conversation Log"]
COL = {h: i + 1 for i, h in enumerate(HEADERS)}   # 1-based column index

FOLLOWUP_LOG = os.path.join(HERE, "followup_sent_log.txt")
HAS_BOOK_FILE = os.path.join(HERE, "has_book.txt")
NEEDS_HELP_FILE = os.path.join(HERE, "needs_help.txt")
SUPPRESSION = os.path.join(HERE, "suppression.txt")

# Auto-queue a follow-up when there's been this many days of no activity
# (no reply from them and no message from you).
AUTO_FOLLOWUP_DAYS = 7

# Fixed template progression for auto-queued nudges, by how many follow-ups the
# person has already received:  1st follow-up -> gentle-nudge, 2nd -> still-
# interested.  Past the end of this list the Template is left BLANK so YOU pick
# (or write a custom message). Not random — a defined sequence.
AUTO_TEMPLATE_SEQUENCE = ["gentle-nudge", "still-interested"]

_TS_IN_LINE = re.compile(r"\d{4}-\d{2}-\d{2}(?: \d{2}:\d{2})?")


def last_activity_dt(row, history):
    """Most recent message in EITHER direction for this recipient — pooled from
    the conversation log, emailed_on/latest_reply_on, and the follow-up log
    (so a just-sent follow-up counts even if the report hasn't been re-run)."""
    dts = []
    for m in _TS_IN_LINE.findall(row.get("conversation") or ""):
        d = parse_dt_any(m)
        if d:
            dts.append(d)
    for v in (row.get("emailed_on"), row.get("latest_reply_on")):
        d = parse_dt_any(v)
        if d:
            dts.append(d)
    for ts, _src in history.get(row["email"], []):
        d = parse_dt_any(ts)
        if d:
            dts.append(d)
    return max(dts) if dts else None


def parse_dt_any(s):
    """Parse any timestamp format we might see — including dd-mm-yyyy, which is
    what Excel writes back into the CSVs under an Indian/European locale."""
    s = (s or "").strip()
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d",       # our native format
                "%d-%m-%Y %H:%M", "%d-%m-%Y",       # Excel locale re-save
                "%m/%d/%Y %H:%M", "%m/%d/%Y"):      # reply-comment stamps
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def norm_ts(s):
    """Normalize a timestamp string to YYYY-MM-DD (+ HH:MM if it had a time)."""
    dt = parse_dt_any(s)
    if not dt:
        return (s or "").strip()
    return dt.strftime("%Y-%m-%d %H:%M") if ":" in (s or "") else dt.strftime("%Y-%m-%d")


_REPLY_LINE = re.compile(r"^(\d{2}/\d{2}/\d{4}|\?\?/\?\?/\?\?\?\?)\s*-\s*(.*)$")


def build_conversation(emailed_on, comments, followups):
    """Interleave outbound (your original + follow-ups) and inbound (their
    replies) into one timestamped, chronological two-way log.
        2026-06-10        →  SENT: original outreach
        2026-06-10        ←  REPLY: congratulations!!
        2026-06-19 14:30  →  SENT: follow-up (still-interested)
    """
    def disp(dt, raw, with_time):
        if dt:
            return dt.strftime("%Y-%m-%d %H:%M" if with_time else "%Y-%m-%d")
        return raw

    events = []   # (sort_dt, display_ts, body)
    if emailed_on:
        dt0 = parse_dt_any(emailed_on)
        events.append((dt0 or datetime.min, disp(dt0, emailed_on, False),
                       "→  SENT: original outreach"))
    for ts, source in (followups or []):
        dt = parse_dt_any(ts)
        label = "→  SENT: follow-up" + (f" ({source})" if source else "")
        events.append((dt or datetime.min, disp(dt, ts, True), label))
    for line in (comments or "").split("\n"):
        line = line.strip()
        if not line:
            continue
        m = _REPLY_LINE.match(line)
        if m:
            dt = parse_dt_any(m.group(1))
            events.append((dt or datetime.min, disp(dt, m.group(1), False),
                           f"←  REPLY: {m.group(2)}"))
        else:
            events.append((datetime.min, "", f"←  REPLY: {line}"))
    events.sort(key=lambda e: e[0])
    return "\n".join((f"{ts}  {body}" if ts else body) for _, ts, body in events)


def load_email_list(path):
    """Set of emails from a one-per-line file (# comments ignored)."""
    out = set()
    if os.path.exists(path):
        with open(path, encoding="utf-8", errors="ignore") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "@" in line:
                    out.add(line.lower())
    return out


def read_csv(path):
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f))


def parse_followup_log(path):
    """email -> list of (timestamp, source) for every follow-up ever sent,
    read from followup_sent_log.txt (columns: email, sent_at, source, mode)."""
    hist = {}
    if not os.path.exists(path):
        return hist
    with open(path, encoding="utf-8", errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line or line.lower().startswith("email,"):
                continue
            parts = [p.strip() for p in line.split(",")]
            email = parts[0].lower()
            ts = parts[1] if len(parts) > 1 else ""
            source = parts[2] if len(parts) > 2 else ""
            hist.setdefault(email, []).append((ts, source))
    return hist


TEMPLATES_SHEET = "Templates"


def file_template_bodies():
    """name -> text, seeded from templates/*.txt on disk."""
    bodies = {}
    for p in sorted(glob.glob(os.path.join(TEMPLATES_DIR, "*.txt"))):
        name = os.path.splitext(os.path.basename(p))[0]
        with open(p, encoding="utf-8") as f:
            bodies[name] = f.read().rstrip("\n")
    return bodies


def load_existing_templates(path):
    """name -> text from a prior followups.xlsx 'Templates' tab (user edits)."""
    bodies = {}
    if not os.path.exists(path):
        return bodies
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        if TEMPLATES_SHEET not in wb.sheetnames:
            wb.close(); return bodies
        ws = wb[TEMPLATES_SHEET]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            name = str(row[0]).strip()
            if name == "(read me)":          # skip the help-note row
                continue
            body = row[1] if len(row) > 1 and row[1] is not None else ""
            bodies[name] = str(body)
        wb.close()
    except Exception as e:
        print(f"(could not read existing Templates tab: {e})")
    return bodies


def merged_template_bodies():
    """Template bodies to write into the sheet: disk seeds, with any prior
    in-sheet edits taking precedence (so your edits survive a rebuild)."""
    bodies = file_template_bodies()
    bodies.update(load_existing_templates(OUT))   # user edits win
    if not bodies:
        print("!! No templates found — add <name>.txt files in templates/.")
    return bodies


def load_existing_choices(path):
    """email -> dict of prior Follow-up?/Template/Custom Message/Status.

    Maps columns by HEADER NAME from the existing file, so it survives changes
    to column order between versions of this script.
    """
    choices = {}
    if not os.path.exists(path):
        return choices
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        header = [str(c.value).strip() if c.value else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]
        idx = {name: i for i, name in enumerate(header)}

        def get(row, name):
            i = idx.get(name)
            return (row[i] if (i is not None and i < len(row) and row[i] is not None) else "")

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            email = str(row[0]).strip().lower()
            choices[email] = {
                "Never Follow-up?": get(row, "Never Follow-up?"),
                "Follow-up?": get(row, "Follow-up?"),
                "Template": get(row, "Template"),
                "Custom Message": get(row, "Custom Message"),
                "Allow Repeat?": get(row, "Allow Repeat?"),
                "Status": get(row, "Status"),
            }
        wb.close()
        print(f"Preserved choices for {len(choices)} existing row(s).")
    except Exception as e:
        print(f"(could not read existing {os.path.basename(path)}: {e})")
    return choices


def main():
    responded = read_csv(RESPONDED)
    no_response = read_csv(NO_RESPONSE)
    if not responded and not no_response:
        sys.exit("No responded.csv / no_response.csv found. Run response_report.py first.")

    template_bodies = merged_template_bodies()
    templates = sorted(template_bodies)
    prior = load_existing_choices(OUT)
    history = parse_followup_log(FOLLOWUP_LOG)
    has_book = load_email_list(HAS_BOOK_FILE)
    needs_help = load_email_list(NEEDS_HELP_FILE)
    suppressed = load_email_list(SUPPRESSION)
    opted = {r.get("email", "").strip().lower() for r in responded
             if r.get("opted_out", "").strip().lower() == "yes"}
    blocked = has_book | opted | suppressed   # never auto-queue these
    now = datetime.now()
    auto_count = 0

    # Merge both CSVs into one list of recipient dicts.
    rows = []
    for r in responded:
        rows.append({
            "email": r.get("email", "").strip().lower(),
            "responded": "yes",
            "emailed_on": r.get("emailed_on", ""),
            "latest_reply_on": r.get("latest_reply_on", ""),
            "comments": r.get("comments", ""),
            "conversation": r.get("conversation", ""),
        })
    for r in no_response:
        rows.append({
            "email": r.get("email", "").strip().lower(),
            "responded": "no",
            "emailed_on": r.get("emailed_on", ""),
            "latest_reply_on": "",
            "comments": "",
            "conversation": r.get("conversation", ""),
        })

    # De-dupe by email (responded wins), then sort: responders first, then email.
    by_email = {}
    for row in rows:
        e = row["email"]
        if not e:
            continue
        if e not in by_email or row["responded"] == "yes":
            by_email[e] = row
    ordered = sorted(by_email.values(),
                     key=lambda r: (0 if r["responded"] == "yes" else 1, r["email"]))

    wb = Workbook()
    ws = wb.active
    ws.title = "Follow-ups"

    header_fill = PatternFill("solid", fgColor="1F3864")
    header_font = Font(bold=True, color="FFFFFF")
    for h, c in COL.items():
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center")

    yes_fill = PatternFill("solid", fgColor="E2EFDA")   # light green for chosen
    wrap = Alignment(wrap_text=True, vertical="top")

    book_fill = PatternFill("solid", fgColor="FCE4D6")   # peach = already has book
    help_fill = PatternFill("solid", fgColor="FFC7CE")   # red = needs help
    auto_fill = PatternFill("solid", fgColor="FFF2CC")   # yellow = auto-queued
    awaiting_fill = PatternFill("solid", fgColor="FF6B6B")  # red = their reply is last (your turn)
    for i, row in enumerate(ordered, start=2):
        ws.cell(row=i, column=COL["email"], value=row["email"])
        ws.cell(row=i, column=COL["responded"], value=row["responded"])
        hb = ws.cell(row=i, column=COL["Has Book?"],
                     value="yes" if row["email"] in has_book else "")
        if row["email"] in has_book:
            hb.fill = book_fill
        nh = ws.cell(row=i, column=COL["Needs Help?"],
                     value="yes" if row["email"] in needs_help else "")
        if row["email"] in needs_help:
            nh.fill = help_fill
        ws.cell(row=i, column=COL["emailed_on"], value=norm_ts(row["emailed_on"]))
        ws.cell(row=i, column=COL["latest_reply_on"], value=norm_ts(row["latest_reply_on"]))
        # Prefer the full two-way conversation from the report (includes your
        # manual Sent-folder replies); fall back to reconstruction for old CSVs.
        convo = row.get("conversation") or build_conversation(
            row["emailed_on"], row["comments"], history.get(row["email"], []))
        cc = ws.cell(row=i, column=COL["Conversation Log"], value=convo)
        cc.alignment = wrap

        # Latest message in the thread. Turn the cell RED (black text) when the
        # last message was FROM THEM (a '← REPLY') — i.e., it's YOUR turn to
        # reply, so you don't miss it.
        last_line = ""
        for ln in convo.split("\n"):
            if ln.strip():
                last_line = ln.strip()
        pin = last_line.find("←")   # inbound marker
        pout = last_line.find("→")  # outbound marker
        inbound_last = pin != -1 and (pout == -1 or pin < pout)
        lm = ws.cell(row=i, column=COL["Latest Message"], value=last_line)
        lm.alignment = wrap
        if inbound_last:
            lm.fill = awaiting_fill
            lm.font = Font(color="000000", bold=True)

        prev = prior.get(row["email"], {})
        # 'Never Follow-up?' is MANUAL ONLY — preserved verbatim, never set by code.
        never = (prev.get("Never Follow-up?") or "").strip().lower()
        never = "yes" if never == "yes" else "no"
        ws.cell(row=i, column=COL["Never Follow-up?"], value=never)

        followup = prev.get("Follow-up?") if prev.get("Follow-up?") in ("No", "Yes") else "No"
        template = prev.get("Template", "")
        custom = prev.get("Custom Message", "")
        allow_repeat = (prev.get("Allow Repeat?") or "").strip().lower()
        allow_repeat = "yes" if allow_repeat == "yes" else "no"

        # Templates this person has ALREADY received (don't repeat them).
        received = {src for _ts, src in history.get(row["email"], [])
                    if src and src != "custom"}

        # Auto-queue: flip to Yes after AUTO_FOLLOWUP_DAYS of no activity, unless
        # 'Never Follow-up?' = yes OR a stop-guardrail applies (has the book /
        # opted out / suppressed).
        # Follow-up? is AUTHORITATIVE on activity: queue only people idle for
        # AUTO_FOLLOWUP_DAYS+, and force No for anyone with recent activity — a
        # reply from them OR any message from you, INCLUDING manual Gmail sends
        # (last_activity_dt pools the conversation, so it sees those too). This
        # clears stale 'Yes' flags once a person becomes active again.
        auto = False
        if never == "yes" or row["email"] in blocked:
            followup = "No"            # never / has-book / opted-out / suppressed
        else:
            la = last_activity_dt(row, history)
            if la and (now - la).days >= AUTO_FOLLOWUP_DAYS:
                followup = "Yes"
                auto = True
            else:
                followup = "No"        # active within the window — don't nudge
        if auto:
            followup = "Yes"
            # Fixed sequence by how many follow-ups already sent:
            #   1st -> gentle-nudge, 2nd -> still-interested, then BLANK.
            # A not-yet-sent manual pick of yours is kept as-is. A template they
            # already received is never re-picked (falls to blank).
            if not custom:
                sent_count = len(history.get(row["email"], []))
                if template and template not in received:
                    pass                       # keep your manual, unsent choice
                elif sent_count < len(AUTO_TEMPLATE_SEQUENCE):
                    nxt = AUTO_TEMPLATE_SEQUENCE[sent_count]
                    template = nxt if nxt not in received else ""
                else:
                    template = ""              # past the sequence — you choose
            auto_count += 1

        fu_cell = ws.cell(row=i, column=COL["Follow-up?"], value=followup)
        if auto:
            fu_cell.fill = auto_fill
        ws.cell(row=i, column=COL["Template"], value=template)
        cm = ws.cell(row=i, column=COL["Custom Message"], value=custom)
        cm.alignment = wrap
        ws.cell(row=i, column=COL["Allow Repeat?"], value=allow_repeat)

        sends = history.get(row["email"], [])
        ws.cell(row=i, column=COL["Follow-ups Sent"], value=len(sends))
        hist_text = "\n".join(f"{ts} — {src}" if src else ts for ts, src in sends)
        hc = ws.cell(row=i, column=COL["Follow-up History"], value=hist_text)
        hc.alignment = wrap
        ws.cell(row=i, column=COL["Status"], value=prev.get("Status", ""))

    n = len(ordered) + 1   # last row

    # Dropdowns (column letters derived from header positions).
    fu_col = ws.cell(row=1, column=COL["Follow-up?"]).column_letter
    tm_col = ws.cell(row=1, column=COL["Template"]).column_letter
    nf_col = ws.cell(row=1, column=COL["Never Follow-up?"]).column_letter

    dv_followup = DataValidation(type="list", formula1='"No,Yes"', allow_blank=False)
    dv_followup.error = 'Choose No or Yes'
    dv_followup.prompt = 'Set Yes to send this person a follow-up'
    ws.add_data_validation(dv_followup)
    dv_followup.add(f"{fu_col}2:{fu_col}{n}")

    dv_never = DataValidation(type="list", formula1='"no,yes"', allow_blank=False)
    dv_never.error = 'Choose no or yes'
    dv_never.prompt = ('Set yes to PERMANENTLY exclude this person from all '
                       'follow-ups. Never changed by the program.')
    ws.add_data_validation(dv_never)
    dv_never.add(f"{nf_col}2:{nf_col}{n}")

    ar_col = ws.cell(row=1, column=COL["Allow Repeat?"]).column_letter
    dv_repeat = DataValidation(type="list", formula1='"no,yes"', allow_blank=False)
    dv_repeat.prompt = ('Set yes to allow re-sending a template this person '
                        'ALREADY received (explicit override). Resets after send.')
    ws.add_data_validation(dv_repeat)
    dv_repeat.add(f"{ar_col}2:{ar_col}{n}")

    # Turn the Never Follow-up? cell red whenever YOU set it to yes.
    # NOTE: conditional-formatting fills render by the *background* color, so
    # set start_color AND end_color (fgColor alone shows no fill in Excel).
    never_red = PatternFill(start_color="C00000", end_color="C00000",
                            fill_type="solid")
    ws.conditional_formatting.add(
        f"{nf_col}2:{nf_col}{n}",
        CellIsRule(operator="equal", formula=['"yes"'], fill=never_red,
                   font=Font(color="FFFFFF", bold=True)))

    if templates:
        tlist = ",".join(templates)
        dv_template = DataValidation(type="list", formula1=f'"{tlist}"', allow_blank=True)
        dv_template.prompt = 'Pick a template (only used when Follow-up? = Yes)'
        ws.add_data_validation(dv_template)
        dv_template.add(f"{tm_col}2:{tm_col}{n}")

    # Column widths + niceties.
    widths = {"email": 30, "responded": 11, "Has Book?": 10, "Needs Help?": 11,
              "emailed_on": 13, "latest_reply_on": 18, "Latest Message": 55,
              "Never Follow-up?": 15,
              "Follow-up?": 12, "Template": 18, "Custom Message": 55,
              "Allow Repeat?": 13, "Follow-ups Sent": 14, "Follow-up History": 34,
              "Status": 28, "Conversation Log": 80}
    for h, w in widths.items():
        ws.column_dimensions[ws.cell(row=1, column=COL[h]).column_letter].width = w
    ws.freeze_panes = "A2"
    last_col = ws.cell(row=1, column=len(HEADERS)).column_letter
    ws.auto_filter.ref = f"A1:{last_col}{n}"

    # ---- second tab: editable master templates ----
    tws = wb.create_sheet(TEMPLATES_SHEET)
    note = ("Edit the Body text below to change a template for EVERYONE who uses "
            "it. Placeholders: {booksprout_link} = your BookSprout link; "
            "{got_it_button} = 'I got it on BookSprout' reply button; "
            "{trouble_button} = 'Having trouble' reply button. The send step "
            "reads templates from THIS tab. (To add a brand-new template, save a "
            "templates/<name>.txt file and re-run build_control_sheet.)")
    tws.cell(row=1, column=1, value="Template").fill = header_fill
    tws.cell(row=1, column=1).font = header_font
    tws.cell(row=1, column=2, value="Body (edit here)").fill = header_fill
    tws.cell(row=1, column=2).font = header_font
    tws.cell(row=2, column=2, value=note).alignment = Alignment(wrap_text=True, vertical="top")
    tws.cell(row=2, column=1, value="(read me)")
    for r, name in enumerate(templates, start=3):
        tws.cell(row=r, column=1, value=name)
        bcell = tws.cell(row=r, column=2, value=template_bodies[name])
        bcell.alignment = Alignment(wrap_text=True, vertical="top")
        tws.row_dimensions[r].height = 150
    tws.column_dimensions["A"].width = 20
    tws.column_dimensions["B"].width = 90
    tws.row_dimensions[2].height = 60
    tws.freeze_panes = "A3"

    wb.save(OUT)
    resp = sum(1 for r in ordered if r["responded"] == "yes")
    print(f"Wrote {OUT}")
    print(f"  {len(ordered)} recipients ({resp} responded, {len(ordered) - resp} no response)")
    print(f"  Templates available: {', '.join(templates) if templates else '(none)'}")
    print(f"  Auto-queued {auto_count} row(s) inactive {AUTO_FOLLOWUP_DAYS}+ days "
          f"(Follow-up? = Yes, highlighted yellow).")
    print(f"  Auto-template sequence: {' -> '.join(AUTO_TEMPLATE_SEQUENCE)} -> (blank, you choose)")
    print("\n'Follow-ups Sent' + 'Follow-up History' show how many times and when")
    print("each person has been nudged (from followup_sent_log.txt).")
    print("\nTabs: 'Follow-ups' (mark who to nudge) and 'Templates' (edit wording).")
    print("Set Follow-up? = Yes, then pick a Template OR type a Custom Message")
    print("(custom wins if both are set). Edit master wording in the Templates tab.")
    print("SAVE, then run 'Execute Follow-ups'.")


if __name__ == "__main__":
    main()
