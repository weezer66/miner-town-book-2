import csv
import imaplib
import re
from pathlib import Path
import os

SUBJECT = "My new novel, Miner Town: Awakening, is out now"
CAMPAIGN = "launch-2026-08-01"
SINCE_DATE = "01-Aug-2026"
REPORT = Path("gmail_duplicate_audit_report.txt")


def load_smtp_from_bat(path: Path) -> tuple[str, str]:
    if not path.exists():
        return "", ""
    text = path.read_text(encoding="utf-8", errors="ignore")
    user_match = re.search(r'SMTP_USER=([^"\r\n]*)', text)
    pass_match = re.search(r'SMTP_RAW_PASS=([^"\r\n]*)', text)
    if not user_match or not pass_match:
        return "", ""
    smtp_user = user_match.group(1).strip()
    smtp_pass = pass_match.group(1).replace(" ", "").strip()
    return smtp_user, smtp_pass


def get_campaign_emails(log_path: Path, campaign_id: str) -> list[str]:
    raw = log_path.read_text(encoding="utf-8", errors="ignore").replace("\\n", "\n")
    emails = []
    for row in csv.DictReader(raw.splitlines()):
        if (row.get("campaign_id") or "").strip() != campaign_id:
            continue
        email_addr = (row.get("email") or "").strip().lower()
        if email_addr:
            emails.append(email_addr)
    return sorted(set(emails))


def detect_sent_folder(mailboxes: list[bytes]) -> str:
    default_folder = "[Gmail]/Sent Mail"
    for mailbox in mailboxes:
        text = mailbox.decode("utf-8", errors="replace")
        if "\\Sent" not in text:
            continue
        match = re.search(r'"([^"]+)"\s*$', text)
        if match:
            return match.group(1)
    return default_folder


def main() -> None:
    user = os.environ.get("SMTP_USER", "")
    password = os.environ.get("SMTP_PASS", "")
    if not user or not password:
        bat_user, bat_pass = load_smtp_from_bat(Path("Execute Launch Day.bat"))
        user = user or bat_user
        password = password or bat_pass
    if not user or not password:
        raise SystemExit("Missing SMTP credentials (env and Execute Launch Day.bat parse failed)")

    campaign_emails = get_campaign_emails(Path("launch_day_sent_log.txt"), CAMPAIGN)

    client = imaplib.IMAP4_SSL("imap.gmail.com")
    client.login(user, password)

    status, mailboxes = client.list()
    if status != "OK":
        raise SystemExit("IMAP LIST failed")

    sent_folder = detect_sent_folder(mailboxes)
    select_status, _ = client.select(f'"{sent_folder}"')
    if select_status != "OK":
        raise SystemExit(f"IMAP SELECT failed for {sent_folder}")

    counts: dict[str, int] = {}
    for email_addr in campaign_emails:
        search_status, data = client.search(
            None,
            "SINCE",
            f'"{SINCE_DATE}"',
            "SUBJECT",
            f'"{SUBJECT}"',
            "TO",
            f'"{email_addr}"',
        )
        if search_status != "OK":
            counts[email_addr] = -1
            continue
        ids = [token for token in (data[0] or b"").split() if token]
        counts[email_addr] = len(ids)

    client.logout()

    duplicates = sorted([email_addr for email_addr, c in counts.items() if c > 1])
    missing = sorted([email_addr for email_addr, c in counts.items() if c == 0])
    search_errors = sorted([email_addr for email_addr, c in counts.items() if c < 0])
    exactly_one = sum(1 for c in counts.values() if c == 1)

    lines = [
        f"sent_folder={sent_folder}",
        f"campaign={CAMPAIGN}",
        f"campaign_recipients={len(campaign_emails)}",
        f"exactly_one={exactly_one}",
        f"duplicates={len(duplicates)}",
        f"missing={len(missing)}",
        f"search_errors={len(search_errors)}",
    ]

    if duplicates:
        lines.append("duplicate_recipients:")
        lines.extend([f"  {email_addr} x{counts[email_addr]}" for email_addr in duplicates])
    if missing:
        lines.append("missing_recipients:")
        lines.extend([f"  {email_addr}" for email_addr in missing])
    if search_errors:
        lines.append("search_error_recipients:")
        lines.extend([f"  {email_addr}" for email_addr in search_errors])

    REPORT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"WROTE_REPORT={REPORT}")


if __name__ == "__main__":
    main()
