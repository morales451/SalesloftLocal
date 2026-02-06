#!/usr/bin/env python3
"""
outreach.py — CLI Sales Cadence Manager

Manages a 6-touchpoint sales cadence from the command line.
Reads contacts from contacts.csv, prioritizes follow-ups over new leads,
sends emails via Outlook/Office365 SMTP, and prompts for manual tasks.
"""

import csv
import os
import random
import smtplib
import sys
import time
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from getpass import getpass
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("Missing dependency: pandas. Run: pip install pandas")

try:
    from colorama import Fore, Style, init as colorama_init
except ImportError:
    sys.exit("Missing dependency: colorama. Run: pip install colorama")

# ─────────────────────────── CONFIG ───────────────────────────

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
SMTP_EMAIL = os.environ.get("OUTREACH_EMAIL", "")
SMTP_PASSWORD = os.environ.get("OUTREACH_PASSWORD", "")

DAILY_EMAIL_LIMIT = 150
THROTTLE_MIN_SECONDS = 30
THROTTLE_MAX_SECONDS = 90

CSV_PATH = Path(__file__).parent / "contacts.csv"

CSV_COLUMNS = [
    "name", "email", "phone", "company", "title",
    "status", "step", "last_contact_date", "notes",
]

VALID_STATUSES = {"active", "paused", "replied", "not_interested", "finished"}

# Wait‑days before each step becomes due (indexed by step number).
# Step 1 has 0 wait — it's the initial outreach.
STEP_WAIT_DAYS = {
    1: 0,
    2: 2,
    3: 3,
    4: 2,
    5: 2,
    6: 4,
}

STEP_TYPE = {
    1: "email",
    2: "call",
    3: "email",
    4: "sms",
    5: "linkedin",
    6: "email",
}

# ─────────────────────────── EMAIL TEMPLATES ──────────────────

def template_a(name: str, company: str, title: str) -> tuple[str, str]:
    """Step 1 — Initial outreach."""
    subject = f"Quick question for {company}"
    body = (
        f"Hi {name},\n\n"
        f"I came across {company} and was impressed by what your team is doing. "
        f"As {title}, you likely deal with [pain point] on a regular basis.\n\n"
        f"We've helped similar companies solve this — would you be open to a "
        f"brief 15‑minute call this week?\n\n"
        f"Best regards"
    )
    return subject, body


def template_b(name: str, company: str, title: str) -> tuple[str, str]:
    """Step 3 — Follow‑up after call attempt."""
    subject = f"Following up — {company}"
    body = (
        f"Hi {name},\n\n"
        f"I wanted to follow up on my earlier message. I understand things get busy, "
        f"so I'll keep this short.\n\n"
        f"I'd love to share a quick case study showing how we helped a company "
        f"similar to {company} achieve [specific result]. Would that be helpful?\n\n"
        f"Looking forward to hearing from you."
    )
    return subject, body


def template_c(name: str, company: str, title: str) -> tuple[str, str]:
    """Step 6 — Final break‑up email."""
    subject = f"Closing the loop — {company}"
    body = (
        f"Hi {name},\n\n"
        f"I've reached out a few times and haven't heard back, so I'll assume "
        f"the timing isn't right.\n\n"
        f"If things change down the road, feel free to reach out. I'm always happy "
        f"to chat about how we can help {company}.\n\n"
        f"Wishing you all the best."
    )
    return subject, body


TEMPLATE_MAP = {
    1: template_a,
    3: template_b,
    6: template_c,
}

# ─────────────────────────── HELPERS ──────────────────────────

def cprint(msg: str, color: str = Fore.WHITE) -> None:
    """Print with color."""
    print(f"{color}{msg}{Style.RESET_ALL}")


def load_contacts() -> pd.DataFrame:
    """Load CSV into a DataFrame, creating it if missing."""
    if not CSV_PATH.exists():
        df = pd.DataFrame(columns=CSV_COLUMNS)
        df.to_csv(CSV_PATH, index=False)
        return df
    df = pd.read_csv(CSV_PATH, dtype=str).fillna("")
    # Ensure step is int‑like
    df["step"] = pd.to_numeric(df["step"], errors="coerce").fillna(0).astype(int)
    return df


def save_contacts(df: pd.DataFrame) -> None:
    """Crash‑safe save — write to tmp then rename."""
    tmp = CSV_PATH.with_suffix(".csv.tmp")
    df.to_csv(tmp, index=False)
    tmp.replace(CSV_PATH)


def next_business_day(dt: datetime) -> datetime:
    """If dt falls on Sat/Sun, push to Monday."""
    weekday = dt.weekday()
    if weekday == 5:  # Saturday
        return dt + timedelta(days=2)
    if weekday == 6:  # Sunday
        return dt + timedelta(days=1)
    return dt


def is_due(row: pd.Series, today: datetime) -> bool:
    """Return True if the contact's next step is due today or earlier."""
    step = int(row["step"])
    if step < 1 or step > 6:
        return False
    last = row.get("last_contact_date", "")
    if not last and step == 1:
        return True  # brand‑new contact, step 1 is always due
    if not last:
        return False
    try:
        last_dt = datetime.strptime(str(last).strip(), "%Y-%m-%d")
    except ValueError:
        return False
    wait = STEP_WAIT_DAYS.get(step, 0)
    due_date = next_business_day(last_dt + timedelta(days=wait))
    return today >= due_date


def connect_smtp() -> smtplib.SMTP:
    """Connect and authenticate to Outlook SMTP."""
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(SMTP_EMAIL, SMTP_PASSWORD)
    return server


def send_email(server: smtplib.SMTP, to_addr: str, subject: str, body: str) -> None:
    """Send a single plain‑text email."""
    msg = MIMEMultipart()
    msg["From"] = SMTP_EMAIL
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))
    server.sendmail(SMTP_EMAIL, to_addr, msg.as_string())


def throttle() -> None:
    """Random sleep between emails to avoid triggering rate limits."""
    delay = random.randint(THROTTLE_MIN_SECONDS, THROTTLE_MAX_SECONDS)
    cprint(f"  Throttling {delay}s …", Fore.YELLOW)
    time.sleep(delay)


def email_exists(df: pd.DataFrame, email: str) -> bool:
    """Check for duplicate email (case‑insensitive)."""
    return email.strip().lower() in df["email"].str.strip().str.lower().values


# ─────────────────────── CORE ACTIONS ─────────────────────────

def handle_email_step(
    server: smtplib.SMTP,
    df: pd.DataFrame,
    idx: int,
    emails_sent: int,
) -> tuple[pd.DataFrame, int, bool]:
    """Send an email for the given contact row. Returns updated df, count, success."""
    row = df.loc[idx]
    step = int(row["step"])
    template_fn = TEMPLATE_MAP.get(step)
    if template_fn is None:
        cprint(f"  No email template for step {step}; skipping.", Fore.RED)
        return df, emails_sent, False

    subject, body = template_fn(row["name"], row["company"], row["title"])
    try:
        send_email(server, row["email"], subject, body)
    except smtplib.SMTPException as exc:
        cprint(f"  SMTP error for {row['email']}: {exc}", Fore.RED)
        return df, emails_sent, False

    emails_sent += 1
    cprint(
        f"  ✓ Email sent to {row['name']} ({row['email']}) — Step {step}",
        Fore.GREEN,
    )

    # Advance contact
    today_str = datetime.now().strftime("%Y-%m-%d")
    if step >= 6:
        df.at[idx, "status"] = "finished"
        df.at[idx, "notes"] = f"{row['notes']} | Finished cadence {today_str}".strip(" |")
    else:
        df.at[idx, "step"] = step + 1
    df.at[idx, "last_contact_date"] = today_str
    save_contacts(df)
    return df, emails_sent, True


def handle_manual_step(df: pd.DataFrame, idx: int) -> pd.DataFrame:
    """Prompt the user for a manual task (Call / SMS / LinkedIn)."""
    row = df.loc[idx]
    step = int(row["step"])
    task = STEP_TYPE.get(step, "task").upper()

    cprint(f"\n{'─' * 50}", Fore.CYAN)
    cprint(f"  MANUAL TASK: {task}", Fore.CYAN)
    cprint(f"  Name    : {row['name']}", Fore.WHITE)
    cprint(f"  Title   : {row['title']}", Fore.WHITE)
    cprint(f"  Company : {row['company']}", Fore.WHITE)
    cprint(f"  Phone   : {row['phone']}", Fore.WHITE)
    cprint(f"  Notes   : {row['notes']}", Fore.WHITE)
    cprint(f"{'─' * 50}", Fore.CYAN)

    while True:
        result = input(
            f"{Fore.YELLOW}  Result? (y=Done / n=Skip / stop=Remove from list): {Style.RESET_ALL}"
        ).strip().lower()
        if result in ("y", "n", "stop"):
            break
        cprint("  Invalid input. Enter y, n, or stop.", Fore.RED)

    today_str = datetime.now().strftime("%Y-%m-%d")

    if result == "stop":
        df.at[idx, "status"] = "not_interested"
        df.at[idx, "notes"] = f"{row['notes']} | Removed {today_str}".strip(" |")
        cprint(f"  Marked {row['name']} as not_interested.", Fore.MAGENTA)
    elif result == "y":
        if step >= 6:
            df.at[idx, "status"] = "finished"
            df.at[idx, "notes"] = f"{row['notes']} | Finished cadence {today_str}".strip(" |")
        else:
            df.at[idx, "step"] = step + 1
        df.at[idx, "last_contact_date"] = today_str
        cprint(f"  Advanced {row['name']} to step {step + 1}.", Fore.GREEN)
    else:
        cprint(f"  Skipped {row['name']}.", Fore.YELLOW)

    save_contacts(df)
    return df


# ──────────────────── ADD CONTACTS ────────────────────────────

def add_contacts_interactive(df: pd.DataFrame) -> pd.DataFrame:
    """Prompt user to add contacts one at a time."""
    cprint("\n── Add New Contacts (type 'done' to finish) ──\n", Fore.CYAN)
    while True:
        email = input(f"{Fore.YELLOW}  Email (or 'done'): {Style.RESET_ALL}").strip()
        if email.lower() == "done":
            break
        if not email or "@" not in email:
            cprint("  Invalid email.", Fore.RED)
            continue
        if email_exists(df, email):
            cprint(f"  {email} already exists — skipping.", Fore.RED)
            continue

        name = input(f"  Name: ").strip()
        phone = input(f"  Phone: ").strip()
        company = input(f"  Company: ").strip()
        title = input(f"  Title: ").strip()

        new_row = {
            "name": name,
            "email": email,
            "phone": phone,
            "company": company,
            "title": title,
            "status": "active",
            "step": 1,
            "last_contact_date": "",
            "notes": "",
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        save_contacts(df)
        cprint(f"  Added {name} ({email}).", Fore.GREEN)

    return df


def import_contacts_csv(df: pd.DataFrame, path: str) -> pd.DataFrame:
    """Bulk‑import from another CSV file, skipping duplicates."""
    import_path = Path(path)
    if not import_path.exists():
        cprint(f"  File not found: {path}", Fore.RED)
        return df

    incoming = pd.read_csv(import_path, dtype=str).fillna("")
    required = {"name", "email"}
    if not required.issubset(set(incoming.columns)):
        cprint(f"  Import CSV must have at least columns: {required}", Fore.RED)
        return df

    added = 0
    skipped = 0
    for _, row in incoming.iterrows():
        if email_exists(df, row["email"]):
            skipped += 1
            continue
        new_row = {
            "name": row.get("name", ""),
            "email": row["email"],
            "phone": row.get("phone", ""),
            "company": row.get("company", ""),
            "title": row.get("title", ""),
            "status": "active",
            "step": 1,
            "last_contact_date": "",
            "notes": "",
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        added += 1

    save_contacts(df)
    cprint(f"  Imported {added} contacts ({skipped} duplicates skipped).", Fore.GREEN)
    return df


# ──────────────────── REVIEW & RUN ────────────────────────────

def classify_due_contacts(df: pd.DataFrame) -> dict:
    """Split due contacts into follow‑ups and new‑outreach buckets."""
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    today = next_business_day(today)

    buckets: dict[str, list[int]] = {
        "followup_emails": [],  # Steps 3, 6
        "followup_manual": [],  # Steps 2, 4, 5
        "new_emails": [],       # Step 1
    }

    active = df[df["status"] == "active"]
    for idx, row in active.iterrows():
        step = int(row["step"])
        if not is_due(row, today):
            continue
        stype = STEP_TYPE.get(step)
        if stype == "email" and step > 1:
            buckets["followup_emails"].append(idx)
        elif stype != "email" and step > 1:
            buckets["followup_manual"].append(idx)
        elif step == 1:
            buckets["new_emails"].append(idx)

    # Sort so higher steps come first (Step 6 before Step 3)
    buckets["followup_emails"].sort(key=lambda i: -int(df.loc[i, "step"]))
    buckets["followup_manual"].sort(key=lambda i: -int(df.loc[i, "step"]))

    return buckets


def review_and_run(df: pd.DataFrame) -> pd.DataFrame:
    """Main cadence loop: review → confirm → execute."""
    buckets = classify_due_contacts(df)

    fu_email_count = len(buckets["followup_emails"])
    fu_manual_count = len(buckets["followup_manual"])
    new_email_count = len(buckets["new_emails"])
    total = fu_email_count + fu_manual_count + new_email_count

    cprint("\n╔══════════════════════════════════════╗", Fore.CYAN)
    cprint("║       TODAY'S OUTREACH SUMMARY       ║", Fore.CYAN)
    cprint("╠══════════════════════════════════════╣", Fore.CYAN)
    cprint(f"║  Follow‑up Emails  : {fu_email_count:>4}            ║", Fore.WHITE)
    cprint(f"║  Manual Tasks      : {fu_manual_count:>4}            ║", Fore.WHITE)
    cprint(f"║  New Outreach      : {new_email_count:>4}            ║", Fore.WHITE)
    cprint(f"║  Daily Email Limit : {DAILY_EMAIL_LIMIT:>4}            ║", Fore.YELLOW)
    cprint(f"║  Total Actions     : {total:>4}            ║", Fore.WHITE)
    cprint("╚══════════════════════════════════════╝", Fore.CYAN)

    if total == 0:
        cprint("\n  Nothing due today. Check back tomorrow!", Fore.GREEN)
        return df

    confirm = input(
        f"\n{Fore.YELLOW}  Type 'GO' to execute, or anything else to cancel: {Style.RESET_ALL}"
    ).strip()
    if confirm != "GO":
        cprint("  Cancelled.", Fore.RED)
        return df

    # ── Acquire SMTP credentials if emails are due ──
    smtp_server = None
    emails_needed = fu_email_count + new_email_count
    if emails_needed > 0:
        email_addr, password = get_credentials()
        if not email_addr or not password:
            cprint("  Cannot send emails without credentials. Aborting.", Fore.RED)
            return df
        try:
            cprint("\n  Connecting to SMTP …", Fore.YELLOW)
            smtp_server = connect_smtp()
            cprint("  Connected.\n", Fore.GREEN)
        except Exception as exc:
            cprint(f"  SMTP connection failed: {exc}", Fore.RED)
            return df

    emails_sent = 0

    # ── Phase 1: Follow‑up emails (highest priority) ──
    if fu_email_count > 0:
        cprint("\n── FOLLOW‑UP EMAILS ──", Fore.CYAN)
    for idx in buckets["followup_emails"]:
        if emails_sent >= DAILY_EMAIL_LIMIT:
            cprint(
                f"\n  Daily email limit ({DAILY_EMAIL_LIMIT}) reached during follow‑ups. "
                "Stopping all emails for today.",
                Fore.RED,
            )
            break
        df, emails_sent, ok = handle_email_step(smtp_server, df, idx, emails_sent)
        if ok and emails_sent < DAILY_EMAIL_LIMIT:
            throttle()

    # ── Phase 2: Manual follow‑up tasks ──
    if fu_manual_count > 0:
        cprint("\n── MANUAL FOLLOW‑UP TASKS ──", Fore.CYAN)
    for idx in buckets["followup_manual"]:
        df = handle_manual_step(df, idx)

    # ── Phase 3: New outreach emails (lowest priority) ──
    limit_hit = emails_sent >= DAILY_EMAIL_LIMIT
    if limit_hit:
        cprint(
            f"\n  Skipping {new_email_count} new‑outreach emails — daily limit reached.",
            Fore.YELLOW,
        )
    else:
        if new_email_count > 0:
            cprint("\n── NEW OUTREACH EMAILS ──", Fore.CYAN)
        for idx in buckets["new_emails"]:
            if emails_sent >= DAILY_EMAIL_LIMIT:
                remaining = len(buckets["new_emails"]) - buckets["new_emails"].index(idx)
                cprint(
                    f"\n  Daily limit reached. {remaining} new emails deferred to tomorrow.",
                    Fore.YELLOW,
                )
                break
            df, emails_sent, ok = handle_email_step(smtp_server, df, idx, emails_sent)
            if ok and emails_sent < DAILY_EMAIL_LIMIT:
                throttle()

    # ── Cleanup ──
    if smtp_server:
        try:
            smtp_server.quit()
        except Exception:
            pass

    cprint(f"\n  Session complete. Emails sent today: {emails_sent}", Fore.GREEN)
    return df


# ──────────────────── CREDENTIALS ─────────────────────────────

def get_credentials() -> tuple[str, str]:
    """Return (email, password) from env vars or prompt."""
    global SMTP_EMAIL, SMTP_PASSWORD

    if SMTP_EMAIL and SMTP_PASSWORD:
        return SMTP_EMAIL, SMTP_PASSWORD

    cprint("\n  SMTP credentials not found in environment.", Fore.YELLOW)
    cprint("  Set OUTREACH_EMAIL and OUTREACH_PASSWORD env vars, or enter below.\n", Fore.YELLOW)

    if not SMTP_EMAIL:
        SMTP_EMAIL = input("  Outlook email: ").strip()
    if not SMTP_PASSWORD:
        SMTP_PASSWORD = getpass("  Outlook password: ")

    return SMTP_EMAIL, SMTP_PASSWORD


# ──────────────────── STATUS DASHBOARD ────────────────────────

def show_status(df: pd.DataFrame) -> None:
    """Print a quick status overview of all contacts."""
    cprint("\n── CONTACT STATUS ──\n", Fore.CYAN)
    if df.empty:
        cprint("  No contacts loaded.", Fore.YELLOW)
        return

    status_counts = df["status"].value_counts()
    for status, count in status_counts.items():
        cprint(f"  {status:<16} {count}", Fore.WHITE)

    active = df[df["status"] == "active"]
    if not active.empty:
        step_counts = active["step"].value_counts().sort_index()
        cprint("\n  Active contacts by step:", Fore.CYAN)
        for step, count in step_counts.items():
            label = STEP_TYPE.get(int(step), "?").upper()
            cprint(f"    Step {step} ({label}): {count}", Fore.WHITE)

    cprint(f"\n  Total contacts: {len(df)}", Fore.GREEN)


# ──────────────────── MAIN MENU ───────────────────────────────

def main_menu() -> None:
    """Interactive main menu."""
    colorama_init(autoreset=True)

    cprint("\n╔══════════════════════════════════════════╗", Fore.CYAN)
    cprint("║   SALESLOFT LOCAL — CLI Outreach Tool    ║", Fore.CYAN)
    cprint("╚══════════════════════════════════════════╝", Fore.CYAN)

    df = load_contacts()

    while True:
        cprint("\n── MAIN MENU ──", Fore.CYAN)
        cprint("  1. Run today's cadence", Fore.WHITE)
        cprint("  2. Add contacts manually", Fore.WHITE)
        cprint("  3. Import contacts from CSV", Fore.WHITE)
        cprint("  4. View contact status", Fore.WHITE)
        cprint("  5. Exit", Fore.WHITE)

        choice = input(f"\n{Fore.YELLOW}  Select [1-5]: {Style.RESET_ALL}").strip()

        if choice == "1":
            df = review_and_run(df)
        elif choice == "2":
            df = add_contacts_interactive(df)
        elif choice == "3":
            path = input("  Path to CSV file: ").strip()
            df = import_contacts_csv(df, path)
        elif choice == "4":
            show_status(df)
        elif choice == "5":
            cprint("\n  Goodbye.\n", Fore.GREEN)
            break
        else:
            cprint("  Invalid choice.", Fore.RED)


if __name__ == "__main__":
    main_menu()
