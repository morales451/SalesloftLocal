#!/usr/bin/env python3
"""
outreach.py — CLI Sales Cadence Manager

Manages a 6-touchpoint sales cadence from the command line.
Reads contacts from contacts.csv, prioritizes follow-ups over new leads,
sends emails via Outlook/Office365 SMTP, and prompts for manual tasks.
"""

import configparser
import copy
import csv
import json
import os
import random
import re
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
UNDO_FILE = Path(__file__).parent / ".undo_history.json"
CONFIG_FILE = Path(__file__).parent / "config.ini"
TEMPLATES_FILE = Path(__file__).parent / "templates.json"
MAX_UNDO_ENTRIES = 20

CSV_COLUMNS = [
    "name", "email", "phone", "company", "title",
    "status", "step", "last_contact_date", "notes",
    "snooze_until", "created_date", "replied_date", "last_error",
    "meeting_date", "meeting_notes",
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

# ─────────────────────────── CONFIG LOADER ────────────────────

def _write_default_config() -> None:
    """Generate a skeleton config.ini with current defaults."""
    CONFIG_FILE.write_text(
        "[smtp]\n"
        "# Office365 SMTP settings\n"
        "server = smtp.office365.com\n"
        "port = 587\n"
        "# Your Outlook email address\n"
        "email =\n"
        "# Leave password blank — use OUTREACH_PASSWORD env var instead\n"
        "\n"
        "[limits]\n"
        "# Max emails per day\n"
        "daily_email_limit = 150\n"
        "# Random throttle delay between emails (seconds)\n"
        "throttle_min_seconds = 30\n"
        "throttle_max_seconds = 90\n"
        "\n"
        "[cadence]\n"
        "# Days to wait before each step becomes due\n"
        "step1 = 0\n"
        "step2 = 2\n"
        "step3 = 3\n"
        "step4 = 2\n"
        "step5 = 2\n"
        "step6 = 4\n"
    )


def load_config() -> None:
    """Load settings from config.ini, overriding hardcoded defaults."""
    global SMTP_SERVER, SMTP_PORT, SMTP_EMAIL, SMTP_PASSWORD
    global DAILY_EMAIL_LIMIT, THROTTLE_MIN_SECONDS, THROTTLE_MAX_SECONDS
    global STEP_WAIT_DAYS

    if not CONFIG_FILE.exists():
        _write_default_config()
        return

    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_FILE)

    if "smtp" in cfg:
        SMTP_SERVER = cfg["smtp"].get("server", SMTP_SERVER)
        SMTP_PORT = cfg["smtp"].getint("port", SMTP_PORT)
        # Env vars take priority over config file for credentials
        if not SMTP_EMAIL:
            SMTP_EMAIL = cfg["smtp"].get("email", "") or SMTP_EMAIL
        if not SMTP_PASSWORD:
            SMTP_PASSWORD = cfg["smtp"].get("password", "") or SMTP_PASSWORD

    if "limits" in cfg:
        DAILY_EMAIL_LIMIT = cfg["limits"].getint("daily_email_limit", DAILY_EMAIL_LIMIT)
        THROTTLE_MIN_SECONDS = cfg["limits"].getint("throttle_min_seconds", THROTTLE_MIN_SECONDS)
        THROTTLE_MAX_SECONDS = cfg["limits"].getint("throttle_max_seconds", THROTTLE_MAX_SECONDS)

    if "cadence" in cfg:
        for key, val in cfg["cadence"].items():
            if key.startswith("step") and key[4:].isdigit():
                STEP_WAIT_DAYS[int(key[4:])] = int(val)


def show_settings() -> None:
    """Display current active settings."""
    cprint("\n── CURRENT SETTINGS ──\n", Fore.CYAN)
    cprint(f"  Config file  : {CONFIG_FILE}", Fore.WHITE)
    cprint(f"  SMTP server  : {SMTP_SERVER}:{SMTP_PORT}", Fore.WHITE)
    cprint(f"  Email        : {SMTP_EMAIL or '(not set)'}", Fore.WHITE)
    cprint(f"  Password     : {'(set)' if SMTP_PASSWORD else '(not set)'}", Fore.WHITE)
    cprint(f"  Daily limit  : {DAILY_EMAIL_LIMIT} emails", Fore.WHITE)
    cprint(f"  Throttle     : {THROTTLE_MIN_SECONDS}–{THROTTLE_MAX_SECONDS}s", Fore.WHITE)
    cprint(f"\n  Step wait days:", Fore.CYAN)
    for step, days in sorted(STEP_WAIT_DAYS.items()):
        stype = STEP_TYPE.get(step, "?").upper()
        cprint(f"    Step {step} ({stype:<8}): {days} days", Fore.WHITE)
    cprint(f"\n  Edit {CONFIG_FILE.name} to change these settings.", Fore.YELLOW)


# ─────────────────────────── EMAIL TEMPLATES ──────────────────

def template_a(name: str, company: str, title: str) -> tuple[str, str]:
    """Step 1 — Initial outreach."""
    subject = "Henry Roof Coatings vs. Current Supplier"
    body = f"""{name},

I'm Alex with Carlisle, the commercial roofing manufacturer.

Do you offer roof coatings for customers who can't afford full commercial roof replacement?

If you already do, I'd love to compare Henry Roof Coatings with your current supplier—just to keep them honest.

Coatings can help you close more deals and drive more revenue with a budget-friendly option.

Can I stop by your office next week?

Thanks,"""
    return subject, body


def template_b(name: str, company: str, title: str) -> tuple[str, str]:
    """Step 3 — Follow‑up after call attempt."""
    subject = "Re: Henry Roof Coatings vs. Current Supplier"
    body = f"""{name},

Circling back on the note below.

Are you currently walking away from leads that can't get budget approval for a full tear-off?

Henry Roof Coatings can turn those lost bids into profitable projects with much lower labor costs.

I'll be in your area next Tuesday—do you have 5 minutes for me to drop off some info?

Thanks,"""
    return subject, body


def template_c(name: str, company: str, title: str) -> tuple[str, str]:
    """Step 6 — Final break‑up email."""
    subject = "Re: Henry Roof Coatings vs. Current Supplier"
    body = f"""{name},

I haven't heard back, so I'll assume you're all set with your current coating strategy for now.

I won't keep following up, but keep us in mind next time you need a competitive number to keep your current supplier honest.

Feel free to reach out if you have a specific project you need a spec for.

Best,"""
    return subject, body


TEMPLATE_MAP = {
    1: template_a,
    3: template_b,
    6: template_c,
}

# ──────────────────── JSON TEMPLATE SYSTEM ────────────────────

def _default_templates() -> dict:
    """Return hardcoded templates as a JSON-serializable dict (uses <<token>> syntax)."""
    return {
        "1": {
            "subject": "Henry Roof Coatings vs. Current Supplier",
            "body": (
                "<<name>>,\n\n"
                "I'm Alex with Carlisle, the commercial roofing manufacturer.\n\n"
                "Do you offer roof coatings for customers who can't afford full commercial roof replacement?\n\n"
                "If you already do, I'd love to compare Henry Roof Coatings with your current supplier"
                "\u2014just to keep them honest.\n\n"
                "Coatings can help you close more deals and drive more revenue with a budget-friendly option.\n\n"
                "Can I stop by your office next week?\n\nThanks,"
            ),
        },
        "3": {
            "subject": "Re: Henry Roof Coatings vs. Current Supplier",
            "body": (
                "<<name>>,\n\n"
                "Circling back on the note below.\n\n"
                "Are you currently walking away from leads that can't get budget approval for a full tear-off?\n\n"
                "Henry Roof Coatings can turn those lost bids into profitable projects with much lower labor costs.\n\n"
                "I'll be in your area next Tuesday\u2014do you have 5 minutes for me to drop off some info?\n\nThanks,"
            ),
        },
        "6": {
            "subject": "Re: Henry Roof Coatings vs. Current Supplier",
            "body": (
                "<<name>>,\n\n"
                "I haven't heard back, so I'll assume you're all set with your current coating strategy for now.\n\n"
                "I won't keep following up, but keep us in mind next time you need a competitive number "
                "to keep your current supplier honest.\n\n"
                "Feel free to reach out if you have a specific project you need a spec for.\n\nBest,"
            ),
        },
    }


def load_templates() -> dict:
    """Load templates from templates.json, falling back to defaults if missing/corrupt."""
    if TEMPLATES_FILE.exists():
        try:
            return json.loads(TEMPLATES_FILE.read_text())
        except (json.JSONDecodeError, OSError):
            pass
    templates = _default_templates()
    TEMPLATES_FILE.write_text(json.dumps(templates, indent=2))
    return templates


def render_template(tmpl: dict, name: str, company: str, title: str) -> tuple[str, str]:
    """Substitute <<name>>, <<company>>, <<title>> tokens in a template dict."""
    def sub(text: str) -> str:
        return (
            text.replace("<<name>>", name)
                .replace("<<company>>", company)
                .replace("<<title>>", title)
        )
    return sub(tmpl["subject"]), sub(tmpl["body"])


def get_template_for_step(
    step: int, name: str, company: str, title: str
) -> tuple[str, str] | None:
    """Return (subject, body) for a step. Prefers JSON templates, falls back to hardcoded."""
    templates = load_templates()
    tmpl = templates.get(str(step))
    if tmpl:
        return render_template(tmpl, name, company, title)
    fn = TEMPLATE_MAP.get(step)
    if fn:
        return fn(name, company, title)
    return None


def edit_templates() -> None:
    """Interactive editor for email templates stored in templates.json."""
    step_labels = {
        "1": "Step 1 \u2014 Initial outreach (email)",
        "3": "Step 3 \u2014 Follow-up (email)",
        "6": "Step 6 \u2014 Break-up (email)",
    }

    while True:
        templates = load_templates()
        cprint("\n\u2500\u2500 EMAIL TEMPLATE EDITOR \u2500\u2500\n", Fore.CYAN)
        for key, label in step_labels.items():
            tmpl = templates.get(key, {})
            subject_preview = tmpl.get("subject", "(none)")[:55]
            cprint(f"  [{key}] {label}", Fore.WHITE)
            cprint(f"      Subject: {subject_preview}", Fore.WHITE)
        cprint("\n  [r] Reset all to defaults", Fore.YELLOW)
        cprint("  [b] Back", Fore.WHITE)

        choice = input(
            f"\n{Fore.YELLOW}  Select step to edit [1/3/6/r/b]: {Style.RESET_ALL}"
        ).strip().lower()

        if choice == "b":
            break
        elif choice == "r":
            confirm = input("  Type RESET to confirm reset to defaults: ").strip()
            if confirm == "RESET":
                defaults = _default_templates()
                TEMPLATES_FILE.write_text(json.dumps(defaults, indent=2))
                cprint("  Templates reset to defaults.", Fore.GREEN)
        elif choice in step_labels:
            tmpl = templates.get(choice, {})
            cprint(f"\n  Editing: {step_labels[choice]}", Fore.CYAN)
            cprint("  Tokens: <<name>>  <<company>>  <<title>>\n", Fore.WHITE)

            current_subject = tmpl.get("subject", "")
            new_subject = input(f"  Subject [{current_subject}]: ").strip()
            if new_subject:
                tmpl["subject"] = new_subject

            cprint(f"\n  Current body:", Fore.WHITE)
            for line in tmpl.get("body", "").splitlines():
                cprint(f"    {line}", Fore.WHITE)
            cprint(
                "\n  Enter new body line by line. Empty line to finish.",
                Fore.YELLOW,
            )
            cprint("  Press Enter immediately to keep current body.", Fore.YELLOW)

            body_lines: list[str] = []
            first = input("  > ").rstrip("\n")
            if first:
                body_lines.append(first)
                while True:
                    ln = input("  > ").rstrip("\n")
                    if not ln:
                        break
                    body_lines.append(ln)
                tmpl["body"] = "\n".join(body_lines)

            templates[choice] = tmpl
            TEMPLATES_FILE.write_text(json.dumps(templates, indent=2))
            cprint("  Template saved.", Fore.GREEN)
        else:
            cprint("  Invalid choice.", Fore.RED)


# ──────────────────── MANUAL TASK SCRIPTS ─────────────────────

def get_manual_script(step: int, name: str) -> str:
    """Return the script text for a manual task step, with name substituted."""
    scripts = {
        2: f"""Opener: "{name}, this is Alex with Henry, the roof coating manufacturer."
(Pause)
"Do you offer roof coatings to your customers?"

IF YES: "Do you use us... Henry, or someone else?"
  -> "Gotcha. The reason I'm calling is because I am the roof coating rep
     in Texas. I wanted to see if we can get coffee or lunch to discuss
     working together. I think I can help you save money, time, and
     headaches compared to your current supplier."

IF NO: "Would you like to get lunch to learn more about roof coatings so
  that you can offer this to your customers who can't afford a roof
  replacement? This could be a new revenue stream for you to win
  commercial coating jobs." """,

        4: f"""Hey {name}, Alex with Carlisle here. Sent you a few emails about Henry Coatings. Just wanted to float this to the top of your inbox in case you have a project stuck on budget. Let me know if you want to chat. Thanks.""",

        5: f"""{name},

I've been trying to reach you via email and phone regarding Henry Roof Coatings. I'm not trying to spam you—I just know I can help you close the owners who can't afford a full replacement right now. We can offer a 20-year NDL warranty at a fraction of the cost of a tear-off. Is this a better place to chat, or should I try your office again?

Thanks,""",
    }
    return scripts.get(step, "")

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
    # Backward compat: add any missing columns
    for col in CSV_COLUMNS:
        if col not in df.columns:
            df[col] = ""
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
    # Snooze check
    snooze = row.get("snooze_until", "")
    if snooze:
        try:
            snooze_dt = datetime.strptime(str(snooze).strip(), "%Y-%m-%d")
            if today < snooze_dt:
                return False
        except ValueError:
            pass
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


def display_notes_log(notes: str) -> None:
    """Print the notes log in a readable format."""
    if not notes or not str(notes).strip():
        cprint("  Notes   : (none)", Fore.WHITE)
        return
    entries = [e.strip() for e in str(notes).split("|") if e.strip()]
    cprint(f"  {'─' * 40}", Fore.MAGENTA)
    cprint(f"  NOTES LOG ({len(entries)} entries):", Fore.MAGENTA)
    for entry in entries:
        cprint(f"    {entry}", Fore.WHITE)
    cprint(f"  {'─' * 40}", Fore.MAGENTA)


def save_undo(df: pd.DataFrame, idx: int, description: str) -> None:
    """Save a snapshot of a contact row before modifying it."""
    row_data = df.loc[idx].to_dict()
    # Convert numpy types to native Python for JSON
    for k, v in row_data.items():
        if hasattr(v, "item"):
            row_data[k] = v.item()
        else:
            row_data[k] = str(v) if not isinstance(v, (str, int, float)) else v

    entry = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "description": description,
        "idx": int(idx),
        "snapshot": row_data,
    }

    history = []
    if UNDO_FILE.exists():
        try:
            history = json.loads(UNDO_FILE.read_text())
        except (json.JSONDecodeError, OSError):
            history = []
    history.append(entry)
    history = history[-MAX_UNDO_ENTRIES:]
    UNDO_FILE.write_text(json.dumps(history, indent=2))


def _show_undo_history(history: list) -> None:
    """Print available undo steps, most recent first."""
    cprint(f"\n── UNDO HISTORY ({len(history)} step{'s' if len(history) != 1 else ''} available) ──\n", Fore.CYAN)
    for i, entry in enumerate(reversed(history), 1):
        snap = entry["snapshot"]
        cprint(
            f"  [{i}] {entry['timestamp']}  {entry['description']}"
            f"  ({snap.get('name', '?')})",
            Fore.WHITE,
        )


def undo_last_action(df: pd.DataFrame) -> pd.DataFrame:
    """Restore modified contacts to previous states, one step at a time."""
    if not UNDO_FILE.exists():
        cprint("\n  Nothing to undo.", Fore.YELLOW)
        return df

    try:
        history = json.loads(UNDO_FILE.read_text())
    except (json.JSONDecodeError, OSError):
        cprint("\n  Undo history is corrupted.", Fore.RED)
        return df

    if not history:
        cprint("\n  Nothing to undo.", Fore.YELLOW)
        return df

    while history:
        _show_undo_history(history)
        entry = history[-1]
        snap = entry["snapshot"]

        cprint(f"\n── NEXT UNDO ──", Fore.CYAN)
        cprint(f"  Action : {entry['description']}", Fore.WHITE)
        cprint(f"  Time   : {entry['timestamp']}", Fore.WHITE)
        cprint(f"  Contact: {snap.get('name', '?')} ({snap.get('email', '?')})", Fore.WHITE)
        cprint(
            f"  Restores to: Step {snap.get('step', '?')}, Status: {snap.get('status', '?')}",
            Fore.WHITE,
        )

        confirm = input(
            f"\n{Fore.YELLOW}  Type UNDO to restore, or anything else to stop: {Style.RESET_ALL}"
        ).strip()
        if confirm != "UNDO":
            cprint("  Stopped.", Fore.YELLOW)
            break

        # Find contact by email (index may have shifted after deletes)
        email = snap.get("email", "")
        matches = df[df["email"].str.strip().str.lower() == email.strip().lower()]
        if matches.empty:
            cprint(f"  Contact {email} not found in CSV. May have been deleted.", Fore.RED)
            history.pop()
            UNDO_FILE.write_text(json.dumps(history, indent=2))
            break

        actual_idx = matches.index[0]
        for col, val in snap.items():
            if col in df.columns:
                df.at[actual_idx, col] = val
        df["step"] = pd.to_numeric(df["step"], errors="coerce").fillna(0).astype(int)

        save_contacts(df)
        history.pop()
        UNDO_FILE.write_text(json.dumps(history, indent=2))
        cprint(
            f"  Restored {snap['name']} to Step {snap['step']}, Status: {snap['status']}.",
            Fore.GREEN,
        )

        if not history:
            cprint("\n  No more undo steps available.", Fore.YELLOW)
            break

    return df


def prompt_for_notes(df: pd.DataFrame, idx: int, step: int) -> pd.DataFrame:
    """Ask the user for notes after an action. Appends as a timestamped log entry."""
    note = input(
        f"{Fore.YELLOW}  Add notes (or press Enter to skip): {Style.RESET_ALL}"
    ).strip()
    if note:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
        existing = str(df.at[idx, "notes"]).strip()
        entry = f"[{timestamp} S{step}] {note}"
        if existing:
            df.at[idx, "notes"] = f"{existing} | {entry}"
        else:
            df.at[idx, "notes"] = entry
        save_contacts(df)
        cprint(f"  Note saved.", Fore.GREEN)
    return df


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
    display_notes_log(row["notes"])
    result = get_template_for_step(step, row["name"], row["company"], row["title"])
    if result is None:
        cprint(f"  No email template for step {step}; skipping.", Fore.RED)
        return df, emails_sent, False

    subject, body = result

    # Preview email before sending
    cprint(f"\n  To: {row['email']}", Fore.GREEN)
    cprint(f"  Subject: {subject}", Fore.GREEN)
    cprint(f"  {'─' * 40}", Fore.WHITE)
    for line in body.splitlines():
        cprint(f"  {line}", Fore.WHITE)
    cprint(f"  {'─' * 40}", Fore.WHITE)

    confirm = input(f"{Fore.YELLOW}  Send this email? (y/n): {Style.RESET_ALL}").strip().lower()
    if confirm != "y":
        cprint(f"  Skipped {row['name']}.", Fore.YELLOW)
        return df, emails_sent, False

    try:
        send_email(server, row["email"], subject, body)
    except smtplib.SMTPException as exc:
        cprint(f"  SMTP error for {row['email']}: {exc}", Fore.RED)
        df.at[idx, "last_error"] = f"{datetime.now().strftime('%Y-%m-%d %H:%M')} {exc}"
        save_contacts(df)
        return df, emails_sent, False

    emails_sent += 1
    df.at[idx, "last_error"] = ""
    cprint(
        f"  ✓ Email sent to {row['name']} ({row['email']}) — Step {step}",
        Fore.GREEN,
    )

    # Save undo snapshot before advancing
    save_undo(df, idx, f"Email Step {step} to {row['name']}")

    # Advance contact
    today_str = datetime.now().strftime("%Y-%m-%d")
    if step >= 6:
        df.at[idx, "status"] = "finished"
        df.at[idx, "notes"] = f"{row['notes']} | Finished cadence {today_str}".strip(" |")
    else:
        df.at[idx, "step"] = step + 1
    df.at[idx, "last_contact_date"] = today_str
    save_contacts(df)
    df = prompt_for_notes(df, idx, step)
    return df, emails_sent, True


def handle_manual_step(df: pd.DataFrame, idx: int) -> pd.DataFrame:
    """Prompt the user for a manual task (Call / SMS / LinkedIn)."""
    row = df.loc[idx]
    step = int(row["step"])
    task = STEP_TYPE.get(step, "task").upper()

    cprint(f"\n{'─' * 60}", Fore.CYAN)
    cprint(f"  MANUAL TASK: {task}", Fore.CYAN)
    cprint(f"  Name    : {row['name']}", Fore.WHITE)
    cprint(f"  Title   : {row['title']}", Fore.WHITE)
    cprint(f"  Company : {row['company']}", Fore.WHITE)
    cprint(f"  Phone   : {row['phone']}", Fore.WHITE)
    cprint(f"{'─' * 60}", Fore.CYAN)
    display_notes_log(row["notes"])

    # Display the script for this manual step
    script = get_manual_script(step, row["name"])
    if script:
        cprint(f"\n  {'═' * 56}", Fore.GREEN)
        cprint(f"  SCRIPT:", Fore.GREEN)
        cprint(f"  {'═' * 56}", Fore.GREEN)
        for line in script.strip().splitlines():
            cprint(f"  {line}", Fore.WHITE)
        cprint(f"  {'═' * 56}\n", Fore.GREEN)

    while True:
        result = input(
            f"{Fore.YELLOW}  Result? (y=Done / n=Skip / mtg=Meeting booked / stop=Remove): {Style.RESET_ALL}"
        ).strip().lower()
        if result in ("y", "n", "stop", "mtg"):
            break
        cprint("  Invalid input. Enter y, n, mtg, or stop.", Fore.RED)

    # Save undo snapshot before modifying
    save_undo(df, idx, f"Manual Step {step} ({task}) for {row['name']}")

    today_str = datetime.now().strftime("%Y-%m-%d")

    if result == "mtg":
        mtg_date = input(f"{Fore.YELLOW}  Meeting date (YYYY-MM-DD): {Style.RESET_ALL}").strip()
        mtg_note = input(f"{Fore.YELLOW}  Meeting details (location/time/topic): {Style.RESET_ALL}").strip()
        df.at[idx, "status"] = "replied"
        df.at[idx, "replied_date"] = today_str
        df.at[idx, "meeting_date"] = mtg_date
        df.at[idx, "meeting_notes"] = mtg_note
        df.at[idx, "last_contact_date"] = today_str
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
        mtg_entry = f"[{timestamp} S{step}] MEETING BOOKED {mtg_date}: {mtg_note}"
        existing = str(row["notes"]).strip()
        df.at[idx, "notes"] = f"{existing} | {mtg_entry}".strip(" |")
        cprint(f"  Meeting booked with {row['name']} on {mtg_date}!", Fore.GREEN)
    elif result == "stop":
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

    if result != "mtg":
        df = prompt_for_notes(df, idx, step)
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
            "created_date": datetime.now().strftime("%Y-%m-%d"),
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
            "created_date": datetime.now().strftime("%Y-%m-%d"),
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


def preview_email_step(df: pd.DataFrame, idx: int) -> None:
    """Dry‑run: render and display an email without sending."""
    row = df.loc[idx]
    step = int(row["step"])
    display_notes_log(row["notes"])
    result = get_template_for_step(step, row["name"], row["company"], row["title"])
    if result is None:
        cprint(f"  No template for step {step}.", Fore.RED)
        return
    subject, body = result
    cprint(f"\n  To: {row['email']}", Fore.GREEN)
    cprint(f"  Subject: {subject}", Fore.GREEN)
    cprint(f"  {'─' * 40}", Fore.WHITE)
    for line in body.splitlines():
        cprint(f"  {line}", Fore.WHITE)
    cprint(f"  {'─' * 40}", Fore.WHITE)
    cprint(f"  [DRY RUN] Would send email.", Fore.YELLOW)


def preview_manual_step(df: pd.DataFrame, idx: int) -> None:
    """Dry‑run: display a manual task without prompting."""
    row = df.loc[idx]
    step = int(row["step"])
    task = STEP_TYPE.get(step, "task").upper()
    cprint(f"  {'─' * 60}", Fore.CYAN)
    cprint(f"  MANUAL TASK: {task}", Fore.CYAN)
    cprint(f"  Name    : {row['name']}", Fore.WHITE)
    cprint(f"  Title   : {row['title']}", Fore.WHITE)
    cprint(f"  Company : {row['company']}", Fore.WHITE)
    cprint(f"  Phone   : {row['phone']}", Fore.WHITE)
    cprint(f"  {'─' * 60}", Fore.CYAN)
    display_notes_log(row["notes"])
    script = get_manual_script(step, row["name"])
    if script:
        cprint(f"\n  {'═' * 56}", Fore.GREEN)
        cprint(f"  SCRIPT:", Fore.GREEN)
        cprint(f"  {'═' * 56}", Fore.GREEN)
        for line in script.strip().splitlines():
            cprint(f"  {line}", Fore.WHITE)
        cprint(f"  {'═' * 56}", Fore.GREEN)
    cprint(f"  [DRY RUN] Would prompt for manual action.", Fore.YELLOW)


def review_and_run(df: pd.DataFrame, dry_run: bool = False) -> pd.DataFrame:
    """Main cadence loop: review → confirm → execute."""
    buckets = classify_due_contacts(df)

    fu_email_count = len(buckets["followup_emails"])
    fu_manual_count = len(buckets["followup_manual"])
    new_email_count = len(buckets["new_emails"])
    total = fu_email_count + fu_manual_count + new_email_count

    mode_label = "DRY RUN PREVIEW" if dry_run else "TODAY'S OUTREACH SUMMARY"
    cprint(f"\n╔══════════════════════════════════════╗", Fore.CYAN)
    cprint(f"║  {mode_label:^36}║", Fore.CYAN)
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

    # ── Dry‑run path ──
    if dry_run:
        cprint("\n── DRY RUN: Previewing all actions (nothing will be sent or saved) ──\n", Fore.YELLOW)
        action_num = 0
        if fu_email_count > 0:
            cprint("── FOLLOW‑UP EMAILS ──", Fore.CYAN)
        for idx in buckets["followup_emails"]:
            action_num += 1
            row = df.loc[idx]
            cprint(f"\n  [{action_num}/{total}] {row['name']} — {row['company']} (Step {int(row['step'])})", Fore.CYAN)
            preview_email_step(df, idx)
        if fu_manual_count > 0:
            cprint("\n── MANUAL FOLLOW‑UP TASKS ──", Fore.CYAN)
        for idx in buckets["followup_manual"]:
            action_num += 1
            row = df.loc[idx]
            cprint(f"\n  [{action_num}/{total}] {row['name']} — {row['company']} (Step {int(row['step'])})", Fore.CYAN)
            preview_manual_step(df, idx)
        if new_email_count > 0:
            cprint("\n── NEW OUTREACH EMAILS ──", Fore.CYAN)
        for idx in buckets["new_emails"]:
            action_num += 1
            row = df.loc[idx]
            cprint(f"\n  [{action_num}/{total}] {row['name']} — {row['company']} (Step {int(row['step'])})", Fore.CYAN)
            preview_email_step(df, idx)
        cprint(f"\n  Dry run complete. {total} actions previewed.", Fore.GREEN)
        return df

    # ── Live run ──
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
    action_num = 0

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
        action_num += 1
        row = df.loc[idx]
        cprint(f"\n  [{action_num}/{total}] {row['name']} — {row['company']} (Step {int(row['step'])})", Fore.CYAN)
        df, emails_sent, ok = handle_email_step(smtp_server, df, idx, emails_sent)
        if ok and emails_sent < DAILY_EMAIL_LIMIT:
            throttle()

    # ── Phase 2: Manual follow‑up tasks ──
    if fu_manual_count > 0:
        cprint("\n── MANUAL FOLLOW‑UP TASKS ──", Fore.CYAN)
    for idx in buckets["followup_manual"]:
        action_num += 1
        row = df.loc[idx]
        cprint(f"\n  [{action_num}/{total}] {row['name']} — {row['company']} (Step {int(row['step'])})", Fore.CYAN)
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
            action_num += 1
            row = df.loc[idx]
            cprint(f"\n  [{action_num}/{total}] {row['name']} — {row['company']} (Step {int(row['step'])})", Fore.CYAN)
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


# ──────────────────── WHO'S DUE TODAY ─────────────────────────

def show_due_today(df: pd.DataFrame) -> None:
    """Print a compact table of contacts whose next step is due."""
    buckets = classify_due_contacts(df)
    all_due = (
        buckets["followup_emails"]
        + buckets["followup_manual"]
        + buckets["new_emails"]
    )
    if not all_due:
        cprint("\n  Nothing due today!", Fore.GREEN)
        return

    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    cprint("\n── WHO'S DUE TODAY ──\n", Fore.CYAN)
    cprint(f"  {'#':<3} {'Name':<20} {'Company':<20} {'Step':<6} {'Type':<10} {'Overdue'}", Fore.WHITE)
    cprint(f"  {'─' * 75}", Fore.WHITE)

    for i, idx in enumerate(all_due, 1):
        row = df.loc[idx]
        step = int(row["step"])
        stype = STEP_TYPE.get(step, "?").upper()
        last = row.get("last_contact_date", "")
        if last:
            try:
                last_dt = datetime.strptime(str(last).strip(), "%Y-%m-%d")
                due_date = next_business_day(last_dt + timedelta(days=STEP_WAIT_DAYS.get(step, 0)))
                overdue = (today - due_date).days
            except ValueError:
                overdue = 0
        else:
            overdue = 0
        overdue_str = f"+{overdue}d" if overdue > 0 else "today"
        color = Fore.RED if overdue > 3 else (Fore.YELLOW if overdue > 0 else Fore.GREEN)
        cprint(f"  {i:<3} {row['name']:<20} {row['company']:<20} {step:<6} {stype:<10} {overdue_str}", color)

    cprint(f"\n  {len(all_due)} contacts due.", Fore.GREEN)


# ──────────────────── CONTACT MANAGER ─────────────────────────

def manage_contacts(df: pd.DataFrame) -> pd.DataFrame:
    """Interactive contact manager with search/list/edit/delete/snooze."""
    cprint("\n── CONTACT MANAGER ──\n", Fore.CYAN)
    cprint("  Commands: search <term> | list [status|stepN] | edit <#> | delete <#> | snooze <#> | back", Fore.WHITE)

    while True:
        cmd = input(f"\n{Fore.YELLOW}  contacts> {Style.RESET_ALL}").strip()
        if not cmd or cmd.lower() == "back":
            break
        parts = cmd.split(maxsplit=1)
        action = parts[0].lower()
        arg = parts[1] if len(parts) > 1 else ""

        if action == "search":
            _contact_search(df, arg)
        elif action == "list":
            _contact_list(df, arg)
        elif action == "edit":
            df = _contact_edit(df, arg)
        elif action == "delete":
            df = _contact_delete(df, arg)
        elif action == "snooze":
            df = _contact_snooze(df, arg)
        else:
            cprint("  Unknown command. Try: search, list, edit, delete, snooze, back", Fore.RED)
    return df


def _print_contact_table(filtered: pd.DataFrame, page_size: int = 20) -> None:
    """Print a paginated table of contacts."""
    if filtered.empty:
        cprint("  No contacts match.", Fore.YELLOW)
        return

    total = len(filtered)
    total_pages = max(1, (total + page_size - 1) // page_size)
    page = 0

    while True:
        start = page * page_size
        chunk = filtered.iloc[start:start + page_size]

        cprint(f"\n  {'#':<4} {'Name':<20} {'Company':<18} {'Status':<15} {'Step':<5} {'Email'}", Fore.WHITE)
        cprint(f"  {'─' * 80}", Fore.WHITE)
        for idx, row in chunk.iterrows():
            color = Fore.GREEN if row["status"] == "active" else Fore.WHITE
            cprint(
                f"  {idx:<4} {row['name']:<20} {row['company']:<18} "
                f"{row['status']:<15} {int(row['step']):<5} {row['email']}",
                color,
            )

        if total_pages > 1:
            cprint(
                f"\n  Page {page + 1}/{total_pages} — {total} contacts total",
                Fore.CYAN,
            )
            nav = input("  [n]ext / [p]rev / [q]uit: ").strip().lower()
            if nav == "n" and page < total_pages - 1:
                page += 1
            elif nav == "p" and page > 0:
                page -= 1
            else:
                break
        else:
            cprint(f"\n  {total} contact{'s' if total != 1 else ''}", Fore.CYAN)
            break


def _contact_search(df: pd.DataFrame, term: str) -> None:
    """Search contacts by name, company, or email."""
    if not term:
        cprint("  Usage: search <term>", Fore.RED)
        return
    mask = (
        df["name"].str.contains(term, case=False, na=False)
        | df["company"].str.contains(term, case=False, na=False)
        | df["email"].str.contains(term, case=False, na=False)
    )
    _print_contact_table(df[mask])


def _contact_list(df: pd.DataFrame, filter_arg: str) -> None:
    """List contacts, optionally filtered by status or step."""
    if not filter_arg:
        _print_contact_table(df)
        return
    arg = filter_arg.strip().lower()
    if arg in VALID_STATUSES:
        _print_contact_table(df[df["status"] == arg])
    elif arg.startswith("step") and arg[4:].isdigit():
        _print_contact_table(df[df["step"] == int(arg[4:])])
    else:
        _print_contact_table(df[df["company"].str.contains(filter_arg, case=False, na=False)])


def _contact_edit(df: pd.DataFrame, idx_str: str) -> pd.DataFrame:
    """Edit a contact's fields."""
    try:
        idx = int(idx_str)
    except (ValueError, TypeError):
        cprint("  Usage: edit <row number>", Fore.RED)
        return df
    if idx not in df.index:
        cprint(f"  No contact at index {idx}.", Fore.RED)
        return df

    row = df.loc[idx]
    cprint(f"\n  Editing: {row['name']} ({row['email']})", Fore.CYAN)
    cprint("  Press Enter to keep current value.\n", Fore.WHITE)

    editable = ["name", "email", "phone", "company", "title", "status"]
    for field in editable:
        current = str(row[field])
        new_val = input(f"  {field} [{current}]: ").strip()
        if new_val:
            if field == "status" and new_val not in VALID_STATUSES:
                cprint(f"  Invalid status. Valid: {VALID_STATUSES}", Fore.RED)
                continue
            if field == "status" and new_val == "replied" and row["status"] != "replied":
                df.at[idx, "replied_date"] = datetime.now().strftime("%Y-%m-%d")
            df.at[idx, field] = new_val

    save_contacts(df)
    cprint(f"  Contact updated.", Fore.GREEN)
    return df


def _contact_delete(df: pd.DataFrame, idx_str: str) -> pd.DataFrame:
    """Delete a contact after confirmation."""
    try:
        idx = int(idx_str)
    except (ValueError, TypeError):
        cprint("  Usage: delete <row number>", Fore.RED)
        return df
    if idx not in df.index:
        cprint(f"  No contact at index {idx}.", Fore.RED)
        return df

    name = df.at[idx, "name"]
    confirm = input(f"  Delete {name}? Type DELETE to confirm: ").strip()
    if confirm == "DELETE":
        df = df.drop(idx).reset_index(drop=True)
        save_contacts(df)
        cprint(f"  {name} deleted.", Fore.GREEN)
    else:
        cprint("  Cancelled.", Fore.YELLOW)
    return df


def _contact_snooze(df: pd.DataFrame, idx_str: str) -> pd.DataFrame:
    """Snooze a contact until a specific date."""
    try:
        idx = int(idx_str)
    except (ValueError, TypeError):
        cprint("  Usage: snooze <row number>", Fore.RED)
        return df
    if idx not in df.index:
        cprint(f"  No contact at index {idx}.", Fore.RED)
        return df

    row = df.loc[idx]
    current_snooze = row.get("snooze_until", "")
    if current_snooze:
        cprint(f"  Currently snoozed until: {current_snooze}", Fore.YELLOW)

    date_str = input("  Snooze until (YYYY-MM-DD, or 'clear'): ").strip()
    if date_str.lower() == "clear":
        df.at[idx, "snooze_until"] = ""
        save_contacts(df)
        cprint(f"  Snooze cleared for {row['name']}.", Fore.GREEN)
    else:
        try:
            datetime.strptime(date_str, "%Y-%m-%d")
            df.at[idx, "snooze_until"] = date_str
            save_contacts(df)
            cprint(f"  {row['name']} snoozed until {date_str}.", Fore.GREEN)
        except ValueError:
            cprint("  Invalid date format. Use YYYY-MM-DD.", Fore.RED)
    return df


# ──────────────────── RETRY FAILED EMAILS ─────────────────────

def retry_failed(df: pd.DataFrame) -> pd.DataFrame:
    """Re‑attempt all emails that previously failed."""
    failed = df[df["last_error"].astype(str).str.strip() != ""]
    if failed.empty:
        cprint("\n  No failed emails to retry.", Fore.GREEN)
        return df

    cprint(f"\n── RETRY FAILED EMAILS ({len(failed)} contacts) ──\n", Fore.CYAN)
    for idx, row in failed.iterrows():
        cprint(f"  {row['name']} ({row['email']}) — Error: {row['last_error']}", Fore.RED)

    confirm = input(f"\n{Fore.YELLOW}  Type 'GO' to retry all, or anything else to cancel: {Style.RESET_ALL}").strip()
    if confirm != "GO":
        cprint("  Cancelled.", Fore.YELLOW)
        return df

    email_addr, password = get_credentials()
    if not email_addr or not password:
        cprint("  Cannot send without credentials.", Fore.RED)
        return df

    try:
        cprint("\n  Connecting to SMTP …", Fore.YELLOW)
        smtp_server = connect_smtp()
        cprint("  Connected.\n", Fore.GREEN)
    except Exception as exc:
        cprint(f"  SMTP connection failed: {exc}", Fore.RED)
        return df

    emails_sent = 0
    for idx in failed.index:
        df, emails_sent, ok = handle_email_step(smtp_server, df, idx, emails_sent)
        if ok:
            throttle()

    try:
        smtp_server.quit()
    except Exception:
        pass

    cprint(f"\n  Retry complete. {emails_sent} emails sent.", Fore.GREEN)
    return df


# ──────────────────── ANALYTICS & REPORTS ─────────────────────

def show_funnel(df: pd.DataFrame) -> None:
    """Display a horizontal bar chart of active contacts by step."""
    active = df[df["status"] == "active"]
    if active.empty:
        cprint("  No active contacts.", Fore.YELLOW)
        return

    cprint("\n── CADENCE FUNNEL ──\n", Fore.CYAN)
    max_bar = 30
    max_count = active["step"].value_counts().max() if not active.empty else 1

    for step in range(1, 7):
        count = len(active[active["step"] == step])
        bar_len = int((count / max_count) * max_bar) if max_count > 0 else 0
        bar = "\u2588" * bar_len
        stype = STEP_TYPE.get(step, "?").upper()
        label = f"  Step {step} ({stype:<8})"
        cprint(f"{label} {bar} {count}", Fore.GREEN if count > 0 else Fore.WHITE)

    cprint(f"\n  ── Outcomes ──", Fore.CYAN)
    for status in ["replied", "not_interested", "finished", "paused"]:
        count = len(df[df["status"] == status])
        if count > 0:
            color = Fore.GREEN if status == "replied" else Fore.WHITE
            cprint(f"  {status:<16} {count}", color)


def show_activity_stats(df: pd.DataFrame) -> None:
    """Parse notes timestamps to compute daily/weekly activity."""
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    week_ago = today - timedelta(days=7)

    today_actions = 0
    week_actions = 0
    step_counts = {s: 0 for s in range(1, 7)}

    pattern = re.compile(r"\[(\d{4}-\d{2}-\d{2}) \d{2}:\d{2} S(\d)\]")
    for _, row in df.iterrows():
        notes = str(row.get("notes", ""))
        for match in pattern.finditer(notes):
            try:
                dt = datetime.strptime(match.group(1), "%Y-%m-%d")
                step_num = int(match.group(2))
            except ValueError:
                continue
            if dt >= today:
                today_actions += 1
            if dt >= week_ago:
                week_actions += 1
                if step_num in step_counts:
                    step_counts[step_num] += 1

    cprint("\n── ACTIVITY STATS ──\n", Fore.CYAN)
    cprint(f"  Actions today:     {today_actions}", Fore.WHITE)
    cprint(f"  Actions this week: {week_actions}", Fore.WHITE)
    cprint(f"\n  This week by step:", Fore.CYAN)
    for step in range(1, 7):
        stype = STEP_TYPE.get(step, "?").upper()
        cprint(f"    Step {step} ({stype:<8}): {step_counts[step]}", Fore.WHITE)


def show_conversion_rates(df: pd.DataFrame) -> None:
    """Compute and display conversion metrics."""
    total = len(df)
    if total == 0:
        cprint("  No contacts to analyze.", Fore.YELLOW)
        return

    replied = len(df[df["status"] == "replied"])
    finished = len(df[df["status"] == "finished"])
    not_int = len(df[df["status"] == "not_interested"])
    active = len(df[df["status"] == "active"])
    paused = len(df[df["status"] == "paused"])

    cprint("\n── CONVERSION RATES ──\n", Fore.CYAN)
    cprint(f"  Total contacts:    {total}", Fore.WHITE)
    cprint(f"  Active:            {active} ({active / total * 100:.0f}%)", Fore.GREEN)
    cprint(f"  Replied:           {replied} ({replied / total * 100:.0f}%)", Fore.GREEN)
    cprint(f"  Finished cadence:  {finished} ({finished / total * 100:.0f}%)", Fore.WHITE)
    cprint(f"  Not interested:    {not_int} ({not_int / total * 100:.0f}%)", Fore.WHITE)
    cprint(f"  Paused:            {paused} ({paused / total * 100:.0f}%)", Fore.YELLOW)

    replied_df = df[df["status"] == "replied"]
    if not replied_df.empty:
        avg_step = replied_df["step"].astype(int).mean()
        cprint(f"\n  Avg step at reply: {avg_step:.1f}", Fore.CYAN)


def analytics_menu(df: pd.DataFrame) -> None:
    """Analytics submenu."""
    while True:
        cprint("\n── ANALYTICS & REPORTS ──", Fore.CYAN)
        cprint("  1. Cadence funnel", Fore.WHITE)
        cprint("  2. Activity stats (today/week)", Fore.WHITE)
        cprint("  3. Conversion rates", Fore.WHITE)
        cprint("  4. Meetings booked", Fore.WHITE)
        cprint("  5. Back to main menu", Fore.WHITE)

        choice = input(f"\n{Fore.YELLOW}  Select [1-5]: {Style.RESET_ALL}").strip()
        if choice == "1":
            show_funnel(df)
        elif choice == "2":
            show_activity_stats(df)
        elif choice == "3":
            show_conversion_rates(df)
        elif choice == "4":
            show_meetings(df)
        elif choice == "5":
            break
        else:
            cprint("  Invalid choice.", Fore.RED)


# ──────────────────── MEETINGS DASHBOARD ──────────────────────

def show_meetings(df: pd.DataFrame) -> None:
    """Display all contacts with meetings booked."""
    has_meeting = df[df["meeting_date"].astype(str).str.strip() != ""]
    if has_meeting.empty:
        cprint("\n  No meetings booked yet.", Fore.YELLOW)
        return

    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    cprint("\n── MEETINGS BOOKED ──\n", Fore.CYAN)
    cprint(f"  {'Name':<20} {'Company':<20} {'Date':<12} {'Details'}", Fore.WHITE)
    cprint(f"  {'─' * 75}", Fore.WHITE)

    for _, row in has_meeting.iterrows():
        mtg_date = str(row["meeting_date"]).strip()
        try:
            dt = datetime.strptime(mtg_date, "%Y-%m-%d")
            if dt < today:
                color = Fore.WHITE  # past
            elif dt == today:
                color = Fore.GREEN  # today
            else:
                color = Fore.CYAN   # upcoming
        except ValueError:
            color = Fore.WHITE
        cprint(
            f"  {row['name']:<20} {row['company']:<20} {mtg_date:<12} {row['meeting_notes']}",
            color,
        )

    upcoming = 0
    for _, row in has_meeting.iterrows():
        try:
            if datetime.strptime(str(row["meeting_date"]).strip(), "%Y-%m-%d") >= today:
                upcoming += 1
        except ValueError:
            pass
    cprint(f"\n  {len(has_meeting)} total meetings ({upcoming} upcoming)", Fore.GREEN)


# ──────────────────── DAILY RECAP ─────────────────────────────

def daily_recap(df: pd.DataFrame) -> None:
    """Generate and optionally email a summary of today's activity."""
    today_str = datetime.now().strftime("%Y-%m-%d")
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    # Parse notes for today's actions
    pattern = re.compile(r"\[" + today_str + r" \d{2}:\d{2} S(\d)\] (.+?)(?:\||$)")
    actions = []
    for _, row in df.iterrows():
        notes = str(row.get("notes", ""))
        for match in pattern.finditer(notes):
            actions.append({
                "name": row["name"],
                "company": row["company"],
                "step": int(match.group(1)),
                "action": match.group(2).strip(),
            })

    # Contacts contacted today
    contacted_today = df[df["last_contact_date"] == today_str]

    # Meetings booked
    meetings_today = df[df["meeting_date"] == today_str]

    # Build recap
    lines = []
    lines.append(f"═══ DAILY RECAP — {today_str} ═══")
    lines.append("")
    lines.append(f"Contacts worked today: {len(contacted_today)}")
    lines.append(f"Actions logged:        {len(actions)}")
    lines.append(f"Meetings booked today: {len(meetings_today)}")
    lines.append("")

    if actions:
        lines.append("── Actions ──")
        for a in actions:
            stype = STEP_TYPE.get(a["step"], "?").upper()
            lines.append(f"  {a['name']} ({a['company']}) — S{a['step']} {stype}: {a['action']}")
        lines.append("")

    if not meetings_today.empty:
        lines.append("── Meetings Booked ──")
        for _, row in meetings_today.iterrows():
            lines.append(f"  {row['name']} ({row['company']}) — {row['meeting_notes']}")
        lines.append("")

    # Due tomorrow
    tomorrow = next_business_day(today + timedelta(days=1))
    due_tomorrow = 0
    active = df[df["status"] == "active"]
    for _, row in active.iterrows():
        if is_due(row, tomorrow):
            due_tomorrow += 1
    lines.append(f"Due next business day: {due_tomorrow} contacts")

    # Status summary
    lines.append("")
    lines.append("── Pipeline ──")
    for status in ["active", "replied", "finished", "not_interested", "paused"]:
        count = len(df[df["status"] == status])
        if count > 0:
            lines.append(f"  {status:<16} {count}")

    recap_text = "\n".join(lines)

    # Print to terminal
    cprint(f"\n", Fore.CYAN)
    for line in lines:
        cprint(f"  {line}", Fore.WHITE)
    cprint("", Fore.CYAN)

    # Offer to email it
    send_it = input(
        f"\n{Fore.YELLOW}  Email this recap to yourself? (y/n): {Style.RESET_ALL}"
    ).strip().lower()
    if send_it == "y":
        email_addr, password = get_credentials()
        if not email_addr or not password:
            cprint("  No credentials available.", Fore.RED)
            return
        try:
            server = connect_smtp()
            send_email(server, email_addr, f"Outreach Recap — {today_str}", recap_text)
            server.quit()
            cprint(f"  Recap sent to {email_addr}.", Fore.GREEN)
        except Exception as exc:
            cprint(f"  Failed to send recap: {exc}", Fore.RED)


# ──────────────────── EXPORT ──────────────────────────────────

def export_report(df: pd.DataFrame) -> None:
    """Export contacts to a clean CSV report."""
    today_str = datetime.now().strftime("%Y-%m-%d")
    default_name = f"outreach_export_{today_str}.csv"

    cprint("\n── EXPORT CONTACTS ──\n", Fore.CYAN)
    cprint("  1. All contacts (full data)", Fore.WHITE)
    cprint("  2. Active contacts only", Fore.WHITE)
    cprint("  3. Meetings booked", Fore.WHITE)
    cprint("  4. Replied contacts", Fore.WHITE)
    cprint("  5. Back", Fore.WHITE)

    choice = input(f"\n{Fore.YELLOW}  Select [1-5]: {Style.RESET_ALL}").strip()

    if choice == "1":
        export_df = df.copy()
        label = "all contacts"
    elif choice == "2":
        export_df = df[df["status"] == "active"].copy()
        label = "active contacts"
    elif choice == "3":
        export_df = df[df["meeting_date"].astype(str).str.strip() != ""].copy()
        label = "meetings"
    elif choice == "4":
        export_df = df[df["status"] == "replied"].copy()
        label = "replied contacts"
    else:
        return

    if export_df.empty:
        cprint(f"  No {label} to export.", Fore.YELLOW)
        return

    filename = input(f"  Filename [{default_name}]: ").strip() or default_name
    export_path = Path(__file__).parent / filename

    # Clean up for export — readable columns
    export_cols = ["name", "email", "phone", "company", "title", "status", "step",
                   "last_contact_date", "meeting_date", "meeting_notes", "notes"]
    available_cols = [c for c in export_cols if c in export_df.columns]
    export_df[available_cols].to_csv(export_path, index=False)
    cprint(f"  Exported {len(export_df)} {label} to {export_path}", Fore.GREEN)


# ──────────────────── DEMO MODE ───────────────────────────────

DEMO_CONTACTS = [
    {"name": "Mike Johnson", "email": "mike@demo.com", "phone": "555-9901",
     "company": "Lone Star Roofing", "title": "Owner", "step": 6,
     "notes": "[2026-02-03 09:15 S1] Sent intro | [2026-02-05 14:00 S2] Called, spoke briefly - busy | [2026-02-08 10:30 S3] Sent follow-up"},
    {"name": "Carlos Rivera", "email": "carlos@demo.com", "phone": "555-9902",
     "company": "Summit Commercial", "title": "VP of Operations", "step": 2,
     "notes": "[2026-02-09 11:00 S1] Sent intro email"},
    {"name": "Dana White", "email": "dana@demo.com", "phone": "555-9903",
     "company": "AllWeather Roofing", "title": "GM", "step": 4,
     "notes": "[2026-02-01 08:00 S1] Intro sent | [2026-02-03 13:45 S2] No answer | [2026-02-06 09:00 S3] Opened email"},
    {"name": "Jenny Park", "email": "jenny@demo.com", "phone": "555-9904",
     "company": "BlueSky Contractors", "title": "President", "step": 1,
     "notes": ""},
    {"name": "Hector Ruiz", "email": "hector@demo.com", "phone": "555-9905",
     "company": "Texas Pro Roofing", "title": "Owner", "step": 5,
     "notes": "[2026-01-28 10:00 S1] Intro | [2026-01-30 14:00 S2] Left VM | [2026-02-02 09:30 S3] Follow-up sent | [2026-02-04 11:00 S4] Texted"},
]


def demo_mode() -> None:
    """Walk through every screen the tool produces using fake data. Nothing is saved or sent."""
    cprint("\n╔══════════════════════════════════════════════╗", Fore.YELLOW)
    cprint("║            DEMO MODE — Nothing is real       ║", Fore.YELLOW)
    cprint("║   No emails sent. No CSV changes. Just UI.   ║", Fore.YELLOW)
    cprint("╚══════════════════════════════════════════════╝\n", Fore.YELLOW)

    # Sort by priority: Step 6 email first, then Step 3 email, manual tasks, then Step 1
    ordered = sorted(DEMO_CONTACTS, key=lambda c: (
        0 if STEP_TYPE.get(c["step"]) == "email" and c["step"] > 1 else
        1 if STEP_TYPE.get(c["step"]) != "email" and c["step"] > 1 else 2,
        -c["step"],
    ))

    fu_emails = [c for c in ordered if STEP_TYPE.get(c["step"]) == "email" and c["step"] > 1]
    manual = [c for c in ordered if STEP_TYPE.get(c["step"]) != "email" and c["step"] > 1]
    new_emails = [c for c in ordered if c["step"] == 1]

    # ── Review screen ──
    cprint("╔══════════════════════════════════════╗", Fore.CYAN)
    cprint("║       TODAY'S OUTREACH SUMMARY       ║", Fore.CYAN)
    cprint("╠══════════════════════════════════════╣", Fore.CYAN)
    cprint(f"║  Follow‑up Emails  : {len(fu_emails):>4}            ║", Fore.WHITE)
    cprint(f"║  Manual Tasks      : {len(manual):>4}            ║", Fore.WHITE)
    cprint(f"║  New Outreach      : {len(new_emails):>4}            ║", Fore.WHITE)
    cprint(f"║  Daily Email Limit : {DAILY_EMAIL_LIMIT:>4}            ║", Fore.YELLOW)
    cprint(f"║  Total Actions     : {len(ordered):>4}            ║", Fore.WHITE)
    cprint("╚══════════════════════════════════════╝", Fore.CYAN)

    input(f"\n{Fore.YELLOW}  Press Enter to start the demo walkthrough …{Style.RESET_ALL}")

    action_num = 0

    # ── Follow-up emails ──
    if fu_emails:
        cprint("\n── FOLLOW‑UP EMAILS ──", Fore.CYAN)
    for c in fu_emails:
        action_num += 1
        step = c["step"]
        cprint(f"\n  [{action_num}/{len(ordered)}] {c['name']} — {c['company']} (Step {step})", Fore.CYAN)
        display_notes_log(c["notes"])
        demo_tmpl = get_template_for_step(step, c["name"], c["company"], c["title"])
        if demo_tmpl:
            subject, body = demo_tmpl
            cprint(f"\n  Subject: {subject}", Fore.GREEN)
            cprint(f"  To: {c['email']}", Fore.GREEN)
            cprint(f"  {'─' * 40}", Fore.WHITE)
            for line in body.splitlines():
                cprint(f"  {line}", Fore.WHITE)
            cprint(f"  {'─' * 40}", Fore.WHITE)
        cprint(f"  [DEMO] Email would be sent here.", Fore.YELLOW)
        _demo_notes_prompt(c)
        input(f"{Fore.YELLOW}  Press Enter for next action …{Style.RESET_ALL}")

    # ── Manual tasks ──
    if manual:
        cprint("\n── MANUAL FOLLOW‑UP TASKS ──", Fore.CYAN)
    for c in manual:
        action_num += 1
        step = c["step"]
        task = STEP_TYPE.get(step, "task").upper()
        cprint(f"\n  [{action_num}/{len(ordered)}]", Fore.CYAN)
        cprint(f"  {'─' * 60}", Fore.CYAN)
        cprint(f"  MANUAL TASK: {task}", Fore.CYAN)
        cprint(f"  Name    : {c['name']}", Fore.WHITE)
        cprint(f"  Title   : {c['title']}", Fore.WHITE)
        cprint(f"  Company : {c['company']}", Fore.WHITE)
        cprint(f"  Phone   : {c['phone']}", Fore.WHITE)
        cprint(f"  {'─' * 60}", Fore.CYAN)
        display_notes_log(c["notes"])

        script = get_manual_script(step, c["name"])
        if script:
            cprint(f"\n  {'═' * 56}", Fore.GREEN)
            cprint(f"  SCRIPT:", Fore.GREEN)
            cprint(f"  {'═' * 56}", Fore.GREEN)
            for line in script.strip().splitlines():
                cprint(f"  {line}", Fore.WHITE)
            cprint(f"  {'═' * 56}", Fore.GREEN)

        cprint(f"\n  [DEMO] You would see: Result? (y=Done / n=Skip / stop=Remove)", Fore.YELLOW)
        _demo_notes_prompt(c)
        input(f"{Fore.YELLOW}  Press Enter for next action …{Style.RESET_ALL}")

    # ── New outreach ──
    if new_emails:
        cprint("\n── NEW OUTREACH EMAILS ──", Fore.CYAN)
    for c in new_emails:
        action_num += 1
        step = c["step"]
        cprint(f"\n  [{action_num}/{len(ordered)}] {c['name']} — {c['company']} (Step {step})", Fore.CYAN)
        display_notes_log(c["notes"])
        demo_tmpl = get_template_for_step(step, c["name"], c["company"], c["title"])
        if demo_tmpl:
            subject, body = demo_tmpl
            cprint(f"\n  Subject: {subject}", Fore.GREEN)
            cprint(f"  To: {c['email']}", Fore.GREEN)
            cprint(f"  {'─' * 40}", Fore.WHITE)
            for line in body.splitlines():
                cprint(f"  {line}", Fore.WHITE)
            cprint(f"  {'─' * 40}", Fore.WHITE)
        cprint(f"  [DEMO] Email would be sent here.", Fore.YELLOW)
        _demo_notes_prompt(c)
        if action_num < len(ordered):
            input(f"{Fore.YELLOW}  Press Enter for next action …{Style.RESET_ALL}")

    # ── Summary ──
    cprint(f"\n{'═' * 50}", Fore.GREEN)
    cprint(f"  DEMO COMPLETE", Fore.GREEN)
    cprint(f"  In a real run, {len(fu_emails) + len(new_emails)} emails would have been sent.", Fore.GREEN)
    cprint(f"  {len(manual)} manual tasks would have been prompted.", Fore.GREEN)
    cprint(f"  Notes you entered would be saved to contacts.csv.", Fore.GREEN)
    cprint(f"{'═' * 50}\n", Fore.GREEN)


def _demo_notes_prompt(contact: dict) -> None:
    """Show the notes prompt in demo mode (input is discarded)."""
    note = input(
        f"{Fore.YELLOW}  Add notes (or press Enter to skip): {Style.RESET_ALL}"
    ).strip()
    if note:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
        cprint(f"  [DEMO] Would save: [{timestamp} S{contact['step']}] {note}", Fore.YELLOW)


# ──────────────────── MAIN MENU ───────────────────────────────

def main_menu() -> None:
    """Interactive main menu."""
    colorama_init(autoreset=True)
    load_config()

    cprint("\n╔══════════════════════════════════════════╗", Fore.CYAN)
    cprint("║   SALESLOFT LOCAL — CLI Outreach Tool    ║", Fore.CYAN)
    cprint("╚══════════════════════════════════════════╝", Fore.CYAN)

    df = load_contacts()

    while True:
        cprint("\n── MAIN MENU ──", Fore.CYAN)
        cprint("  1.  Run today's cadence", Fore.WHITE)
        cprint("  2.  Dry-run (preview with real data)", Fore.WHITE)
        cprint("  3.  Retry failed emails", Fore.WHITE)
        cprint(f"  {'─' * 35}", Fore.WHITE)
        cprint("  4.  Who's due today?", Fore.WHITE)
        cprint("  5.  Search / Edit / Delete contacts", Fore.WHITE)
        cprint("  6.  Add contacts manually", Fore.WHITE)
        cprint("  7.  Import contacts from CSV", Fore.WHITE)
        cprint(f"  {'─' * 35}", Fore.WHITE)
        cprint("  8.  View contact status", Fore.WHITE)
        cprint("  9.  Analytics & reports", Fore.WHITE)
        cprint("  10. Daily recap", Fore.WHITE)
        cprint("  11. Export contacts to CSV", Fore.WHITE)
        cprint("  12. Undo (up to 20 steps)", Fore.YELLOW)
        cprint(f"  {'─' * 35}", Fore.WHITE)
        cprint("  13. Demo mode (test drive with fake data)", Fore.YELLOW)
        cprint("  14. Edit email templates", Fore.WHITE)
        cprint("  15. Settings", Fore.WHITE)
        cprint("  0.  Exit", Fore.WHITE)

        choice = input(f"\n{Fore.YELLOW}  Select [0-15]: {Style.RESET_ALL}").strip()

        if choice == "1":
            df = review_and_run(df)
        elif choice == "2":
            df = review_and_run(df, dry_run=True)
        elif choice == "3":
            df = retry_failed(df)
        elif choice == "4":
            show_due_today(df)
        elif choice == "5":
            df = manage_contacts(df)
        elif choice == "6":
            df = add_contacts_interactive(df)
        elif choice == "7":
            path = input("  Path to CSV file: ").strip()
            df = import_contacts_csv(df, path)
        elif choice == "8":
            show_status(df)
        elif choice == "9":
            analytics_menu(df)
        elif choice == "10":
            daily_recap(df)
        elif choice == "11":
            export_report(df)
        elif choice == "12":
            df = undo_last_action(df)
        elif choice == "13":
            demo_mode()
        elif choice == "14":
            edit_templates()
        elif choice == "15":
            show_settings()
        elif choice == "0":
            cprint("\n  Goodbye.\n", Fore.GREEN)
            break
        else:
            cprint("  Invalid choice.", Fore.RED)


if __name__ == "__main__":
    main_menu()
