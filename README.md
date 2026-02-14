# SalesloftLocal — CLI Sales Cadence Tool

A single-file Python CLI that manages a 6-step B2B sales outreach cadence:
**Email → Call → Email → SMS → LinkedIn → Email**

Reads contacts from `contacts.csv`, prioritises follow-ups over new leads,
sends emails via Office365/Outlook SMTP, and prompts for manual tasks.

---

## Quick Start

```bash
pip install pandas colorama
python outreach.py          # interactive menu
python outreach.py --dry-run  # preview today's actions without sending
```

---

## Configuration

On first run a `config.ini` is generated automatically. Edit it to set your
SMTP server, daily email limit, and cadence wait-days:

```ini
[smtp]
server = smtp.office365.com
port = 587
email = you@yourcompany.com
# Leave password blank — use OUTREACH_PASSWORD env var instead

[limits]
daily_email_limit = 150
throttle_min_seconds = 30
throttle_max_seconds = 90

[cadence]
step1 = 0   # initial outreach — due immediately
step2 = 2
step3 = 3
step4 = 2
step5 = 2
step6 = 4
```

Store your password as an environment variable, not in the file:

```bash
export OUTREACH_EMAIL=you@yourcompany.com
export OUTREACH_PASSWORD=yourpassword
```

---

## Email Templates

Templates for Steps 1, 3, and 6 (the email steps) are stored in
`templates.json`. Edit them from the menu (**option 14**) or directly in the
JSON file. Use `<<name>>`, `<<company>>`, `<<title>>` as substitution tokens.

---

## Automated Daily Run (`--auto`)

The `--auto` flag runs the full cadence without any interactive prompts.
All due email steps are sent automatically. Manual steps (call, SMS, LinkedIn)
are printed as reminders to stdout and skipped.

**Requirements:** `OUTREACH_EMAIL` and `OUTREACH_PASSWORD` must be set as
environment variables — no interactive credential prompt in headless mode.

### Cron setup (Linux/macOS)

Run at 8 am Monday–Friday, append output to `auto.log`:

```cron
0 8 * * 1-5 cd /path/to/SalesloftLocal && \
  OUTREACH_EMAIL=you@co.com \
  OUTREACH_PASSWORD=secret \
  python outreach.py --auto >> auto.log 2>&1
```

Add this line via `crontab -e`.

### Log output format

```
[2026-02-13 08:00:01] Starting auto-run: 3 email(s), 1 manual reminder(s).
[2026-02-13 08:00:02] SMTP connected.
[2026-02-13 08:00:03] SENT  Step 1 → Alice Smith <alice@acme.com>
[2026-02-13 08:00:34] SENT  Step 3 → Bob Jones <bob@beta.com>
[2026-02-13 08:01:05] REMINDER  CALL Step 2 due for Carlos Rivera (Summit Commercial) — 555-9902
[2026-02-13 08:01:05] Done. Emails sent this run: 2.
```

---

## Contact Manager Commands

Open the contact manager with menu option **5**.

| Command | Description |
|---------|-------------|
| `search <term>` | Search by name, company, or email |
| `list` | Show all contacts (paginated, 20/page) |
| `list active` | Filter by status |
| `list step2` | Filter by cadence step |
| `list tag:texas` | Filter by tag |
| `edit <#>` | Edit a single contact's fields |
| `delete <#>` | Delete a contact (requires `DELETE` confirmation) |
| `snooze <#>` | Snooze a contact until a date |
| `tag <#> <tag>` | Add a tag to a contact |
| `untag <#> <tag>` | Remove a tag from a contact |
| `bulk-status <status> <target>` | Set status on many contacts at once |
| `bulk-snooze <date\|clear> <target>` | Snooze many contacts at once |
| `bulk-delete <target>` | Delete many contacts (requires `DELETE ALL`) |
| `back` | Return to main menu |

**Bulk targets:** `all` · `stepN` (e.g. `step1`) · `status:<val>` · `1,4,7` (index list)

---

## Dependencies

```
pandas>=1.5.0
colorama>=0.4.6
```

Install: `pip install pandas colorama`

---

## File Layout

```
SalesloftLocal/
├── outreach.py       # main application
├── contacts.csv      # contact database (auto-created)
├── config.ini        # settings (auto-created on first run)
├── templates.json    # email templates (auto-created on first run)
├── .undo_history.json  # undo snapshots (auto-managed)
└── auto.log          # headless run log (if using --auto with cron)
```
