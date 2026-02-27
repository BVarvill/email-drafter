# Email Drafter

Creates personalised draft emails in Outlook from a spreadsheet of contacts.

Reads a list of leads and generates one Outlook draft per row, with each email
personalised to the contact's name, organisation, and a pre-written hook sentence.
Drafts are saved (not sent) and throttled with a random delay to avoid spam filters.

## What It Does

For each row in the spreadsheet:

1. Reads the contact's name, title, organisation, and hook sentence
2. Builds an HTML email body from a template
3. Saves it as an Outlook draft under the specified send account
4. Waits 5-10 seconds before the next draft

## Setup

Windows only — uses `win32com` to control Outlook.

```bash
pip install pandas openpyxl pywin32
```

Outlook must be open and configured with the send account.

## Configuration

Edit the constants at the top of `email_drafter.py`:

```python
XLSX_PATH   = Path(r"send_list.xlsx")   # path to your contact spreadsheet
TARGET_SMTP = "you@yourdomain.com"       # Outlook account to send from
```

## Input Spreadsheet Format

| Column | Description |
|--------|-------------|
| `Email` | Recipient email address |
| `Last` | Recipient last name |
| `Title` | Salutation (Dr., Prof., Mr., etc.) |
| `ParentOrg` | Organisation name (used in subject + body) |
| `Hook` | Personalised hook sentence for this contact |

## Usage

```bash
python email_drafter.py
```

Drafts are saved to your Outlook Drafts folder. Review them before sending.
