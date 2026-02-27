"""
Create draft outreach emails in Outlook from a spreadsheet.

Reads a list of contacts from an Excel file and creates one Outlook draft per row,
personalised with the contact's name, organisation, and a pre-written hook sentence.
Drafts are saved (not sent) and throttled to avoid triggering spam filters.

Usage:
  Set TARGET_SMTP and XLSX_PATH at the top, then run:
  python email_drafter.py

Requirements:
  pip install pandas openpyxl pywin32
  Windows only (uses win32com to control Outlook)
"""

import time
import random
from pathlib import Path

import pandas as pd
import win32com.client as win32


XLSX_PATH   = Path(r"send_list.xlsx")
SHEET       = "Sheet1"
TARGET_SMTP = "Ben@websedge.com"

# Expected column names in the spreadsheet
COL_EMAIL      = "Email"
COL_LAST       = "Last"
COL_TITLE      = "Title"
COL_PARENT_ORG = "ParentOrg"
COL_HOOK       = "Hook"


def build_html(title: str, last: str, parent_org: str, hook: str) -> str:
    """Assemble the full HTML email body."""
    style = 'style="font-family: Arial, sans-serif; font-size: 9pt; color: black;"'

    return f"""<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <style type="text/css" style="display:none;"> P {{margin-top:0;margin-bottom:0;}} </style>
</head>
<body dir="ltr">
  <p><span {style}>Dear {title}. {last},</span></p>
  <p><span {style}>&nbsp;</span></p>
  <p><span {style}>
    I would like to schedule a call between you and Mark Rose, IAF Congress TV director, about the
    opportunity to feature {parent_org} in a pre-recorded video case study as part of the official
    broadcast for the 76th International Astronautical Congress (IAC) in Sydney
    (29 September – 3 October 2025) and online.
  </span></p>
  <p><span {style}>&nbsp;</span></p>
  <p><span {style}>{hook}</span></p>
  <p><span {style}>&nbsp;</span></p>
  <p><span {style}>
    As I'm sure you are aware, the IAC is the world's premier global space event and a leading forum
    for presenting innovative developments in space science, technology, and exploration. The
    International Astronautical Federation has again partnered with global knowledge-driven media
    company, WebsEdge, to produce IAF Congress TV, the official broadcast at the International
    Astronautical Congress. For the last 10 years, IAF Congress TV has captured and shared the most
    exciting breakthroughs in both fundamental and applied space research — spotlighting the people,
    projects, and institutions shaping the future of the space sector.
  </span></p>
  <p><span {style}>&nbsp;</span></p>
  <p><span {style}>
    As a key part of this year's broadcast, we are once again inviting a select number of leading
    organisations at the cutting edge of space technology to feature in five-minute documentary-style
    video profiles, shown throughout the Congress and available to attendees online. These videos offer
    a high-impact way to showcase your key research, programmes, or technologies to a global,
    interdisciplinary audience.
  </span></p>
  <p><span {style}>&nbsp;</span></p>
  <p><span {style}>
    Through our research, we are considering a number of companies, centres and institutes as
    potential candidates including {parent_org}, and I am keen to arrange a conversation between you
    and Mark Rose to make sure there is a strong fit. <b>I must emphasise that there is a cost
    involved in this opportunity to be profiled, which covers the production, distribution, and full
    ownership of the film and all additional footage.</b>
  </span></p>
  <p><span {style}>&nbsp;</span></p>
  <p><span {style}>
    In advance of the conversation, it would be useful for you to have a look at one or two of the
    groups that we profiled in the same way at previous IAC meetings, to give you an idea of the
    style of film we would produce with you.
  </span></p>
  <p><span {style}>&nbsp;</span></p>
  <p><span {style}>
    As such, please could you email back with some suitable times over the next few days when Mark
    can call you to discuss this? He will be in meetings for the majority of today, but is fairly
    open over the next week or so.
  </span></p>
  <p><span {style}>&nbsp;</span></p>
  <p><span {style}>I look forward to hearing back from you with a convenient time to speak.</span></p>
  <p><span {style}>&nbsp;</span></p>
  <p><span {style}>Best wishes,</span></p>
  <p><span {style}>&nbsp;</span></p>
  <p><span {style}>Ben</span></p>
  <p><span {style}>&nbsp;</span></p>
  <p>
    <span style="font-family: Arial; color: rgb(37,37,37);"><b>Ben Varvill</b> | </span>
    <span style="font-family: Arial; color: rgb(0,32,96);"><b>Researcher – ICMA TV – A WebsEdge Channel</b></span>
  </p>
  <p><span {style}>&nbsp;</span></p>
  <p><span {style}>
    For further information on IAF Congress TV, please visit:
    <a href="https://www.iafastro.org/events/iac/international-astronautical-congress-2025/media/iaf-congress-tv.html">iaf-congress-tv.html</a>
    or contact Evelina Hedman at <a href="mailto:evelina.hedman@iafastro.org">evelina.hedman@iafastro.org</a>.
  </span></p>
  <p><span style="font-family: Arial; font-size: 10pt; color: black;">
    UK: 6 Henrietta Street | London | WC2E 8PT | UK<br>
    USA: 3244 Prospect Street NW | Washington DC | 20007 | USA<br>
    W: <a href="http://www.websedge.com/">www.websedge.com</a>
  </span></p>
  <p><span style="font-size: 10.5pt; color: black;">D-U-N-S number: 235211278</span></p>
  <p><span style="font-family: Arial; font-size: 7pt; color: gray;">
    WebsEdge is a trading name of WebsEdge Limited, registered in England and Wales, number 3520183.
    This email is confidential. If you are not the intended recipient, please notify the sender and
    delete this message. Internet emails are not necessarily secure.
  </span></p>
</body>
</html>"""


def main():
    df = pd.read_excel(XLSX_PATH, sheet_name=SHEET, engine="openpyxl")
    print(f"Loaded {len(df)} rows from {XLSX_PATH}")

    outlook = win32.Dispatch("Outlook.Application")
    session = outlook.GetNamespace("MAPI")

    # Use the specified send account, fall back to the first account
    send_acct = None
    for acct in session.Accounts:
        if acct.SmtpAddress.lower() == TARGET_SMTP.lower():
            send_acct = acct
            break
    if not send_acct:
        send_acct = session.Accounts.Item(1)
        print(f"Target account not found — using {send_acct.SmtpAddress}")

    count = 0
    for _, row in df.iterrows():
        email      = str(row.get(COL_EMAIL, "")).strip()
        last       = str(row.get(COL_LAST, "")).strip()
        title      = str(row.get(COL_TITLE, "")).strip()
        parent_org = str(row.get(COL_PARENT_ORG, "")).strip()
        hook       = str(row.get(COL_HOOK, "")).strip()

        if not email:
            continue

        mail = outlook.CreateItem(0)
        mail.SendUsingAccount = send_acct
        mail.To      = email
        mail.Subject = f"{parent_org} – Film - 2025 IAF Congress Thought Leadership Film Series"
        mail.HTMLBody = build_html(title, last, parent_org, hook)
        mail.Save()
        count += 1

        print(f"  Draft saved: {last} @ {parent_org}")
        time.sleep(random.uniform(5, 10))

    print(f"\nDone — {count} drafts saved to Outlook.")


if __name__ == "__main__":
    main()
