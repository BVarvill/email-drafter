# email_drafter.py
import pandas as pd
import win32com.client as win32
import time, random
from pathlib import Path

# ——— 1) CONFIG ——————————————————————————————————————————
XLSX_PATH = Path("/Users/benvarvill/Documents/Automation:IAF/send_list.xlsx")
SHEET        = "Sheet1"
TARGET_SMTP  = "Ben@websedge.com"

# Excel column names
COL_EMAIL       = "Email"
COL_LAST        = "Last"
COL_TITLE       = "Title"
COL_PARENT_ORG  = "ParentOrg"
COL_LINK        = "Link"  # (if you ever want to insert a dynamic link)
COL_HOOK        = "Hook"

# ——— 2) READ THE EXCEL ————————————————————————————————————
df = pd.read_excel(XLSX_PATH, sheet_name=SHEET, engine="openpyxl")

# ——— 3) START OUTLOOK & PICK ACCOUNT —————————————————————
outlook = win32.Dispatch("Outlook.Application")
session = outlook.GetNamespace("MAPI")
send_acct = None
for acct in session.Accounts:
    if acct.SmtpAddress.lower() == TARGET_SMTP.lower():
        send_acct = acct
        break
if not send_acct:
    send_acct = session.Accounts.Item(1)

# ——— 4) LOOP & DRAFT ————————————————————————————————————
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

    html_template = f"""
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <style type="text/css" style="display:none;"> P {{margin-top:0;margin-bottom:0;}} </style>
</head>
<body dir="ltr">
  <p class="elementToProof">
    <span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">
      Dear {title}. {last},
    </span>
  </p>
  <p class="elementToProof">
    <span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">&nbsp;</span>
  </p>
  <p class="elementToProof">
    <span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">
      I would like to schedule a call between you and Mark Rose, IAF Congress TV director, about the opportunity to feature VITO Remote Sensing in a pre-recorded video case study as part of the official broadcast for the 76th International Astronautical Congress (IAC) in Sydney (29 September – 3 October 2025) and online.
    </span>
  </p>
  <p class="elementToProof">
    <span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">&nbsp;</span>
  </p>
  <p class="elementToProof">    
    <span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">
      {hook}      
    </span>
  <p class="elementToProof"><span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">&nbsp;</span></p>
  <p class="elementToProof">
    <span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">
      As I’m sure you are aware, the IAC is the world’s premier global space event and a leading forum for presenting innovative developments in space science, technology, and exploration. The International Astronautical Federation has again partnered with global knowledge-driven media company, WebsEdge, to produce IAF Congress TV, the official broadcast at the International Astronautical Congress.  For the last 10 years, IAF Congress TV has captured and shared the most exciting breakthroughs in both fundamental and applied space research — spotlighting the people, projects, and institutions shaping the future of the space sector.
    </span>
    <span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">
  <p class="elementToProof">
    <span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">
      As a key part of this year’s broadcast, we are once again inviting a select number of leading organizations at the cutting edge of space technology to feature in five-minute documentary-style video profiles, which will be shown throughout the Congress and available to attendees online. These videos offer a high-impact way to showcase your key research, programs, or technologies to a global, interdisciplinary audience.
    </span>
  </p>
  <p class="elementToProof"><span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">&nbsp;</span></p>
  <p class="elementToProof">
    <span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">
      Through our research, we are considering a number of companies, centers and institutes as potential candidates to sponsor these documentary features including {parent_org}, and I am keen to arrange a conversation between you and Mark
      Rose to make sure that there is a strong fit. <b>I must emphasise that there is a cost involved in this
      opportunity to be profiled, which covers the production, distribution, and full ownership of the film and all
      additional footage.</b>
    </span>
  </p>
  <p class="elementToProof"><span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">&nbsp;</span></p>
  <p class="elementToProof">
    <span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">
      In advance of the conversation, it would be useful for you to have a look at one or two of the groups that we profiled in the same way at previous IAC meetings as this will give you an idea on the style of film we would produce with you. You can see a few of those films here:
    </span>
  </p>
  <p class="elementToProof"><span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">&nbsp;</span></p>
  <p class="elementToProof">
    <span style="font-family: Arial, sans-serif; font-size: 9pt; color: rgb(25, 25, 25);">
      As such, please could you email back with some suitable times over the next few days when Mark can call you to
      discuss this? He will be in meetings for the majority of today but is fairly open over the next week or so, if you can suggest a couple of times for an initial call?
    </span>
  </p>
  <p class="elementToProof"><span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">&nbsp;</span></p>
  <p class="elementToProof">
    <span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">
      I look forward to hearing back from you with a convenient time to speak.
    </span>
  </p>
  <p class="elementToProof"><span style="font-family: Arial, sans-serif; font-size: 9pt;">&nbsp;</span></p>
  <p class="elementToProof"><span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">Best wishes,</span></p>
  <p class="elementToProof"><span style="font-family: Arial, sans-serif; font-size: 9pt;">&nbsp;</span></p>
  <p class="elementToProof"><span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">Ben</span></p>
  <p class="elementToProof"><span style="font-family: Arial, sans-serif; font-size: 9pt;">&nbsp;</span></p>
  <p class="elementToProof">
    <span style="font-family: Arial, sans-serif; color: rgb(37, 37, 37);"><b>Ben Varvill</b> | </span>
    <span style="font-family: Arial, sans-serif; color: rgb(0, 32, 96);"><b>Researcher – ICMA TV – A WebsEdge Channel</b></span>
  </p>
  <p class="elementToProof"><span style="font-family: Arial, sans-serif; font-size: 10pt; color: black;">&nbsp;</span></p>
  <p class="elementToProof">
    <span style="font-family: Arial, sans-serif; font-size: 9pt; color: black; background-color: white;">
      For further information on IAF Congress TV, please visit: https://www.iafastro.org/events/iac/international-astronautical-congress-2025/media/iaf-congress-tv.html or feel free to contact Evelina Hedman at <a href="mailto:evelina.hedman@iafastro.org">evelina.hedman@iafastro.org</a>.
    </span>
  </p>
  <p class="elementToProof"><span style="font-family: Arial, sans-serif; font-size: 8pt; color: gray; background-color: white;">&nbsp;</span></p>
  <p class="elementToProof">
    <span style="font-family: Arial, sans-serif; font-size: 10pt; color: black;">
      UK: 6 Henrietta Street | London | WC2E 8PT | UK
      <br>USA: 3244 Prospect Street NW | Washington DC | 20007 | USA
    </span>
  </p>
  <p class="elementToProof">
    <span style="font-family: Arial, sans-serif; font-size: 10pt; color: black;">
      W: <a href="http://www.websedge.com/">www.websedge.com</a>
    </span>
  </p>
  <p class="elementToProof"><span style="font-family: Arial, sans-serif; font-size: 9pt; color: black;">&nbsp;</span></p>
  <p class="elementToProof"><span style="font-size: 10.5pt; color: black;">D-U-N-S number: 235211278</span></p>
  <p class="elementToProof"><span style="font-size: 10.5pt; color: rgb(64, 64, 64);">&nbsp;</span></p>
  <p class="elementToProof">
    <span style="font-family: Arial, sans-serif; font-size: 7pt; color: gray; background-color: white;">
      WebsEdge is a trading name of WebsEdge Limited, registered in England and Wales, registered number 3520183.
      Confidentiality Notice: The information contained in this email is confidential and may be legally privileged.
      If you are not the intended recipient, you are hereby notified that any disclosure, copying, distribution,
      or reliance upon the contents of this email is strictly prohibited. The statements and opinions expressed
      in this email are those of the author and do not necessarily reflect those of the company. The company
      does not take any responsibility for the views of the author. If you have received this email transmission
      in error, please reply to the sender and then delete the message from your inbox. Internet e-mails are not
      necessarily secure. WebsEdge Limited does not accept responsibility for changes made to this message after
      it was sent. Whilst all reasonable care has been taken to avoid the transmission of viruses, it is the
      responsibility of the recipient to ensure that the onward transmission, opening or use of this message
      and any attachments will not adversely affect its systems or data. No responsibility is accepted by
      WebsEdge Limited in this regard and the recipient should carry out such virus and other checks as it
      considers appropriate.
    </span>
  </p>
</body>
</html>
    """

    mail.HTMLBody = html_template
    mail.Save()
    count += 1

    # throttle
    time.sleep(random.uniform(5, 10))

print(f"Done: created {count} draft emails in Outlook.")