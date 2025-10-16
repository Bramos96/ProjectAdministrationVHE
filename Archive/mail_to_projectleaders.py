import os
import pandas as pd
from datetime import datetime
import win32com.client as win32

# ───────────────────────────────────────────────────────────
OUTPUT_FOLDER = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Output"
EMAILMAP_FILE = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Projectleiders.xlsx"
EMAILMAP_SHEET = "Sheet1"
OVERVIEW_SHEET = "Overzicht"
TESTMODE = True  # Alles naar jezelf
SIGNATURE_NAME = "Bram Gerrits.htm"
# ───────────────────────────────────────────────────────────

def find_latest_overview(folder):
    files = [f for f in os.listdir(folder) if f.startswith("Overzicht_Projectadministratie_Week") and f.endswith(".xlsx")]
    files.sort(key=lambda x: os.path.getmtime(os.path.join(folder, x)), reverse=True)
    return os.path.join(folder, files[0])

def get_signature_html():
    sig_path = os.path.join(os.environ["APPDATA"], "Microsoft", "Signatures", SIGNATURE_NAME)
    if os.path.exists(sig_path):
        with open(sig_path, encoding="utf-8") as f:
            return f.read()
    else:
        print(f"⚠️ Geen handtekeningbestand gevonden: {sig_path}")
        return "<p>Groet,<br><br>Bram.</p>"

# 1. Laad gegevens
overview_path = find_latest_overview(OUTPUT_FOLDER)
df = pd.read_excel(overview_path, sheet_name=OVERVIEW_SHEET, header=0)
df.columns = df.columns.str.strip()

df = df[["Projectnummer", "Projectleider", "Actiepunten Projectleider", "Actiepunten Elders"]]
# Laat df intact voor 'Elders'; filter alleen voor PL-mails
df_pl = df.dropna(subset=["Projectleider"]).copy()


# 2. Laad e-mailmapping
df_emails = pd.read_excel(EMAILMAP_FILE, sheet_name=EMAILMAP_SHEET)
df_emails.columns = df_emails.columns.str.strip()
email_map = dict(zip(df_emails["Naam"], df_emails["Email"]))

# Mapping van actiepunt → ontvangernaam (zoals die in je Projectleiders.xlsx staat)
RECIPIENT_BY_PHRASE = {
    "Gesloten SO met openstaande bestelling": "Inkoop",
    "Gesloten SO met openstaande PO - Prod": "Judith",
    "Gesloten SO met openstaande PO - Proto": "Inkoop",
}

# Verzamelbak per ontvanger: lijst met (Projectnummer, Actiepunt)
from collections import defaultdict
elders_bucket = defaultdict(list)

# Vul de verzamelbak op basis van 'Actiepunten Elders' (niet filteren op Projectleider!)
for _, r in df.iterrows():
    txt = str(r.get("Actiepunten Elders", "") or "")
    if not txt.strip():
        continue
    for phrase, recipient in RECIPIENT_BY_PHRASE.items():
        if phrase in txt:
            elders_bucket[recipient].append((r["Projectnummer"], phrase))


# 3. Haal handtekening
signature_html = get_signature_html()

# 4. Verstuur mails
outlook = win32.Dispatch("Outlook.Application")
mail_count = 0

for leider, sub in df_pl.groupby("Projectleider"):
    if leider not in email_map:
        print(f"⚠️ Geen e-mailadres gevonden voor: {leider}")
        continue

    sub = sub.dropna(subset=["Actiepunten Projectleider"])
    if sub.empty:
        continue

    voornaam = leider.split()[0]

    # Bouw HTML-tabel
    table_rows = ""
    for _, row in sub.iterrows():
        actiepunten_html = str(row['Actiepunten Projectleider']).replace('\n', '<br>')
        table_rows += f"<tr><td>{row['Projectnummer']}</td><td>{actiepunten_html}</td></tr>"

    html_body = f"""
    <p>Beste {voornaam},</p>
    <p>Hieronder vind je een overzicht van je openstaande actiepunten per project:</p>
    <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">
        <tr style="background-color:#4F81BD;color:white;">
            <th style="width: 150px;">Projectnummer</th>
            <th style="width: 300px;">Actiepunten</th>
        </tr>
        {table_rows}
    </table>
    {signature_html}
    """

    to_address = email_map[leider] if not TESTMODE else "bram.gerrits@vhe.nl"

    mail = outlook.CreateItem(0)
    mail.To = to_address
    mail.Subject = f"[Projectadministratie] Actiepunten – {voornaam}"
    mail.HTMLBody = html_body
    mail.Display()

    print(f"✅ Mail klaargezet voor: {voornaam} → {to_address}")
    mail_count += 1

# 4b. Mails voor 'Actiepunten Elders' per ontvanger (PER PROJECTNUMMER)
for recipient_name, items in elders_bucket.items():
    if not items:
        continue
    if recipient_name not in email_map:
        print(f"⚠️ Geen e-mailadres gevonden voor: {recipient_name}")
        continue

    # Bouw per projectnummer de lijst met actiepunten (dedup)
    per_project = {}
    for proj, phrase in items:
        per_project.setdefault(proj, [])
        if phrase not in per_project[proj]:
            per_project[proj].append(phrase)

    if not per_project:
        continue

    # HTML-tabel: Projectnummer | Actiepunten (meerdere regels onder elkaar)
    table_rows = ""
    for proj in sorted(per_project, key=lambda x: str(x)):
        acties_html = "<br>".join(per_project[proj])
        table_rows += f"<tr><td>{proj}</td><td>{acties_html}</td></tr>"

    html_body_elders = f"""
    <p>Beste {recipient_name},</p>
    <p>Hieronder vind je een overzicht van de openstaande actiepunten per project:</p>
    <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">
        <tr style="background-color:#4F81BD;color:white;">
            <th style="width: 150px;">Projectnummer</th>
            <th style="width: 300px;">Actiepunten</th>
        </tr>
        {table_rows}
    </table>
    {signature_html}
    """

    to_address = email_map[recipient_name] if not TESTMODE else "bram.gerrits@vhe.nl"

    mail = outlook.CreateItem(0)
    mail.To = to_address
    mail.Subject = f"[Projectadministratie] Actiepunten – {recipient_name}"
    mail.HTMLBody = html_body_elders
    mail.Display()

    print(f"✅ Mail (Elders) klaargezet voor: {recipient_name} → {to_address}")
    mail_count += 1


print(f"\n📬 Totaal {mail_count} mails klaargezet.")
