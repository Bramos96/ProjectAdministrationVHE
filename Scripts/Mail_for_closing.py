import os
import re
import html
import pandas as pd
from datetime import datetime
import win32com.client as win32

# ───────────────────────────────────────────────────────────
OUTPUT_FOLDER    = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Output"
EMAILMAP_FILE    = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Projectleiders.xlsx"
EMAILMAP_SHEET   = "Sheet1"
OVERVIEW_SHEET   = "Overzicht"
TESTMODE         = False
SIGNATURE_NAME   = "Bram Gerrits.htm"
BASE_FONT_STYLE  = "font-family: Aptos, Calibri, Segoe UI, Arial, sans-serif; font-size: 11pt; line-height: 1.35;"
# ───────────────────────────────────────────────────────────

TARGET_PHRASES = [
    "Opbrengsten binnen",
    "Gesloten SO, project sluiten na goedkeuring",
]
REASON_LABEL = "Gesloten verkooporder & volledig gefactureerd."

def find_latest_overview(folder: str) -> str:
    files = [f for f in os.listdir(folder) if f.startswith("Overzicht_Projectadministratie_Week") and f.endswith(".xlsx")]
    if not files:
        raise FileNotFoundError(f"Geen Overzicht-bestanden gevonden in: {folder}")
    files.sort(key=lambda x: os.path.getmtime(os.path.join(folder, x)), reverse=True)
    return os.path.join(folder, files[0])

def get_signature_html() -> str:
    sig_path = os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Signatures", SIGNATURE_NAME)
    if os.path.exists(sig_path):
        with open(sig_path, encoding="utf-8") as f:
            return f.read()
    print(f"⚠️ Geen handtekeningbestand gevonden: {sig_path}")
    return "<p>Vriendelijke groet,<br><br>Bram.<br><br>[Automatisch verstuurd]</p>"

def to_bool(x) -> bool:
    s = str(x).strip().lower()
    return s in {"true", "1", "ja", "yes", "y", "waar", "ok", "oké", "x"}

def is_valid_email(addr: str) -> bool:
    return isinstance(addr, str) and "@" in addr and "." in addr and addr.strip() != ""

def norm_txt(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = s.replace("\u00A0", " ")
    s = s.replace("–", "-").replace("—", "-")
    s = re.sub(r"\s+", " ", s)
    return s.strip().lower()

overview_path = find_latest_overview(OUTPUT_FOLDER)
df = pd.read_excel(overview_path, sheet_name=OVERVIEW_SHEET, header=0)
df.columns = df.columns.str.strip()

CANDIDATE_COL_PRIMARY = "Bespreekpunten"
CANDIDATE_COL_FALLBACKS = ["Bespreekpunten (PL)", "Actiepunten Projectleider"]
ACTION_BRAM_COL = "Actiepunten Bram"

required_base = ["Projectnummer", "Projectleider"]
missing_base = [c for c in required_base if c not in df.columns]
if missing_base:
    raise ValueError(f"Ontbrekende kolommen in '{OVERVIEW_SHEET}': {missing_base}")

content_col = None
if CANDIDATE_COL_PRIMARY in df.columns:
    content_col = CANDIDATE_COL_PRIMARY
else:
    for fb in CANDIDATE_COL_FALLBACKS:
        if fb in df.columns:
            content_col = fb
            break
if content_col is None:
    raise ValueError("Geen geschikte bronkolom gevonden. Verwacht één van: 'Bespreekpunten', 'Bespreekpunten (PL)', 'Actiepunten Projectleider'")

cols = [c for c in ["Projectnummer", "Projectnaam", "Projectleider", content_col, ACTION_BRAM_COL] if c in df.columns]
df = df[cols].copy()

emap = pd.read_excel(EMAILMAP_FILE, sheet_name=EMAILMAP_SHEET)
emap.columns = emap.columns.str.strip()
if "Naam" not in emap.columns or "Email" not in emap.columns:
    raise ValueError("Kolommen 'Naam' en/of 'Email' ontbreken in Projectleiders.xlsx")

e_map = dict(zip(emap["Naam"], emap["Email"]))
w_map = dict(zip(emap["Naam"], emap["Whitelist"].map(to_bool) if "Whitelist" in emap.columns else [False]*len(emap)))

def get_email(name: str) -> str:
    addr = (e_map.get(name) or "").strip()
    return addr if is_valid_email(addr) else ""

phrases_norm = [norm_txt(p) for p in TARGET_PHRASES]

def meets_both_conditions(txt: str) -> bool:
    if not isinstance(txt, str) or not txt.strip():
        return False
    t = norm_txt(txt)
    return all(p in t for p in phrases_norm)

records = []
for _, r in df.iterrows():
    actie_bram = str(r.get(ACTION_BRAM_COL, "")).strip()
    if actie_bram:
        continue  # overslaan als er actiepunten bij Bram staan
    if meets_both_conditions(r.get(content_col, "")):
        records.append({
            "Projectnummer": r.get("Projectnummer"),
            "Projectnaam": r.get("Projectnaam", ""),
            "Projectleider": r.get("Projectleider"),
            "Reasons": [REASON_LABEL],
        })
closing_df = pd.DataFrame(records)

if closing_df.empty:
    print("ℹ️ Geen projecten gevonden die aan beide sluit-voorwaarden voldoen of actiepunten bij Bram hebben.")
    raise SystemExit(0)

signature_html = get_signature_html()
try:
    outlook = win32.Dispatch("Outlook.Application")
    _ = outlook.GetNamespace("MAPI")
except Exception as e:
    raise RuntimeError(f"Outlook niet beschikbaar: {e}")

week_str = datetime.now().strftime("Week %W, %Y")
test_tag = "[TEST] " if TESTMODE else ""
mail_count = 0

for leader, sub in closing_df.groupby("Projectleider"):
    if not w_map.get(leader, False):
        print(f"⛔ Niet op whitelist: {leader} – mail overgeslagen")
        continue
    real_addr = get_email(leader)
    if not real_addr:
        print(f"⛔ Geen geldig e-mailadres voor: {leader} – mail overgeslagen")
        continue

    rows_html = ""
    for _, row in sub.iterrows():
        reasons_html = html.escape(REASON_LABEL)
        proj = html.escape(str(row.get("Projectnummer", "")))
        rows_html += (
            f"<tr>"
            f"<td style='text-align:left; font-family: Aptos, Calibri, Segoe UI, Arial, sans-serif; font-size:11pt;'>{proj}</td>"
            f"<td style='text-align:left; font-family: Aptos, Calibri, Segoe UI, Arial, sans-serif; font-size:11pt;'>{reasons_html}</td>"
            f"</tr>"
        )

    if not rows_html:
        continue

    leader_first = str(leader).split()[0]
    html_body = f"""
    <div style="{BASE_FONT_STYLE}">
        Hi {leader_first},<br><br>
        Het viel me op dat de volgende projecten geleverd/gefactureerd zijn:<br><br>
        <table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse; font-family: Aptos, Calibri, Segoe UI, Arial, sans-serif; font-size:11pt;">
            <tr style="background-color:#4F81BD; color:white; font-family: Aptos, Calibri, Segoe UI, Arial, sans-serif; font-size:11pt;">
                <th style="text-align:left; width:160px; font-family: Aptos, Calibri, Segoe UI, Arial, sans-serif; font-size:11pt;">Projectnummer</th>
                <th style="text-align:left; width:360px; font-family: Aptos, Calibri, Segoe UI, Arial, sans-serif; font-size:11pt;">Reden</th>
            </tr>
            {rows_html}
        </table>
        <br>
        Kunnen deze projecten volgens jou gesloten worden?<br><br>
        Alvast dank voor je reactie!<br><br>
        Vriendelijke groet,<br><br>
        Bram<br><br>
        [Automatisch verstuurd]
    </div>
    """

    to_address = "bram.gerrits@vhe.nl" if TESTMODE else real_addr
    mail = outlook.CreateItem(0)
    mail.To = to_address
    mail.Subject = f"{test_tag}[Projectadministratie] Bevestiging projectsluiting – {leader_first} ({week_str})"
    mail.HTMLBody = html_body
    mail.Display()

    print(f"✅ Mail (closing) klaargezet voor: {leader} → {to_address}")
    mail_count += 1

print(f"\n📬 Totaal {mail_count} mails klaargezet (closing).")