import os
import pandas as pd
from datetime import date
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# ───────────────────────────────────────────
# PAS JE PAD HIER AAN
OUTPUT_FILE = (
    r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration"
    r"\Output\Overzicht_Projectadministratie_Week{week}_{year}.xlsx"
)
# ───────────────────────────────────────────

def make_actions_projectleider(row, today):
    """Return bullet list for Actiepunten Projectleider."""
    bullets = []

    # 1) Einddatum verlopen?
    ed_ts = pd.to_datetime(row.get("Einddatum"), errors="coerce")
    if pd.notna(ed_ts) and today > ed_ts.date():
        bullets.append("• Einddatum verlopen")

    # 2) Leverdatum(s) verlopen?
    lv_ts = pd.to_datetime(row.get("Leverdatum"), errors="coerce")
    if pd.notna(lv_ts) and today > lv_ts.date():
        bullets.append("• Leverdatum(s) verlopen")

    # 3) Budget Kosten ontbreekt?
    bk = row.get("Budget Kosten")
    try:
        if pd.notna(bk) and float(bk) == 0:
            bullets.append("• Budget kosten toevoegen")
    except:
        pass

    # 4) Budget Opbrengsten ontbreekt?
    bo = row.get("Budget Opbrengsten")
    try:
        if pd.notna(bo) and float(bo) == 0:
            bullets.append("• Budget opbrengsten toevoegen")
    except:
        pass

    return "\n".join(bullets)

def make_bespreekpunten(row):
    """Return bullet list for Bespreekpunten (automatic conclusions)."""
    bullets = []

    def to_float(x):
        try:
            return float(x)
        except:
            return None

    def is_yes(v):
        return str(v).strip().lower() in {"ja", "true", "1", "y", "yes"}

    def is_nee(v):
        # Alleen 'nee' telt als Nee; lege cellen tellen niet mee
        s = str(v).strip().lower()
        return s in {"nee", "no", "n", "false", "0"}

    # 1) Conclusie: Negatief resultaat bespreken
    vr = to_float(row.get("Verwacht resultaat"))
    if vr is not None and vr < 0:
        bullets.append("• Negatief resultaat bespreken")

    # 2) Conclusie: Klaar om te sluiten?
    bo = to_float(row.get("Budget Opbrengsten"))
    wo = to_float(row.get("Werkelijke opbrengsten"))
    if bo is not None and wo is not None:
        if bo != 0 and bo <= wo:
            bullets.append("• Klaar om te sluiten?")

    # 3) Bestaande regel: Openstaande SO = Ja → dochterproject
    val_so = row.get("Openstaande SO", "")
    if is_yes(val_so):
        bullets.append("• Gesloten SO, maar openstaande SO dochterproject")

    # 4) NIEUW: als alle drie 'Nee' zijn → vervang evt. 'Klaar om te sluiten?' door definitieve conclusie
    cols_sc = ["Openstaande bestelling", "Openstaande SO", "Openstaande PO"]
    vals = [row.get(c, "") for c in cols_sc]

    # Alleen true als ELK van de drie expliciet 'Nee' is (lege cellen tellen niet)
    all_three_nee = all(is_nee(v) for v in vals) and all(str(v).strip() != "" for v in vals)

    if all_three_nee:
        # Verwijder eventuele 'Klaar om te sluiten?'
        bullets = [b for b in bullets if "Klaar om te sluiten?" not in b]
        # Voeg nieuwe conclusie toe
        bullets.append("• Gesloten SO, project sluiten na goedkeuring")

    return "\n".join(bullets)


def make_actiepunten_elders(row):
    """
    Actiepunten Elders:
    - Openstaande bestelling == Ja  → 'Gesloten SO met openstaande bestelling' (géén Proto/Prod)
    - Openstaande PO == Ja         → 'Gesloten SO met openstaande PO - <Proto/Prod>' o.b.v. Type
    - Niet toegewezen regels (niet leeg/niet 0) → 'Orderregel(s) toewijzen aan PR'
    """
    # Normaliseer input
    val_best = str(row.get("Openstaande bestelling", "")).strip().lower()
    val_po   = str(row.get("Openstaande PO", "")).strip().lower()

    yes_best = val_best in {"ja", "true", "1", "y", "yes"}
    yes_po   = val_po   in {"ja", "true", "1", "y", "yes"}

    t = str(row.get("Type", "")).strip().lower()
    proto_types = {"orders handel", "orders projecten", "service orders"}
    prod_types  = {"orders kabelafdeling", "orders kastenbouw asml"}

    bullets = []

    # 1) Géén onderscheid meer voor 'Openstaande bestelling'
    if yes_best:
        bullets.append("• Gesloten SO met openstaande bestelling")

    # 2) 'Openstaande PO' met Proto/Prod
    if yes_po:
        if t in proto_types:
            bullets.append("• Gesloten SO met openstaande PO - Proto")
        elif t in prod_types:
            bullets.append("• Gesloten SO met openstaande PO - Prod")
        # Anders geen bullet voor PO bij onbekend type

    # 3) NIEUW: Niet toegewezen regels => altijd toevoegen zodra er iets staat (of > 0)
    ntr = row.get("Niet toegewezen regels")
    has_ntr = False
    if pd.notna(ntr):
        if isinstance(ntr, (int, float)):
            has_ntr = ntr != 0
        else:
            s = str(ntr).strip()
            has_ntr = (s != "") and (s not in {"0", "0.0"})
    if has_ntr:
        bullets.append("• Orderregel(s) toewijzen aan PR")

    return "\n".join(bullets)

def main():
    today = date.today()
    week = today.isocalendar()[1]
    year = today.year

    from pathlib import Path

    # Zoek nieuwste centrale bestand
    files = list(Path(r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Output").glob(
        "Overzicht_Projectadministratie_Week*.xlsx"
    ))
    if not files:
        raise FileNotFoundError("Geen centrale overzichten gevonden in Output.")
    path = max(files, key=lambda f: f.stat().st_mtime)
    print(f"📄 Gekozen centraal bestand: {path.name}")

    # 1) Lees het overzicht
    df = pd.read_excel(path, header=0, engine="openpyxl")

    # Voeg kolom toe als hij niet bestaat
    if 'Handmatig verwacht resultaat' not in df.columns:
        df['Handmatig verwacht resultaat'] = False

    # 3) Bereken conclusies (automatic)
    df["Actiepunten Projectleider"] = df.apply(lambda r: make_actions_projectleider(r, today), axis=1)
    df["Bespreekpunten"] = df.apply(make_bespreekpunten, axis=1)
    df["Actiepunten Elders"] = df.apply(make_actiepunten_elders, axis=1)


    # 4) Schrijf terug naar Excel
    wb = load_workbook(path)
    ws = wb["Overzicht"]

    # Vind bestaande kolommen (header rij 1)
    col_letters = {}
    for cell in ws[1]:
        if cell.value:
            col_letters[str(cell.value).strip()] = cell.column_letter

    # Als 'Bespreekpunten' nog niet bestaat, voeg header toe op het einde
    if "Bespreekpunten" not in col_letters:
        last_col_idx = ws.max_column + 1
        header_cell = ws.cell(row=1, column=last_col_idx)
        header_cell.value = "Bespreekpunten"
        col_letter_new = header_cell.column_letter
        col_letters["Bespreekpunten"] = col_letter_new
        ws.column_dimensions[col_letter_new].width = 50  # startbreedte

    # Update Actiepunten Projectleider
    if "Actiepunten Projectleider" in col_letters:
        col_letter = col_letters["Actiepunten Projectleider"]
        for idx, text in enumerate(df["Actiepunten Projectleider"], start=2):
            cell = ws[f"{col_letter}{idx}"]
            if text:
                cell.value = text
                cell.alignment = Alignment(wrap_text=True)
                lines = text.count("\n") + 1
                # houd bestaande hoogte aan als die groter is
                current_h = ws.row_dimensions[idx].height or 0
                ws.row_dimensions[idx].height = max(current_h, lines * 15)

    # Update Bespreekpunten (automatic) — NIET Actiepunten Bram
    if "Bespreekpunten" in col_letters:
        col_letter = col_letters["Bespreekpunten"]
        for idx, text in enumerate(df["Bespreekpunten"], start=2):
            cell = ws[f"{col_letter}{idx}"]
            if text:
                cell.value = text
                cell.alignment = Alignment(wrap_text=True)
                lines = text.count("\n") + 1
                current_h = ws.row_dimensions[idx].height or 0
                ws.row_dimensions[idx].height = max(current_h, lines * 15)

    
    # Update Actiepunten Elders
    if "Actiepunten Elders" in col_letters:
        col_letter = col_letters["Actiepunten Elders"]
        for idx, text in enumerate(df["Actiepunten Elders"], start=2):
            if text:
                cell = ws[f"{col_letter}{idx}"]
                cell.value = text
                cell.alignment = Alignment(wrap_text=True)
                lines = text.count("\n") + 1
                current_h = ws.row_dimensions[idx].height or 0
                ws.row_dimensions[idx].height = max(current_h, lines * 15)


    # Maak kolommen breder
    for col_name in ["Actiepunten Projectleider", "Bespreekpunten"]:
        if col_name in col_letters:
            ws.column_dimensions[col_letters[col_name]].width = 50

    wb.save(path)
    print(f"✅ Conclusies bijgewerkt in bestand: {os.path.basename(path)}")

if __name__ == "__main__":
    main()
