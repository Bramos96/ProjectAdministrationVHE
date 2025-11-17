import os
import pandas as pd
from datetime import date
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# ───────────────────────────────────────────
# PAD NAAR OUTPUT
# ───────────────────────────────────────────
OUTPUT_FILE = (
    r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration"
    r"\Output\Overzicht_Projectadministratie_Week{week}_{year}.xlsx"
)

# ───────────────────────────────────────────
# ACTIEPUNTEN PROJECTLEIDER
# ───────────────────────────────────────────
def make_actions_projectleider(row, today):
    bullets = []
    ed_ts = pd.to_datetime(row.get("Einddatum"), errors="coerce")
    if pd.notna(ed_ts) and today > ed_ts.date():
        bullets.append("• Einddatum verlopen")

    lv_ts = pd.to_datetime(row.get("Leverdatum"), errors="coerce")
    if pd.notna(lv_ts) and today > lv_ts.date():
        bullets.append("• Leverdatum(s) verlopen")

    bk = row.get("Budget Kosten")
    try:
        if pd.notna(bk) and float(bk) == 0:
            bullets.append("• Budget kosten toevoegen")
    except:
        pass

    bo = row.get("Budget Opbrengsten")
    try:
        if pd.notna(bo) and float(bo) == 0:
            bullets.append("• Budget opbrengsten toevoegen")
    except:
        pass

    return "\n".join(bullets)


# ───────────────────────────────────────────
# BESPREekpunten (ALLEEN deze twee!)
# ───────────────────────────────────────────
def make_bespreekpunten(row):
    bullets = []

    def to_float(x):
        try:
            return float(x)
        except:
            return None

    def is_yes(v):
        return str(v).strip().lower() in {"ja", "true", "1", "y", "yes"}

    # 1) Negatief resultaat bespreken
    vr = to_float(row.get("Verwacht resultaat"))
    if vr is not None and vr < 0:
        bullets.append("• Negatief resultaat bespreken")

    # 2) Dochterproject-indicator
    val_so = row.get("Openstaande SO", "")
    if is_yes(val_so):
        bullets.append("• Gesloten SO, maar openstaande SO dochterproject")

    return "\n".join(bullets)


# ───────────────────────────────────────────
# INFORMATIE (opbrengsten binnen + gesloten SO definitief)
# ───────────────────────────────────────────
def make_informatie(row):
    bullets = []

    def to_float(x):
        try:
            return float(x)
        except:
            return None

    def is_nee(v):
        s = str(v).strip().lower()
        return s in {"nee", "no", "n", "false", "0"}

    # 1) Opbrengsten binnen
    bo = to_float(row.get("Budget Opbrengsten"))
    wo = to_float(row.get("Werkelijke opbrengsten"))
    if bo is not None and wo is not None:
        if bo != 0 and bo <= wo:
            bullets.append("• Opbrengsten binnen")

    # 2) Gesloten SO → project sluiten
    cols_sc = ["Openstaande bestelling", "Openstaande SO", "Openstaande PO"]
    vals = [row.get(c, "") for c in cols_sc]
    all_three_nee = (
        all(is_nee(v) for v in vals)
        and all(str(v).strip() != "" for v in vals)
    )

    if all_three_nee:
        bullets.append("• Gesloten SO, project sluiten na goedkeuring")

    return "\n".join(bullets)


# ───────────────────────────────────────────
# PROTO/PROD
# ───────────────────────────────────────────
def make_proto_prod(row):
    t = str(row.get("Type", "")).strip().lower()

    mapping = {
        "former qnq customers": "Proto",
        "orders handel": "Proto",
        "orders kabelafdeling": "Prod",
        "orders kastenbouw asml": "Proto",
        "orders projecten": "Proto",
        "orders xt sets": "Prod",
        "proto": "Proto",
        "service orders": "Proto",
    }

    return mapping.get(t, "")


# ───────────────────────────────────────────
# ACTIEPUNTEN ELDERS
# ───────────────────────────────────────────
def make_actiepunten_elders(row):
    val_best = str(row.get("Openstaande bestelling", "")).strip().lower()
    val_po = str(row.get("Openstaande PO", "")).strip().lower()

    yes_best = val_best in {"ja", "true", "1", "y", "yes"}
    yes_po = val_po in {"ja", "true", "1", "y", "yes"}

    t = str(row.get("Type", "")).strip().lower()
    proto_types = {"orders handel", "orders projecten", "service orders"}
    prod_types  = {"orders kabelafdeling", "orders kastenbouw asml"}

    bullets = []

    if yes_best:
        bullets.append("• Gesloten SO met openstaande bestelling")

    if yes_po:
        if t in proto_types:
            bullets.append("• Gesloten SO met openstaande PO - Proto")
        elif t in prod_types:
            bullets.append("• Gesloten SO met openstaande PO - Prod")

    ntr = row.get("Niet toegewezen regels")
    if pd.notna(ntr):
        s = str(ntr).strip()
        if s not in {"", "0", "0.0"}:
            bullets.append("• Orderregel(s) toewijzen aan PR")

    return "\n".join(bullets)


# ───────────────────────────────────────────
# MAIN
# ───────────────────────────────────────────
def main():

    today = date.today()
    week = today.isocalendar()[1]
    year = today.year

    from pathlib import Path
    files = list(Path(r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Output").glob(
        "Overzicht_Projectadministratie_Week*.xlsx"
    ))
    if not files:
        raise FileNotFoundError("Geen centrale overzichten gevonden.")
    path = max(files, key=lambda f: f.stat().st_mtime)
    print(f"\n Gekozen bestand: {path.name}\n")

    df = pd.read_excel(path, header=0, engine="openpyxl")

    # BEREKEN KOLLOMMEN
    df["Actiepunten Projectleider"] = df.apply(lambda r: make_actions_projectleider(r, today), axis=1)
    df["Bespreekpunten"] = df.apply(make_bespreekpunten, axis=1)
    df["Informatie"] = df.apply(make_informatie, axis=1)
    df["Actiepunten Elders"] = df.apply(make_actiepunten_elders, axis=1)
    df["Proto/Prod"] = df.apply(make_proto_prod, axis=1)

    wb = load_workbook(path)
    ws = wb["Overzicht"]

    # HEADERS
    col_letters = {}
    for cell in ws[1]:
        if cell.value:
            col_letters[cell.value.strip()] = cell.column_letter

    # Voeg Informatie toe als header niet bestaat
    if "Informatie" not in col_letters:
        last = ws.max_column + 1
        ws.cell(row=1, column=last).value = "Informatie"
        col_letters["Informatie"] = ws.cell(row=1, column=last).column_letter
        ws.column_dimensions[col_letters["Informatie"]].width = 50

    # Voeg Proto/Prod toe indien nodig
    if "Proto/Prod" not in col_letters:
        last = ws.max_column + 1
        ws.cell(row=1, column=last).value = "Proto/Prod"
        col_letters["Proto/Prod"] = ws.cell(row=1, column=last).column_letter
        ws.column_dimensions[col_letters["Proto/Prod"]].width = 12

    # WEGSCHRIJVEN VAN ALLE KOLLOMMEN
    def write_column(name):
        if name in col_letters:
            col = col_letters[name]
            for idx, val in enumerate(df[name], start=2):
                ws[f"{col}{idx}"].value = val or None
                ws[f"{col}{idx}"].alignment = Alignment(wrap_text=True)

    write_column("Actiepunten Projectleider")
    write_column("Bespreekpunten")
    write_column("Informatie")
    write_column("Actiepunten Elders")
    write_column("Proto/Prod")

    # ZET INFORMATIE OP BREEDTE 50
    ws.column_dimensions[col_letters["Informatie"]].width = 50

    wb.save(path)
    print(f"\n  Script afgerond — gegevens opgeslagen in: {os.path.basename(path)}\n")


if __name__ == "__main__":
    main()
