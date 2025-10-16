import os
import pandas as pd

# ───────────────────────────────────────────────────────────
# PAS JE PADEN HIER AAN
INPUT_FOLDER  = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Input"
CENTRAL_FILE  = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Overzicht Projectadministratie.xlsx"
MAPPING_FILE  = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Kolommenmapping per bron.xlsx"
OUTPUT_FOLDER = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Output"
# ───────────────────────────────────────────────────────────

# Helper voor index-normalisatie
def normalize_idx(x):
    num = pd.to_numeric(x, errors="coerce")
    if pd.notna(num) and float(num).is_integer():
        return str(int(num))
    return x

# 1) Mappings-dict + strip whitespace
mappings = {
    "Projectoverzicht Sumatra": {
        "Project": "Projectnummer", "Omschrijving ": "Omschrijving",
        "Projectleider": "Projectleider", "Klant ": "Klant",
        "Einddatum": "Einddatum", "Bud.Kost.": "Budget Kosten",
        "Bud.Opbr.": "Budget Opbrengsten", "Kosten": "Werkelijke kosten",
        "Opbrengsten": "Werkelijke opbrengsten", "Lst. leverdatum": "Eerstvolgende leverdatum"
    },
    "Verkoopdummy Sumatra": {
        "Order": "Projectnummer", "Niet toegewezen regel(s)": "Niet toegewezen regels",
        "Niet toegewezen": "Niet toegewezen regels"
    },
    "Werkbestand Projectadministratie": {
        "Project": "Projectnummer", "Omschrijving": "Omschrijving",
        "Verwacht resultaat": "Verwacht resultaat",
        "Actiepunten projectleider": "Actiepunten Projectleider",
        "Actiepunten Bram": "Actiepunten Bram ", "Algemene informatie": "Algemene informatie"
    },
    "Overzicht Projectadministratie": {
        "Projectnummer": "Projectnummer", "Omschrijving": "Omschrijving",
        "Projectleider": "Projectleider", "Klant": "Klant",
        "Einddatum": "Einddatum", "B. Kosten": "Budget Kosten",
        "B. Opbrengst": "Budget Opbrengsten", "W. Kosten": "Werkelijke kosten",
        "W. Opbrengst": "Werkelijke opbrengsten", "Leverdatum ": "Eerstvolgende leverdatum",
        "Niet toegewezen": "Niet toegewezen regels",
        "Verwacht resultaat": "Verwacht resultaat",
        "Actiepunten Projectleider": "Actiepunten Projectleider",
        "Actiepunten Bram": "Actiepunten Bram ", "Whitelist ": "Whitelist ",
        "Algemene informatie": "Algemene informatie", "Versielog ": "Versielog"
    }
}
for key, cmap in mappings.items():
    mappings[key] = {k.strip(): v.strip() for k, v in cmap.items()}

# 2) Lees & standaardiseer het centrale bestand
df_central = pd.read_excel(CENTRAL_FILE, header=1)
df_central.columns = df_central.columns.str.strip()
df_central.rename(columns=mappings["Overzicht Projectadministratie"], inplace=True)
df_central.set_index("Projectnummer", inplace=True)
central_cols = df_central.columns.tolist()

# 3) Bepaal input-kolommen uit mapping-Excel
mapping_df = pd.read_excel(MAPPING_FILE, header=0)
mapping_df.rename(columns={mapping_df.columns[0]: "Header"}, inplace=True)
mapping_df.columns = mapping_df.columns.str.strip()
mask = mapping_df["Soort"].str.strip().str.lower() == "input"
input_cols = mapping_df.loc[mask.fillna(False), "Header"].dropna().tolist()

# 4) Kies de 2 nieuwste echte .xlsx-bestanden
files = [f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith('.xlsx') and not f.startswith('~$')]
files.sort(key=lambda x: os.path.getmtime(os.path.join(INPUT_FOLDER, x)), reverse=True)
latest = files[:2]

# 5) Verwerk exports (dummy vs. overige)
all_inputs = []
dummy_list = []
for fname in latest:
    fp = os.path.join(INPUT_FOLDER, fname)
    df = pd.read_excel(fp, header=1)
    df.columns = df.columns.str.strip()
    cols_l = df.columns.str.lower().str.strip()

    if 'bud.kost.' in cols_l and 'projectleider' in cols_l:
        typ = 'Projectoverzicht Sumatra'
    elif any('niet toegewezen' in c for c in cols_l):
        typ = 'Verkoopdummy Sumatra'
    elif 'verwacht resultaat' in cols_l and 'actiepunten bram' in cols_l:
        typ = 'Werkbestand Projectadministratie'
    elif 'versielog' in cols_l:
        typ = 'Overzicht Projectadministratie'
    else:
        continue

    sub = df[[c for c in mappings[typ] if c in df.columns]].rename(columns=mappings[typ])
    sub.set_index('Projectnummer', inplace=True)
    sub.index = sub.index.map(normalize_idx)
    sub.index.name = 'Projectnummer'

    if typ == 'Verkoopdummy Sumatra':
        dummy_list.append(sub[['Niet toegewezen regels']])
    else:
        sel = [c for c in input_cols if c != 'Projectnummer']
        all_inputs.append(sub.reindex(columns=sel))

# 6) Concat overige inputs
df_input = pd.concat(all_inputs, axis=0)

# 7) Verwerk dummy
if dummy_list:
    df_dummy = pd.concat(dummy_list, axis=0)
    df_dummy.index = df_dummy.index.map(normalize_idx)
    df_dummy.index.name = 'Projectnummer'
    df_dummy = df_dummy[df_dummy.index.isin(df_input.index)]
    df_input.update(df_dummy)

# 8) Update bestaanden & voeg nieuwe toe
exist = df_input.index.intersection(df_central.index)
new = df_input.index.difference(df_central.index)
df_central.update(df_input.loc[exist])
df_merged = pd.concat([df_central, df_input.loc[new]])

# 9) Reset index & herorder kolommen
df_merged.reset_index(inplace=True)
df_merged = df_merged[['Projectnummer'] + central_cols]

# 10) Bereken Verwacht resultaat
if 'Budget Opbrengsten' in df_merged.columns and 'Budget Kosten' in df_merged.columns:
    df_merged['Verwacht resultaat'] = df_merged['Budget Opbrengsten'] - df_merged['Budget Kosten']

# 11) Format datums
for col in ['Einddatum', 'Eerstvolgende leverdatum']:
    if col in df_merged:
        df_merged[col] = pd.to_datetime(df_merged[col], errors='coerce').dt.strftime('%Y-%m-%d')

# 12) Houd numeriek voor valuta
currency_cols = ['Budget Kosten','Budget Opbrengsten','Werkelijke kosten','Werkelijke opbrengsten','Verwacht resultaat']
for col in currency_cols:
    if col in df_merged:
        df_merged[col] = df_merged[col].fillna(0).round(0).astype(int)

# 13) Schrijf weg met styling & auto-fit
dt = pd.Timestamp.today()
fn = f'Overzicht_Projectadministratie_Week{dt.week}_{dt.year}.xlsx'
out_path = os.path.join(OUTPUT_FOLDER, fn)
with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
    df_merged.to_excel(writer, index=False, sheet_name='Overzicht')
    wb = writer.book
    ws = writer.sheets['Overzicht']
    money_fmt = wb.add_format({'num_format':'€#,##0'})
    header_fmt = wb.add_format({'bold':True,'font_color':'white','bg_color':'#4F81BD','align':'center','valign':'vcenter','font_size':12})
    for i, col in enumerate(df_merged.columns):
        ws.write(0, i, col, header_fmt)
    ws.set_row(0, 25)
    max_row, max_col = df_merged.shape
    ws.autofilter(0, 0, max_row, max_col-1)
    ws.freeze_panes(1, 0)
    for i, col in enumerate(df_merged.columns):
        width = max(df_merged[col].astype(str).map(len).max(), len(col)) + 2
        fmt = money_fmt if col in currency_cols else None
        ws.set_column(i, i, width, fmt)
print('✅ Klaar, opgeslagen als:', fn)

