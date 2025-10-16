import os # Loads Python's Operating System module. Lets it get access to file and folder operations. 
import glob # Loads a module that finds file paths matching patterns xlsx, csv. Note: this isnt used so might be removed. 
import pandas as pd # Panda is a data librarty in Python. It provides a DataFrame (Spreadsheet-like structure). It enables data modification like excel/ SQL
import logging # Enables the script to log info, bugs, errors, warnings etc. 

# ───────────────────────────────────────────────────────────
# PAS JE PADEN HIER AAN
INPUT_FOLDER = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Input" # The location to where it checks input files
CENTRAL_FILE = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Overzicht Projectadministratie.xlsx" # The script will use this as lay-out. Better to look into this! 
MAPPING_FILE = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Kolommenmapping per bron.xlsx" # 
OUTPUT_FOLDER = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Output" # The location to where it puts new files> ()
# ───────────────────────────────────────────────────────────

logging.basicConfig(level=logging.DEBUG, format="%(message)s") # Dit zorgt er voor dat de loggings correct gaat. 
log = logging.getLogger()


def normalize_idx(x):
    num = pd.to_numeric(x, errors="coerce")
    return str(int(num)) if pd.notna(num) and float(num).is_integer() else str(x) # Hulpt functie die projectnummers in de juiste format zet, zodat het leesbaar is. 


# Handmatige kolommen (nooit overschrijven met input)
HANDMATE_COLUMNS = [
    "Algemene informatie",
    "Bespreekpunten",         
    "Actiepunten Bram",       
    "Verwacht resultaat"
]

FLAG_COL = "Handmatig verwacht resultaat" # Een variable die die later gebruikt wordt om "Verwacht resultaat" te berekenen.


def normalize_manual_columns(df: pd.DataFrame) -> pd.DataFrame: # Zet oude datasheet om in nieuwe datasheet, zonder gegevens te verliezen. 
    
    cols = {c.strip(): c for c in df.columns}
    if "Actiepunten Overig" in cols and "Actiepunten Bram" not in cols:
        df.rename(columns={"Actiepunten Overig": "Actiepunten Bram"}, inplace=True) # Zet Actiepunten Overig om in Actiepunten Bram. Onnodig als het in het beginsel al goed is. 

    cols = {c.strip(): c for c in df.columns}
    if "Actiepunten Bram" in cols and "Bespreekpunten" not in cols:
        df.rename(columns={"Actiepunten Bram": "Bespreekpunten"}, inplace=True) # Zet Actiepunten Bram om in Bespreekpunten. Onnodig als het in het beginsel al goed is. 
    return df


# 1) Mappingslijst (overige velden). Handmatige kolommen vangen we via normalize_manual_columns(). MAPPINGS ZIJN NIET MEER NODIG, WANT DIE WORDEN AL BEPAALD IN EXCEL FILE 
mappings = {
    "Projectoverzicht Sumatra": {
        "Project": "Projectnummer",
        "Omschrijving": "Omschrijving",
        "Projectleider": "Projectleider",
        "Klant": "Klant",
        "Einddatum": "Einddatum",
        "Bud.Kost.": "Budget Kosten",
        "Bud.Opbr.": "Budget Opbrengsten",
        "Kosten": "Werkelijke kosten",
        "Opbrengsten": "Werkelijke opbrengsten",
        "Selcode": "Selcode",
        "Volg.lev.dat.": "Leverdatum"
    },
    "Verkoopdummy Sumatra": {
        "Order": "Projectnummer",
        "Niet toegewezen regel(s)": "Niet toegewezen regels",
        "Niet toegewezen": "Niet toegewezen regels"
    },
    "Werkbestand Projectadministratie": {
        "Project": "Projectnummer",
        "Omschrijving": "Omschrijving",
        "Verwacht resultaat": "Verwacht resultaat",
        "Actiepunten projectleider": "Actiepunten Projectleider",
        # Oude namen fallback (worden daarna alsnog genormaliseerd):
        "Actiepunten Bram": "Actiepunten Bram",
        "Algemene informatie": "Algemene informatie",
        "Bespreekpunten": "Bespreekpunten"
    },
    "Overzicht Projectadministratie": {
        "Projectnummer": "Projectnummer",
        "Omschrijving": "Omschrijving",
        "Projectleider": "Projectleider",
        "Klant": "Klant",
        "Einddatum": "Einddatum",
        "B. Kosten": "Budget Kosten",
        "B. Opbrengst": "Budget Opbrengsten",
        "W. Kosten": "Werkelijke kosten",
        "W. Opbrengst": "Werkelijke opbrengsten",
        "Leverdatum": "Leverdatum",
        "Niet toegewezen": "Niet toegewezen regels",
        "Verwacht resultaat": "Verwacht resultaat",
        "Actiepunten Projectleider": "Actiepunten Projectleider",
        # Breng oude kolommen op nieuwe standaard
        "Actiepunten Bram": "Actiepunten Bram",
        "Bespreekpunten": "Bespreekpunten",
        "Whitelist ": "Whitelist ",
        "Algemene informatie": "Algemene informatie",
        "Versielog ": "Versielog",
        "Type": "Type"
    }
}

for key in mappings:
    mappings[key] = {k.strip(): v.strip() for k, v in mappings[key].items()} # << deze code is onnodig als de mappings via Mappings_File gaat. 


# 2) Lees & standaardiseer CENTRAL
df_central = pd.read_excel(CENTRAL_FILE, header=1) # Leest de CENTRAL_FILE, pakt de 2e rij.
df_central.columns = df_central.columns.str.strip() # Formatting lijn. Zorgt dat spaties worden genegeerd.
df_central.rename(columns=mappings["Overzicht Projectadministratie"], inplace=True) # Hier pakt ie de hardcoded mappings ipv die van de centrale file. Deze zou wegkunnen. 
df_central = normalize_manual_columns(df_central) # Past de code toe in normalize manual codes zodat overige actiepunten worden omgezet naar actiepunten bram etc. Kan weg. 
df_central.set_index("Projectnummer", inplace=True) # Wordt de index van de data. Elk project wordt uniek geidentificeerd door dit. 
central_cols = df_central.columns.tolist() # Bewaard de huidige kolom namen van de centrale file in een lijst. 

# Zorg dat de statuskolom uit central aanwezig blijft
if FLAG_COL not in df_central.columns:
    df_central[FLAG_COL] = False
if FLAG_COL not in central_cols:
    central_cols.append(FLAG_COL)


# 3) Input-kolommen uit mappingbestand (alleen niet-handmatig)
mapping_df = pd.read_excel(MAPPING_FILE, header=0)
mapping_df.rename(columns={mapping_df.columns[0]: "Header"}, inplace=True)
mapping_df.columns = mapping_df.columns.str.strip()
mask = mapping_df["Soort"].str.strip().str.lower() == "input"
input_cols = mapping_df.loc[mask.fillna(False), "Header"].dropna().tolist()
if "Selcode" not in input_cols:
    input_cols.append("Selcode")


# 4) Vind 2 nieuwste .xlsx in INPUT
files = [f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith(".xlsx") and not f.startswith("~$")]
files.sort(key=lambda f: os.path.getmtime(os.path.join(INPUT_FOLDER, f)), reverse=True)
latest = files[:2]


# 5) Verwerk bronnen
all_inputs = []
dummy_list = []

for fname in latest:
    fp = os.path.join(INPUT_FOLDER, fname)
    df = pd.read_excel(fp, header=1)
    df.columns = df.columns.str.strip()
    cols_l = df.columns.str.lower().str.strip()

    if "bud.kost." in cols_l and ("projectleider" in cols_l or "selcode" in cols_l):
        typ = "Projectoverzicht Sumatra"
    elif any("niet toegewezen" in c for c in cols_l):
        typ = "Verkoopdummy Sumatra"
    elif "verwacht resultaat" in cols_l and ("actiepunten bram" in cols_l or "bespreekpunten" in cols_l):
        typ = "Werkbestand Projectadministratie"
    elif "versielog" in cols_l:
        typ = "Overzicht Projectadministratie"
    else:
        continue

    sub = df[[c for c in mappings[typ] if c in df.columns]].rename(columns=mappings[typ])
    sub = normalize_manual_columns(sub)

    if "Projectnummer" not in sub.columns:
        continue

    sub.set_index("Projectnummer", inplace=True)
    sub.index = sub.index.map(normalize_idx)
    sub.index.name = "Projectnummer"

    if typ == "Verkoopdummy Sumatra":
        dummy_list.append(sub[["Niet toegewezen regels"]])
    else:
        sel = [c for c in input_cols if c != "Projectnummer"]
        for col in HANDMATE_COLUMNS:
            if col in sub.columns and col not in sel:
                sel.append(col)
        all_inputs.append(sub.reindex(columns=sel, fill_value=pd.NA))


# 6) Combineer inputs
if all_inputs:
    df_input = pd.concat(all_inputs, axis=0)
else:
    df_input = pd.DataFrame(columns=input_cols + HANDMATE_COLUMNS)

# Bescherm handmatige kolommen: input mag ze niet overschrijven
for col in HANDMATE_COLUMNS:
    if col in df_input.columns:
        df_input[col] = pd.NA


# 7) Dummyregels updaten
if dummy_list:
    df_dummy = pd.concat(dummy_list, axis=0)
    df_dummy.index = df_dummy.index.map(normalize_idx)
    df_dummy.index.name = "Projectnummer"
    df_input.update(df_dummy)

exist = df_input.index.intersection(df_central.index)
new = df_input.index.difference(df_central.index)


# 8.1) Update alleen NIET-handmatige kolommen (en laat 'Verwacht resultaat' hier buiten)
cols_to_update = [c for c in df_input.columns if c not in HANDMATE_COLUMNS and c != "Verwacht resultaat"]
if not exist.empty and cols_to_update:
    df_central.loc[exist, cols_to_update] = df_input.loc[exist, cols_to_update]

# 8.2) 'Verwacht resultaat' alléén updaten waar GEEN handmatige lock (FLAG_COL) staat
if "Verwacht resultaat" in df_input.columns and "Verwacht resultaat" in df_central.columns:
    mask_locked = df_central.loc[exist, FLAG_COL].astype(bool)
    can_update = exist[~mask_locked]
    if len(can_update) > 0:
        df_central.loc[can_update, "Verwacht resultaat"] = df_input.loc[can_update, "Verwacht resultaat"]

# 8.3) Handmatige kolommen: alleen vullen als input echt waarde heeft
for col in HANDMATE_COLUMNS:
    if col in df_input.columns and col in df_central.columns:
        for idx in exist:
            val = df_input.at[idx, col]
            if pd.notna(val) and str(val).strip():
                df_central.at[idx, col] = val


# 9) Voeg nieuwe projecten toe
df_merged = pd.concat([df_central, df_input.loc[new]], axis=0).reset_index()


# 10) Type toevoegen (optioneel)
if "Selcode" in df_merged.columns:
    df_merged["Type"] = df_merged["Selcode"].apply(
        lambda x: "Production" if str(x).strip() == "Orders Kabelafdeling" else "Proto"
    )
if "Type" not in central_cols:
    central_cols.append("Type")


# 11) Kolomvolgorde herstellen
cols_final = ["Projectnummer"] + [c for c in central_cols if c in df_merged.columns]
df_merged = df_merged[cols_final]


# 12) Bereken Verwacht resultaat als leeg
if "Budget Opbrengsten" in df_merged.columns and "Budget Kosten" in df_merged.columns:
    mask = (
        (df_merged["Verwacht resultaat"].isna() | (df_merged["Verwacht resultaat"] == "")) &
        (~df_merged[FLAG_COL].astype(bool))
    )
    df_merged.loc[mask, "Verwacht resultaat"] = (
        pd.to_numeric(df_merged.loc[mask, "Budget Opbrengsten"], errors="coerce").fillna(0) -
        pd.to_numeric(df_merged.loc[mask, "Budget Kosten"], errors="coerce").fillna(0)
    )


# 13) Datumformatting
for col in ["Einddatum", "Eerstvolgende leverdatum", "Leverdatum"]:
    if col in df_merged.columns:
        df_merged[col] = pd.to_datetime(df_merged[col], errors="coerce").dt.strftime("%Y-%m-%d")


# 14) Versielog
if "Versielog" in df_merged.columns:
    df_merged["Versielog"] = pd.Timestamp.today().strftime("%Y-%m-%d")


# 15) Valutavelden afronden/opmaken
currency_cols = [
    "Budget Kosten", "Budget Opbrengsten",
    "Werkelijke kosten", "Werkelijke opbrengsten",
    "Verwacht resultaat"
]

for col in currency_cols:
    if col in df_merged.columns:
        df_merged[col] = pd.to_numeric(df_merged[col], errors="coerce").fillna(0).round(0).astype(int)


# 16) Export (zonder werkbestand-/oranje-logica)
dt = pd.Timestamp.today()
fn = f"Overzicht_Projectadministratie_Week{dt.week}_{dt.year}.xlsx"
out_path = os.path.join(OUTPUT_FOLDER, fn)

with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
    df_merged.to_excel(writer, index=False, sheet_name="Overzicht")
    wb = writer.book
    ws = writer.sheets["Overzicht"]

    # formats
    money_fmt = wb.add_format({"num_format": "€#,##0"})
    header_fmt = wb.add_format({
        "bold": True,
        "font_color": "white",
        "bg_color": "#4F81BD",
        "align": "center",
        "valign": "vcenter",
        "font_size": 12
    })

    # write headers
    for i, col in enumerate(df_merged.columns):
        ws.write(0, i, col, header_fmt)

    ws.set_row(0, 25)
    max_row, max_col = df_merged.shape
    ws.autofilter(0, 0, max_row, max_col - 1)
    ws.freeze_panes(1, 0)

    # adjust widths and apply money fmt
    for i, col in enumerate(df_merged.columns):
        try:
            width = max(df_merged[col].astype(str).map(len).max(), len(col)) + 2
        except ValueError:
            width = len(col) + 2
        fmt = money_fmt if col in currency_cols else None
        ws.set_column(i, i, width, fmt)

print(f"Klaar, opgeslagen als: {fn}")
