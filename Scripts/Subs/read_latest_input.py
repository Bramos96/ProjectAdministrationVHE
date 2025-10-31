import os # Loads Python's Operating System module. Lets it get access to file and folder operations. 
import pandas as pd # Panda is a data librarty in Python. It provides a DataFrame (Spreadsheet-like structure). It enables data modification like excel/ SQL
import logging # Enables the script to log info, bugs, errors, warnings etc. 
import xlsxwriter.utility as xl_utils

# ───────────────────────────────────────────────────────────
# PAS JE PADEN HIER AAN
INPUT_FOLDER = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Input" # The location to where it checks input files
CENTRAL_FILE = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Overzicht Projectadministratie.xlsx" # The script will use this as lay-out. 
MAPPING_FILE = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Kolommenmapping per bron.xlsx" # Overview of what columnheaders are used in which file in order to match. Also define if information should be imported or ignored. 
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
    "Verwacht resultaat",
    "Actiepunten Projectleider",
    "2e Projectleider",    
]

FLAG_COL = "Handmatig verwacht resultaat" # Een variable die die later gebruikt wordt om "Verwacht resultaat" te berekenen.


def normalize_manual_columns(df: pd.DataFrame) -> pd.DataFrame: # Zet oude datasheet om in nieuwe datasheet, zonder gegevens te verliezen. 
    
    return df

# 2) Lees layout uit central (kolomnamen, geen data)
df_layout = pd.read_excel(CENTRAL_FILE, header=0, nrows=0)
central_cols = df_layout.columns.str.strip().tolist()

# 2b) Zoek laatste Overzicht-bestand als geheugen
overview_files = [f for f in os.listdir(OUTPUT_FOLDER) if f.startswith("Overzicht_Projectadministratie_Week") and f.endswith(".xlsx")]
if overview_files:
    # sorteer op tijd, pak de nieuwste
    overview_files.sort(key=lambda f: os.path.getmtime(os.path.join(OUTPUT_FOLDER, f)), reverse=True)
    last_overview = os.path.join(OUTPUT_FOLDER, overview_files[0])

    df_central = pd.read_excel(last_overview, header=0)
    df_central.columns = df_central.columns.str.strip()
    # herindexeer naar de juiste layout
    df_central = df_central.reindex(columns=central_cols, fill_value=pd.NA)

    log.info(f"Geheugen geladen uit {last_overview}: {len(df_central)} projecten, {len(df_central.columns)} kolommen.")

else:
    log.warning("Geen vorig overzicht gevonden. Start met lege centrale dataset.")
    df_central = pd.DataFrame(columns=central_cols)

# 3) Input-kolommen uit mappingbestand (alleen niet-handmatig) # Deze code leest het mapping-bestand in en maakt daaruit een lijst met alle kolomnamen die als “Input” gemarkeerd zijn
mapping_df = pd.read_excel(MAPPING_FILE, header=0) 
mapping_df.columns = mapping_df.columns.str.strip()
mapping_df = mapping_df.applymap(lambda x: str(x).replace("\u00A0", " ").strip() if isinstance(x, str) else x)
mapping_df.columns = mapping_df.columns.str.strip()
mask = mapping_df["Soort"].str.strip().str.lower() == "input"
input_cols = mapping_df.loc[mask.fillna(False), "Header"].dropna().tolist()
# Safety: SC-kolommen altijd meenemen als input
for c in ["Openstaande bestelling", "Openstaande SO", "Openstaande PO"]:
    if c not in input_cols:
        input_cols.append(c)

if "Selcode" not in input_cols:
    input_cols.append("Selcode")

# 4) Vind 2 nieuwste .xlsx in INPUT
files = [f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith(".xlsx") and not f.startswith("~$")]
files.sort(key=lambda f: os.path.getmtime(os.path.join(INPUT_FOLDER, f)), reverse=True)
latest = files[:3]


# 5) Verwerk bronnen
all_inputs = []  # dit is een lege lijst voor de gewone inputtabellen.
dummy_list = []  # dit is een speciale verzameling voor "verkoopdummy" regels.

for fname in latest:  # hier worden de twee laatste bestanden in input gelezen en gestript.
    fp = os.path.join(INPUT_FOLDER, fname)
    df = pd.read_excel(fp, header=1)
    df.columns = df.columns.str.strip()
    cols_l = df.columns.str.lower().str.strip()

    log.info(f"--- Verwerken bestand: {fname} ---")
    log.info(f"Kolommen in bronbestand: {df.columns.tolist()}")

    norm = lambda s: "".join(str(s).lower().split())
    norm_cols = {norm(c) for c in df.columns}

    if "bud.kost." in cols_l and ("projectleider" in cols_l or "selcode" in cols_l):
        typ = "Projectoverzicht Sumatra"
    elif any("niet toegewezen" in c for c in cols_l):
        typ = "Verkoopdummy Sumatra"
    elif "verwacht resultaat" in cols_l and ("actiepunten bram" in cols_l or "bespreekpunten" in cols_l):
        typ = "Werkbestand Projectadministratie"
    elif "versielog" in cols_l:
        typ = "Overzicht Projectadministratie"
    elif (
        {"openstaandebestelling", "openstaandeorders", "openstaandepo"} <= norm_cols
        or {"openstaandebestelling", "openstaandeso", "openstaandepo"} <= norm_cols
        or {"besteld_open", "order_open", "prod_open"} <= norm_cols  # herken ook op tellerkolommen
    ):
        typ = "Overzicht Te Sluiten"
    else:
        continue

    # Selecteer de mapping voor dit bestandstype
    colmap = mapping_df[["Header", typ]].dropna()
    rename_dict = dict(zip(colmap[typ], colmap["Header"]))

    log.info(f"Mapping gebruikt voor bestandstype {typ}: {rename_dict}")
    log.info(f"Zoekt naar kolommen {list(rename_dict.keys())} in {fname}")

    #  Eerst renamen, dan filteren
    sub = df.rename(columns=rename_dict)

    # Neem alleen de kolommen die we in mapping hebben gebruikt (waarden, niet keys!)
    sub = sub[[col for col in rename_dict.values() if col in sub.columns]]

    # Debug: toon kolommen na mapping
    log.info(f"Kolommen in sub na mapping: {sub.columns.tolist()}")

    # Check dat Projectnummer er nu in zit
    if "Projectnummer" not in sub.columns:
        log.error(f" Projectnummer ontbreekt nog steeds in {fname}!")
        continue

    # Index instellen
    sub.set_index("Projectnummer", inplace=True)
    sub.index = sub.index.map(normalize_idx)
    sub.index.name = "Projectnummer"

    #  Zet de 3 SC-vlaggen op basis van tellerkolommen (>0 ⇒ 'Ja', anders 'Nee')
    if typ == "Overzicht Te Sluiten":
        src = df.copy()
        # sleutel gelijk trekken
        if "Projectcode" in src.columns:
            src["Projectnummer"] = src["Projectcode"].map(normalize_idx)
        elif "Projectnummer" in src.columns:
            src["Projectnummer"] = src["Projectnummer"].map(normalize_idx)
        else:
            src["Projectnummer"] = pd.NA

        num_map = {
            "Openstaande PO": "prod_open",
            "Openstaande SO": "order_open",
            "Openstaande bestelling": "besteld_open",
        }
        avail = {t: s for t, s in num_map.items() if s in src.columns}

        if avail:
            need = ["Projectnummer"] + list(avail.values())
            num = src[need].copy()
            for s in avail.values():
                num[s] = pd.to_numeric(num[s], errors="coerce").fillna(0)

            g = num.groupby("Projectnummer", dropna=False).sum(numeric_only=True)

            for target, src_col in avail.items():
                if target in sub.columns:
                    sub[target] = (
                        g[src_col]
                        .reindex(sub.index)
                        .fillna(0)
                        .gt(0)
                        .map({True: "Ja", False: "Nee"})
                        .astype("string")
                    )

            # logging als sanity check
            for c in ["Openstaande bestelling", "Openstaande SO", "Openstaande PO"]:
                if c in sub.columns:
                    log.info(f"[SC via tellers] {c}:\n{sub[c].value_counts(dropna=False)}")

    if typ == "Verkoopdummy Sumatra":
        if "Niet toegewezen regels" in sub.columns:
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

# Zorg dat Projectnummer als gewone kolom bestaat
if isinstance(df_input.index, pd.Index) and df_input.index.name == "Projectnummer":
    df_input = df_input.reset_index()

# === NIEUW: consolideren per Projectnummer ===
def pick_status(series: pd.Series):
    # Prioriteit: "Ja" wint > "Nee" > leeg
    s = series.dropna().astype(str).str.strip()
    if (s == "Ja").any():
        return "Ja"
    if (s == "Nee").any():
        return "Nee"
    return pd.NA

def pick_first_nonempty(series: pd.Series):
    for v in series:
        if pd.notna(v) and str(v).strip() != "":
            return v
    return pd.NA

STATUS_COLS = ["Openstaande bestelling", "Openstaande SO", "Openstaande PO"]

grouped_rows = []
for projnummer, chunk in df_input.groupby("Projectnummer"):
    row = {"Projectnummer": projnummer}
    for col in df_input.columns:
        if col == "Projectnummer":
            continue
        if col in STATUS_COLS:
            row[col] = pick_status(chunk[col])
        else:
            row[col] = pick_first_nonempty(chunk[col])
    grouped_rows.append(row)

df_input = pd.DataFrame(grouped_rows)
# === EINDE CONSOLIDATIE ===


# Zorg dat 'Verwacht resultaat' en de flagkolom altijd bestaan
if "Verwacht resultaat" not in df_input.columns:
    df_input["Verwacht resultaat"] = pd.NA

if FLAG_COL not in df_input.columns:
    df_input[FLAG_COL] = False

#  Zorg dat sleutelkolommen altijd correct zijn
# Projectnummer kan in bronnen "Project" heten
if "Project" in df_input.columns and "Projectnummer" not in df_input.columns:
    df_input.rename(columns={"Project": "Projectnummer"}, inplace=True)

# Selcode moet altijd naar Type → nooit dubbel bestaan
if "Selcode" in df_input.columns and "Type" not in df_input.columns:
    df_input.rename(columns={"Selcode": "Type"}, inplace=True)
elif "Selcode" in df_input.columns and "Type" in df_input.columns:
    df_input.drop(columns=["Selcode"], inplace=True)

# Debug: toon kolommen
log.info("Kolommen in df_input na sectie 6:")
for c in df_input.columns:
    log.info(f"- {repr(c)}")

# NA sectie 6, vlak na: if all_inputs: df_input = pd.concat(...)

log.info("="*60)
log.info("DEBUG: Aantal projecten per bron")
log.info("="*60)

for i, inp in enumerate(all_inputs):
    log.info(f"Input bron {i+1}: {len(inp)} projecten")
    log.info(f"  Index duplicaten: {inp.index.duplicated().sum()}")
    log.info(f"  Unieke projecten: {len(inp.index.unique())}")

log.info(f"\nTotaal in df_input NA concat: {len(df_input)} rijen")
log.info(f"Unieke projectnummers in df_input: {len(df_input['Projectnummer'].unique())}")
log.info("="*60)


# 7) Dummyregels updaten
if dummy_list:
    df_dummy = pd.concat(dummy_list, axis=0)
    df_dummy.index = df_dummy.index.map(normalize_idx)
    df_dummy.index.name = "Projectnummer"
    df_input.update(df_dummy)

exist = df_input.index
new = []  # geen nieuwe projecten, want central is leeg

# Maak sleutelkolom veilig
if "Projectnummer" not in df_input.columns:
    raise KeyError("Kolom 'Projectnummer' ontbreekt in df_input. Controleer mappingbestand!")

# Zet indexen en maak uniek
df_central = df_central.drop_duplicates(subset=["Projectnummer"], keep="last").set_index("Projectnummer").copy()
df_input = df_input.drop_duplicates(subset=["Projectnummer"], keep="last").set_index("Projectnummer").copy()

# Debug: zie wat er werkelijk in df_input staat
dbg_cols = [c for c in ["Openstaande bestelling","Openstaande SO","Openstaande PO"] if c in df_input.columns]
log.info(f"SC-kolommen in df_input: {dbg_cols}")
for c in dbg_cols:
    log.info(f"{c} value_counts:\n{df_input[c].value_counts(dropna=False)}")

# Vervang sectie 8 volledig door dit:

# --- BEGIN: BEHOUD ALLEEN PROJECTEN UIT INPUT ---
KEY = "Projectnummer"

df_central.index.name = KEY
df_input.index.name = KEY

# Welke projecten zitten er in de huidige input?
input_projecten = df_input.index.unique()

log.info(f"Projecten in input: {len(input_projecten)}")
log.info(f"Projecten in oud overzicht: {len(df_central)}")

# Stap 1: Gooi alle projecten weg die NIET in input staan
df_central = df_central.loc[df_central.index.intersection(input_projecten)].copy()

log.info(f"Projecten na filter (behouden): {len(df_central)}")

# Stap 2: Voeg nieuwe projecten uit input toe
nieuwe_projecten = input_projecten.difference(df_central.index)
if len(nieuwe_projecten) > 0:
    log.info(f"Nieuwe projecten toevoegen: {len(nieuwe_projecten)}")
    new_rows = pd.DataFrame(index=nieuwe_projecten, columns=df_central.columns)
    df_central = pd.concat([df_central, new_rows])

# Stap 3: Update automatische kolommen (NIET handmatige)

auto_cols = [c for c in df_input.columns if c not in HANDMATE_COLUMNS + ["Whitelist", FLAG_COL]]

# Zorg dat df_central alle kolommen kent
for c in auto_cols:
    if c not in df_central.columns:
        df_central[c] = pd.NA
for c in HANDMATE_COLUMNS:
    if c not in df_central.columns:
        df_central[c] = pd.NA

# 3a. Snapshot van alle handmatige kolommen zoals ze NU in df_central staan
manual_snapshot = {}
for col in HANDMATE_COLUMNS:
    if col in df_central.columns:
        manual_snapshot[col] = df_central[col].copy()

# 3b. Update alleen de automatische kolommen
present = [c for c in auto_cols if c in df_input.columns and c in df_central.columns]
if present:
    # Consistente types voor de statuskolommen
    for c in ["Openstaande bestelling", "Openstaande SO", "Openstaande PO"]:
        if c in df_input.columns:
            df_input[c] = df_input[c].astype("string")

    df_central.update(df_input[present], overwrite=True)

# 3c. Handmatige kolommen terugzetten zoals ze waren
for col, series in manual_snapshot.items():
    df_central[col] = series


log.info(f"Totaal projecten na sync: {len(df_central)}")
# --- END: BEHOUD ALLEEN PROJECTEN UIT INPUT ---

# 9) Definieer einddataset (nu inclusief geheugen + updates)
df_merged = df_central.reset_index()


# 10) Type toevoegen (optioneel)
if "Selcode" in df_merged.columns:
    df_merged["Type"] = df_merged["Selcode"].apply(
        lambda x: "Production" if str(x).strip() == "Orders Kabelafdeling" else "Proto"
    )

# 11) Kolomvolgorde herstellen
# Volgorde uit layoutbestand
layout_cols = df_layout.columns.str.strip().tolist()

# Extra kolommen die niet in layout zitten
extra_cols = []
if "Type" in df_merged.columns and "Type" not in layout_cols:
    extra_cols.append("Type")

for col in HANDMATE_COLUMNS + ["Whitelist"]:
    if col not in df_merged.columns:
        df_merged[col] = pd.NA
    if col not in layout_cols and col not in extra_cols:
        extra_cols.append(col)

# Bouw definitieve volgorde
cols_final = [c for c in layout_cols if c in df_merged.columns] + extra_cols

# Zorg dat Projectnummer altijd helemaal vooraan staat
if "Projectnummer" in cols_final:
    cols_final.remove("Projectnummer")
cols_final = ["Projectnummer"] + cols_final

df_merged = df_merged[cols_final]


# 12) Bereken Verwacht resultaat alleen als handmatig resultaat = Onwaar
if "Verwacht resultaat" in df_merged.columns and \
   "Budget Opbrengsten" in df_merged.columns and \
   "Budget Kosten" in df_merged.columns:

    # Zorg dat flag kolom booleans bevat
    df_merged[FLAG_COL] = df_merged[FLAG_COL].astype(bool).fillna(False)

    # Mask: alleen als leeg én niet handmatig
    mask = (
        (df_merged["Verwacht resultaat"].isna() | (df_merged["Verwacht resultaat"].astype(str).str.strip() == "")) &
        (~df_merged[FLAG_COL])
    )

    df_merged.loc[mask, "Verwacht resultaat"] = (
        pd.to_numeric(df_merged.loc[mask, "Budget Opbrengsten"], errors="coerce").fillna(0) -
        pd.to_numeric(df_merged.loc[mask, "Budget Kosten"], errors="coerce").fillna(0)
    )

# 12b) Bereken Werkelijk resultaat
if "Werkelijke opbrengsten" in df_merged.columns and "Werkelijke kosten" in df_merged.columns:

    if "Werkelijk resultaat" not in df_merged.columns:
        df_merged["Werkelijk resultaat"] = pd.NA

    # Alleen berekenen als er nog niets in staat
    mask_wr = (
        df_merged["Werkelijk resultaat"].isna() |
        (df_merged["Werkelijk resultaat"].astype(str).str.strip() == "")
    )

    df_merged.loc[mask_wr, "Werkelijk resultaat"] = (
        pd.to_numeric(df_merged.loc[mask_wr, "Werkelijke opbrengsten"], errors="coerce").fillna(0) -
        pd.to_numeric(df_merged.loc[mask_wr, "Werkelijke kosten"], errors="coerce").fillna(0)
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
    "Verwacht resultaat", "Werkelijk resultaat"
]

for col in currency_cols:
    if col in df_merged.columns:
        df_merged[col] = pd.to_numeric(df_merged[col], errors="coerce").fillna(0).round(0).astype(int)

# ⚔️ IMPORTANT CHANGE:
# Do NOT touch the flag column ("Handmatig verwacht resultaat").
# It remains exactly as entered manually in Excel or by your other tool.


# 16) Export
dt = pd.Timestamp.today()
fn = f"Overzicht_Projectadministratie_Week{dt.week}_{dt.year}.xlsx"
out_path = os.path.join(OUTPUT_FOLDER, fn)

print("CHECK >>>")
print("Openstaande PO value_counts:")
print(df_central["Openstaande PO"].value_counts(dropna=False).head())
print("Openstaande SO value_counts:")
print(df_central["Openstaande SO"].value_counts(dropna=False).head())
print("Openstaande bestelling value_counts:")
print(df_central["Openstaande bestelling"].value_counts(dropna=False).head())

print("\nVoorbeeld handmatige kolommen (niet leeg als het goed is):")
cols_check = ["Algemene informatie","Actiepunten Bram","2e Projectleider"]
cols_check = [c for c in cols_check if c in df_central.columns]
print(df_central[cols_check].dropna(how="all").head(5))
print("<<< END CHECK")


with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
    df_merged.to_excel(writer, index=False, sheet_name="Overzicht")
    wb = writer.book
    ws = writer.sheets["Overzicht"]

    # === Formats ===
    money_fmt = wb.add_format({"num_format": "€#,##0"})
    header_fmt = wb.add_format({
        "bold": True,
        "font_color": "white",
        "bg_color": "#4F81BD",
        "align": "center",
        "valign": "vcenter",
        "font_size": 12
    })
    orange_fmt = wb.add_format({"bg_color": "#FFA500", "font_color": "black"})

    # === Header row ===
    for i, col in enumerate(df_merged.columns):
        ws.write(0, i, col, header_fmt)

    ws.set_row(0, 25)  # taller header row
    max_row, max_col = df_merged.shape
    ws.autofilter(0, 0, max_row, max_col - 1)
    ws.freeze_panes(1, 0)

    # === Column widths & money formatting ===
    for i, col in enumerate(df_merged.columns):
        try:
            width = max(df_merged[col].astype(str).map(len).max(), len(col)) + 2
        except ValueError:
            width = len(col) + 2
        fmt = money_fmt if col in currency_cols else None
        ws.set_column(i, i, width, fmt)

    # === Conditional formatting ===
    if FLAG_COL in df_merged.columns and "Verwacht resultaat" in df_merged.columns:
        flag_col_idx = df_merged.columns.get_loc(FLAG_COL)
        target_col_idx = df_merged.columns.get_loc("Verwacht resultaat")
        flag_col_letter = xl_utils.xl_col_to_name(flag_col_idx)

        ws.conditional_format(
            1, target_col_idx, max_row, target_col_idx,
            {
                "type": "formula",
                "criteria": f'=${flag_col_letter}2=TRUE',
                "format": wb.add_format({"font_color": "orange", "bold": True})
            }
        )



print(f"Klaar, opgeslagen als: {fn}")


