import os
import pandas as pd

# ───────────────────────────────────────────────────────────
# PAS JE PADEN HIER AAN
INPUT_FOLDER  = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Input"
CENTRAL_FILE  = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Overzicht Projectadministratie.xlsx"
MAPPING_FILE  = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Kolommenmapping per bron.xlsx"
OUTPUT_FOLDER = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Output"
# ───────────────────────────────────────────────────────────

# 1) mappings-dict (v0.02) + strip whitespace
mappings = {
    "Projectoverzicht Sumatra": {
        "Project": "Projectnummer",
        "Omschrijving ": "Omschrijving",
        "Projectleider": "Projectleider",
        "Klant ": "Klant",
        "Einddatum": "Einddatum",
        "Bud.Kost.": "Budget Kosten",
        "Bud.Opbr.": "Budget Opbrengsten",
        "Kosten": "Werkelijke kosten",
        "Opbrengsten": "Werkelijke opbrengsten",
        "Lst. leverdatum": "Eerstvolgende leverdatum"
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
        "Actiepunten Bram": "Actiepunten Bram ",
        "Algemene informatie": "Algemene informatie"
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
        "Leverdatum ": "Eerstvolgende leverdatum",
        "Niet toegewezen": "Niet toegewezen regels",
        "Verwacht resultaat": "Verwacht resultaat",
        "Actiepunten Projectleider": "Actiepunten Projectleider",
        "Actiepunten Bram": "Actiepunten Bram ",
        "Whitelist ": "Whitelist ",
        "Algemene informatie": "Algemene informatie",
        "Versielog ": "Versielog"
    }
}
# strip whitespace in keys & values
for key, cmap in mappings.items():
    mappings[key] = {k.strip(): v.strip() for k, v in cmap.items()}


# 2) Lees & standaardiseer het centrale bestand
df_central = pd.read_excel(CENTRAL_FILE, header=1)
df_central.columns = df_central.columns.str.strip()
df_central.rename(columns=mappings["Overzicht Projectadministratie"], inplace=True)
df_central.columns = df_central.columns.str.strip()
df_central.set_index("Projectnummer", inplace=True)
central_cols = df_central.columns.tolist()


# 3) Bepaal input-kolommen uit mapping-Excel
mapping_df = pd.read_excel(MAPPING_FILE, header=0)
mapping_df = mapping_df.rename(columns={mapping_df.columns[0]: "Header"})
mapping_df.columns = mapping_df.columns.str.strip()

mask = mapping_df["Soort"].str.strip().str.lower() == "input"
mask = mask.fillna(False)
input_cols = mapping_df.loc[mask, "Header"].dropna().tolist()


# 4) Kies de 2 nieuwste echte .xlsx-bestanden
files = [
    f for f in os.listdir(INPUT_FOLDER)
    if f.lower().endswith(".xlsx") and not f.startswith("~$")
]
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

    # herken type
    if "bud.kost." in cols_l and "projectleider" in cols_l:
        typ = "Projectoverzicht Sumatra"
    elif any("niet toegewezen" in c for c in cols_l):
        typ = "Verkoopdummy Sumatra"
    elif "verwacht resultaat" in cols_l and "actiepunten bram" in cols_l:
        typ = "Werkbestand Projectadministratie"
    elif "versielog" in cols_l:
        typ = "Overzicht Projectadministratie"
    else:
        continue

    cmap = mappings[typ]
    sub = df[[c for c in cmap if c in df.columns]].rename(columns=cmap)
    sub.set_index("Projectnummer", inplace=True)

    if typ == "Verkoopdummy Sumatra":
        # Alleen de Niet toegewezen regels
        dummy_list.append(sub[["Niet toegewezen regels"]])
    else:
        # Input-kolommen selecteren
        sel = [c for c in input_cols if c != "Projectnummer"]
        all_inputs.append(sub.reindex(columns=sel))


# 6) Concat overige inputs
df_input = pd.concat(all_inputs, axis=0)

# 7) Verwerk dummy-list via isin()
if dummy_list:
    df_dummy = pd.concat(dummy_list, axis=0)
    df_dummy = df_dummy[df_dummy.index.isin(df_input.index)]
    df_input.update(df_dummy)


# 8) Update bestaanden & voeg nieuwe toe
exist = df_input.index.intersection(df_central.index)
new   = df_input.index.difference(df_central.index)

df_central.update(df_input.loc[exist])
df_merged = pd.concat([df_central, df_input.loc[new]])


# 9) Reset index & herorder kolommen
df_merged.reset_index(inplace=True)
final_cols = ["Projectnummer"] + central_cols
df_merged = df_merged[final_cols]


# 10) Format datums & euro’s
for col in ["Einddatum", "Eerstvolgende leverdatum"]:
    if col in df_merged:
        df_merged[col] = pd.to_datetime(df_merged[col], errors="coerce").dt.strftime("%Y-%m-%d")

for col in ["Budget Kosten", "Budget Opbrengsten", "Werkelijke kosten", "Werkelijke opbrengsten"]:
    if col in df_merged:
        df_merged[col] = (
            df_merged[col]
              .fillna(0)
              .round(0)
              .astype(int)
              .map(lambda x: f"€{x:,}".replace(",", "."))
        )


# 11) Save with week + year
weeknr = pd.Timestamp.today().week
jaar   = pd.Timestamp.today().year
fn     = f"Overzicht_Projectadministratie_Week{weeknr}_{jaar}.xlsx"
df_merged.to_excel(os.path.join(OUTPUT_FOLDER, fn), index=False)

print("✅ Done:", fn)
