import os
import pandas as pd

# ───────────────────────────────────────────────────────────
# PAS JE PADEN AAN
INPUT_FOLDER  = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Input"
CENTRAL_FILE  = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Overzicht Projectadministratie.xlsx"
MAPPING_FILE  = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Kolommenmapping per bron.xlsx"
OUTPUT_FOLDER = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Output"
# ───────────────────────────────────────────────────────────


# 1) Je volledige mappings-dict uit v0.02
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
        "Order": "Projectnummer",                     # ← nieuw
        "Niet toegewezen regel(s)": "Niet toegewezen regels",
        "Niet toegewezen": "Niet toegewezen regels"
    },
    "Werkbestand Projectadministratie": {
        "Project": "Projectnummer",                   # ← belangrijk
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



# 2) Inlezen van het centrale bestand
df_central = pd.read_excel(CENTRAL_FILE, header=1)
df_central.set_index("Projectnummer", inplace=True)
print("✅ Centrale bestand ingelezen:", df_central.shape)


# 3) Uit mapping-Excel bepalen welke standaard-kolommen Input zijn
mapping_df = pd.read_excel(MAPPING_FILE, header=0)
mapping_df = mapping_df.rename(columns={mapping_df.columns[0]: "Header"})
mapping_df.columns = mapping_df.columns.str.strip()
print("🔍 Mapping-kolommen:", mapping_df.columns.tolist())

# Pak alle waarden uit "Header" waar Soort == Input
input_cols = (
    mapping_df
      .loc[mapping_df["Soort"].str.strip().str.lower() == "input", "Header"]
      .dropna()
      .tolist()
)
print("✅ Te updaten kolommen:", input_cols)


# 4) Vind de 2 nieuwste .xlsx-bestanden
files = [os.path.join(INPUT_FOLDER, f)
         for f in os.listdir(INPUT_FOLDER)
         if f.lower().endswith(".xlsx")]
files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
latest_files = files[:2]
print("🔄 In te lezen bestanden:", [os.path.basename(f) for f in latest_files])


# 5) Verwerk elk bestand: herkennen, hernoemen, indexeren én selecteren
all_inputs = []
for fp in latest_files:
    print(f"\n--- Reading file: {os.path.basename(fp)} ---")
    df = pd.read_excel(fp, header=1)
    print("📋 Kolommen:", df.columns.tolist())

    cols_lower = df.columns.str.lower().str.strip()
    if "bud.kost." in cols_lower and "projectleider" in cols_lower:
        bestandstype = "Projectoverzicht Sumatra"
    elif any("niet toegewezen" in c for c in cols_lower):
        bestandstype = "Verkoopdummy Sumatra"
    elif "verwacht resultaat" in cols_lower and "actiepunten bram" in cols_lower:
        bestandstype = "Werkbestand Projectadministratie"
    elif "versielog" in cols_lower:
        bestandstype = "Overzicht Projectadministratie"
    else:
        print(f"⚠️ Onbekend type, overslaan: {os.path.basename(fp)}")
        continue

    print("✅ Bestandstype:", bestandstype)
    kolom_mapping = mappings[bestandstype]

    # Keep alleen de kolommen die in jouw mapping-dict staan
    df = df[[c for c in kolom_mapping.keys() if c in df.columns]]
    df = df.rename(columns=kolom_mapping)

    # Zet Projectnummer op index
    df.set_index("Projectnummer", inplace=True)

    # **Hier de fix**: maak een DataFrame met precies de standaard 'input_cols'
    # - reindex voegt ontbrekende kolommen (zoals voor andere types) als NaN
    df_inputs = df.reindex(columns=[c for c in input_cols if c != "Projectnummer"])

    all_inputs.append(df_inputs)

# 6) Concat alle input-dataframes
df_input = pd.concat(all_inputs, axis=0)
print("\n✅ Samengevoegd input-data:", df_input.shape)


# 7) Update bestaande & voeg nieuwe projecten toe
existing = df_input.index.intersection(df_central.index)
new_idx   = df_input.index.difference(df_central.index)

df_central.update(df_input.loc[existing])
print(f"🔄 {len(existing)} bestaande projecten geüpdatet")

df_new    = df_input.loc[new_idx]
df_merged = pd.concat([df_central, df_new])
print(f"➕ {len(new_idx)} nieuwe projecten toegevoegd")


# 8) Reset index & sla op met weeknummer + jaar
df_merged.reset_index(inplace=True)
weeknr = pd.Timestamp.today().week
jaar   = pd.Timestamp.today().year
out_fn = f"Overzicht_Projectadministratie_Week{weeknr}_{jaar}.xlsx"
df_merged.to_excel(os.path.join(OUTPUT_FOLDER, out_fn), index=False)
print("✅ Bestand weggeschreven als:", out_fn)
