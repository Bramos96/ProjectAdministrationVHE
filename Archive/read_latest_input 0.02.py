import os
import pandas as pd

# 📂 Pad naar de Input-folder
input_folder = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Input"

# 📊 Kolommenmapping per bestandstype
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

# 📁 Haal alle .xlsx bestanden op
excel_files = [
    os.path.join(input_folder, f)
    for f in os.listdir(input_folder)
    if f.endswith('.xlsx')
]

# 🕒 Sorteer op laatste wijziging (nieuwste boven)
excel_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

# 🎯 Alleen de 2 nieuwste meenemen
latest_files = excel_files[:2]

# 🔄 Loop door bestanden
for file_path in latest_files:
    print(f"\n--- Reading file: {os.path.basename(file_path)} ---")
    try:
        # 📌 Lees bestand met kolomnamen op rij 2
        df = pd.read_excel(file_path, header=1)

        # 🔎 Toon kolommen die Python ziet
        print("📋 Kolomnamen in dit bestand:")
        for col in df.columns:
            print(f"- '{col}'")

        # 🔍 Herken bestandstype o.b.v. kolomnamen
        columns = df.columns.str.lower().str.strip()

        if "bud.kost." in columns and "projectleider" in columns:
            bestandstype = "Projectoverzicht Sumatra"
        elif any("niet toegewezen" in col for col in columns):
            bestandstype = "Verkoopdummy Sumatra"
        elif "verwacht resultaat" in columns and "actiepunten bram" in columns:
            bestandstype = "Werkbestand Projectadministratie"
        elif "versielog" in columns:
            bestandstype = "Overzicht Projectadministratie"
        else:
            print(f"⚠️ Onbekend bestandstype op basis van kolommen in {file_path}")
            continue

        # 🧠 Pas mapping toe
        kolom_mapping = mappings[bestandstype]
        df = df[[col for col in kolom_mapping.keys() if col in df.columns]]
        df = df.rename(columns=kolom_mapping)

        # ✅ Preview
        print(f"\n✅ Bestandstype herkend: {bestandstype}")
        print(df.head())

    except Exception as e:
        print(f"❌ Fout bij lezen van {file_path}: {e}")
