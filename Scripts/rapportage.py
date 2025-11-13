import os
import sqlite3
from datetime import date
import pandas as pd

# ─────────────────────────────────────────────
# PADEN AANPASSEN NAAR JOUW SITUATIE
# ─────────────────────────────────────────────
OUTPUT_FOLDER = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Output"
DB_FILE       = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\projectadmin.db"

def find_latest_overview(folder: str) -> str:
    """Zoekt het nieuwste Overzicht-bestand in de Output-map."""
    files = [
        f for f in os.listdir(folder)
        if f.lower().startswith("overzicht") and f.lower().endswith(".xlsx")
    ]

    if not files:
        raise FileNotFoundError("Geen Overzicht-bestanden gevonden!")

    # Sorteer op bewerkingsdatum
    files = sorted(files, key=lambda f: os.path.getmtime(os.path.join(folder, f)), reverse=True)
    newest = os.path.join(folder, files[0])
    return newest


def load_overview(excel_path: str) -> pd.DataFrame:
    """Leest het Excel-overzicht in (tabblad 'Overzicht', koprij in rij 1)."""
    df = pd.read_excel(excel_path, sheet_name="Overzicht", header=0)
    return df


def write_snapshot(df_central: pd.DataFrame) -> None:
    """
    Schrijft een snapshot met:
    - snapshot_date  (datum script-run)
    - snapshot_week  (weeknummer van snapshot)
    - ALLE kolommen uit het projectoverzicht
    naar de tabel 'project_snapshots_full' in SQLite.
    """
    import sqlite3
    from datetime import date

    conn = sqlite3.connect(DB_FILE)

    today = date.today()
    snapshot_date = today.isoformat()
    snapshot_week = today.isocalendar().week  # ISO-weeknummer

    # Kopie van het hele overzicht
    df_snap = df_central.copy()

    # Extra meta-kolommen toevoegen
    # We zetten ze vooraan zodat ze makkelijk te vinden zijn in Power BI
    df_snap.insert(0, "snapshot_date", snapshot_date)
    df_snap.insert(1, "snapshot_week", snapshot_week)

    # Wegschrijven naar NIEUWE tabelnaam
    df_snap.to_sql("project_snapshots_full", conn, if_exists="append", index=False)

    conn.close()
    print(f"{len(df_snap)} regels opgeslagen in 'project_snapshots_full'.")



def main():
    print("🔍 Nieuwste Overzicht-bestand zoeken...")
    latest_file = find_latest_overview(OUTPUT_FOLDER)
    print(f"→ Geselecteerd bestand: {latest_file}")

    print("📖 Bestand inlezen...")
    df = load_overview(latest_file)

    print("💾 Wegschrijven naar database...")
    write_snapshot_to_csv(df)

    print("🎉 Klaar! Snapshot opgeslagen in projectadmin.db")

def write_snapshot_to_csv(df_central: pd.DataFrame) -> None:
    from datetime import date
    import os

    today = date.today()
    snapshot_date = today.isoformat()
    snapshot_week = today.isocalendar().week

    # Map voor snapshots
    SNAPSHOT_FOLDER = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Snapshots"
    os.makedirs(SNAPSHOT_FOLDER, exist_ok=True)

    df_snap = df_central.copy()
    df_snap.insert(0, "snapshot_date", snapshot_date)
    df_snap.insert(1, "snapshot_week", snapshot_week)

    output_file = os.path.join(SNAPSHOT_FOLDER, f"snapshot_{snapshot_date}.csv")

    df_snap.to_csv(output_file, index=False, sep=';')  # ; werkt beter voor NL Excel/Power BI

    print(f"📦 Snapshot opgeslagen als CSV: {output_file}")


if __name__ == "__main__":
    main()
