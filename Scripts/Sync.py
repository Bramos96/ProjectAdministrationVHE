import subprocess
import os
import logging
from datetime import datetime

print("✅ sync.py is gestart!")


# ───────────────────────────────────────────────
# PAS PADEN AAN NAAR JOUW MAPPEN
SCRIPTS_DIR = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Scripts\Subs"
LOGS_DIR = r"C:\Users\bram.gerrits\Desktop\Automations\ProjectAdministration\Logs"
os.makedirs(LOGS_DIR, exist_ok=True)

# ───────────────────────────────────────────────

# Maak logs-dir aan als die nog niet bestaat
os.makedirs(LOGS_DIR, exist_ok=True)

# Maak log-filename met datum/tijd
now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
log_filename = os.path.join(LOGS_DIR, f"sync_log_{now}.txt")

# Configure logging
logging.basicConfig(
    filename=log_filename,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)

def run_script(script_name):
    """
    Run a Python script via subprocess and log output.
    """
    path = os.path.join(SCRIPTS_DIR, script_name)
    logging.info(f"➡ Start script: {script_name}")
    try:
        result = subprocess.run(
            ["python", path],
            capture_output=True,
            text=True,
            check=True
        )
        logging.info(result.stdout)
        logging.info(f"✅ Script {script_name} succesvol afgerond.")
    except subprocess.CalledProcessError as e:
        logging.error(f"❌ Fout in script: {script_name}")
        logging.error(e.stdout)
        logging.error(e.stderr)
        print(f"Fout in script {script_name}. Zie logbestand: {log_filename}")
        raise

def main():
    logging.info("🚀 START SYNC PROCESS")

    try:
        # 1. Upload werkbestand naar centrale bestand
        run_script("workingfile_to_overview.py")

        # 2. Lees inputfiles en werk centrale bestand bij
        run_script("read_latest_input.py")

        # 3. Bereken conclusies
        run_script("calculate_conclusions.py")

        # 4. Schrijf centrale bestand terug naar nieuw werkbestand
        run_script("overview_to_workingfile.py")

    except Exception as ex:
        logging.exception("❌ Er is een fout opgetreden tijdens het sync-proces.")
    else:
        logging.info("✅ SYNC-PROCES SUCCESVOL AFGEROND")

    print(f"✅ Sync-proces afgerond. Logbestand: {log_filename}")

if __name__ == "__main__":
    main()
