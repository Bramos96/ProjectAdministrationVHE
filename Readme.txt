README – Projectadministratie Automatisering

📁 ARCHIVE  
Bevat oude versies van:  
- Overzicht Projectadministratie (back-ups)  
- Geëxporteerde overzichten per week  
Bestandsnaam bij voorkeur met datumnotatie (YYYY-MM-DD)  

📁 INPUT  
Hier worden nieuwe projectexports geplaatst (wekelijks of 2-wekelijks).  
Bijv. afkomstig uit Sumatra.  
Bestanden moeten .xlsx zijn met vaste kolomnamen.  

📁 OUTPUT  
Hier worden de gegenereerde overzichten geplaatst (mailbare versies).  
Bijv. Overzicht_Projectadministratie_Week24.xlsx  

📁 LOGS  
Scriptoutput komt hier terecht.  
Bijv. log_2025-06-01.txt met:  
- Aantal nieuwe projecten  
- Fouten  
- Overgeslagen projecten  

📁 OVERIG  
Voor losse bestanden, checklists of tijdelijke testjes.  
Bijv. opmerkingenbestand, kolommapping, brainstormnotities.  

📄 Overzicht Projectadministratie.xlsx  
Dit is het centrale projectbestand.  
Wordt continu aangevuld via script.  
Bevat per project: algemene info, actiepunten, status en historie.  

📄 Werkbestand Projectadministratie.xlsx  
Dummyversie van het centrale bestand.  
Gebruik dit bestand bij testen van scripts.  
Zorg dat je 'Overzicht Projectadministratie.xlsx' nooit direct laat overschrijven.  

📄 Dummy projectexport  
Plaats testdata in `/Input`, bijv. dummy_projectexport_2025-05-15.xlsx  
Gebruik dit om scripts veilig te testen zonder echte gegevens te beïnvloeden.  

Laatste update README: 15-05-2025

📄 read_latest_input 0.02.py  
Python-script dat automatisch de twee nieuwste `.xlsx` bestanden in de map `/Input` verwerkt.  
Het script herkent bestandstypes op basis van kolomnamen, past een kolommenmapping toe en toont een preview van de gestandaardiseerde data.

📁 INPUT – Aanvulling  
Bestandsnamen mogen wisselen (bijv. export(1).xlsx, export(42).xlsx).  
De herkenning gebeurt niet op basis van bestandsnaam, maar op basis van **kolomnamen op rij 2** (Excel-rij 2 = `header=1`).  
Bestanden mogen meerdere soorten zijn, waaronder:  
- Projectoverzicht Sumatra  
- Verkoopdummy Sumatra  
- Werkbestand Projectadministratie  
- Overzicht Projectadministratie

📁 SCRIPTS (optioneel, aan te maken map)  
Voor al je `.py` scripts zoals `read_latest_input 0.02`.  
Handig om los van de data te bewaren.

📄 Kolommenmapping per bron.xlsx  
Excelbestand dat de mapping bevat tussen kolomnamen in diverse bronnen en de standaardkolommen in het centrale bestand.  
Wordt gebruikt door het script om kolommen correct te hernoemen.