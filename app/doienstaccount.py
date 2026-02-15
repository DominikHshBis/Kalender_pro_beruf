from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timedelta
from dateutil import parser
from zoneinfo import ZoneInfo
import calendar
from Excel_eintrag import excel_setter
import json
from openpyxl import load_workbook
from pathlib import Path
from dotenv import load_dotenv
from pathlib import Path
import json
from openpyxl import load_workbook
import os

load_dotenv()

def find_file(filename: str, start_dir: Path) -> Path:
    for path in start_dir.rglob(filename): 
        return path 

PROJECT_ROOT = Path(__file__).resolve().parent.parent # Config automatisch finden

def env_path(var_name: str, default: str) -> Path:  # fallback version
    """Nimmt ENV-Wert oder Default relativ zum Projekt."""
    value = os.getenv(var_name) 
    if value: 
        return Path(value) 
    return find_file(default,PROJECT_ROOT)

CONFIG_PATH = env_path("CONFIG_PATH", "config_dominik.json") 
SERVICE_ACCOUNT_FILE = env_path("SERVICE_ACCOUNT_FILE", "credentials.json") 
EXCEL_LOAD_PATH = env_path("EXCEL_LOAD_PATH", "Muster_Honorarrechnung-Lehrkräfte_pytest.xlsx") 
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", PROJECT_ROOT / "output")) 
OUTPUT_DIR.mkdir(exist_ok=True)

# lädt json config
with open(CONFIG_PATH) as f:
    config = json.load(f)

#teilt congif inhalt den variablen zu
CALENDAR_ID = config["calendar_id"]
TAGS = config["tags"] #list of tags to search for in calendar event
SCOPES = [config["scopes"]] # api adress to access calendar data
#EXCEL_LOAD_PATH = config["excel_load_path"]

wb = load_workbook(EXCEL_LOAD_PATH) # excel laden
ws = wb.active # excel aktiv schalten
First_day = ""
Last_day = ""

#lädt die dredentials aus der json datei und gibt die Berechtigungen an, damit die API auf den Kalender zugreifen kann
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
) 
#erstellt einen Dienst, um mit der Google Calendar API zu kommunizieren
service = build("calendar", "v3", credentials=credentials)

# Aktuellen Monat berechnen
#now = datetime(2026,4,1) #bt jahr monat und Tag an

BERLIN = ZoneInfo("Europe/Berlin")
now = datetime.now(BERLIN)

first_local = datetime(now.year, now.month, 1, 0, 0, 0, tzinfo=BERLIN)
last_local = datetime(
    now.year,
    now.month,
    calendar.monthrange(now.year, now.month)[1],
    23, 59, 59,
    tzinfo=BERLIN
)

first_utc = first_local.astimezone(ZoneInfo("UTC"))
last_utc = last_local.astimezone(ZoneInfo("UTC"))

time_min = first_utc.isoformat().replace("+00:00", "Z")
time_max = last_utc.isoformat().replace("+00:00", "Z")

events_result = service.events().list(
    calendarId=CALENDAR_ID,
    timeMin=time_min,
    timeMax=time_max,
    singleEvents=True,
    orderBy="startTime",
    timeZone = "Europe/Berlin"
).execute()

events = events_result.get("items", []) # gibt die Termine zurück, die im aktuellen Monat liegen als Liste von Ereignissen zurück. Jedes Ereignis enthält Informationen wie Start- und Endzeit, Titel, Beschreibung usw.
# wenn in der Liste der Ereignisse keine Termine gefunden werden, wird eine Nachricht ausgegeben, dass keine Termine in diesem Monat vorhanden sind. Andernfalls wird für jedes Ereignis in der Liste eine Schleife durchlaufen, um die relevanten Informationen zu extrahieren und in die Excel-Datei einzutragen.
if not events:
    print("Keine Termine in diesem Monat.")
else:
    i = 32
    for event in events:
           
        start = event["start"].get("dateTime", event["start"].get("date"))
        end = event["end"].get("dateTime", event["end"].get("date"))

        summary = event.get("summary", "(kein Titel)",)
       
        description = event.get("description", "")
       
        # wenn einer der Tags in den Überschriften ist, dann führe das untere aus
        if any(tag in summary for tag in TAGS):
            
            start_dt = parser.parse(start) 
            end_dt = parser.parse(end) # Datum und Uhrzeit getrennt formatieren
            start_date = start_dt.strftime("%d.%m.%Y") 
            # wenn der startwert für die Excel 32 ist dann setze das Firstdate (der Tag an dem das erste mal ein Termin oder die Vorbereitung stattfindet)
            if i == 32:
                First_day = start_date
            # passe immer das Lastday an das start_date an, somit wird der letzte Tag durchgehen ermittelt     
            Last_day = start_date 
            start_time = start_dt.strftime("%H:%M") 
            end_date = end_dt.strftime("%d.%m.%Y") 
            end_time = end_dt.strftime("%H:%M")

            #berechnet die Stundendiferenz
            dif =  end_dt - start_dt 
            decimal_hours = dif.total_seconds() / 3600
            total_minutes = int(dif.total_seconds() // 60) 
            hours = total_minutes // 60 
            minutes = total_minutes % 60
            month = datetime.now().strftime("%B")
            #month = datetime(2026,3,1).strftime("%B")
            #print(f"{decimal_hours} Stunden")
          
          
            excel_setter(i,ws, datum=start_date, decimal_hours=decimal_hours, description=description,First_day=First_day, Last_day=Last_day) 
              
            wb.save(OUTPUT_DIR / f"Muster_Honorarrechnung-Lehrkräfte_{month}.xlsx")
            i += 1 
          
