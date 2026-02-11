from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timedelta
from dateutil import parser
import calendar
from Excel_eintrag import excel_setter
import json

from openpyxl import load_workbook

# lädt json config
with open("config/config_dominik.json","r") as config:
    config = json.load(config)

#teilt congif inhalt den variablen zu
CALENDAR_ID = config["calendar_id"]
TAGS = config["tags"] #list of tags to search for in calendar events
SERVICE_ACCOUNT_FILE = "config/credentials.json"
SCOPES = [config["scopes"]] # api adress to access calendar data
EXCEL_LOAD_PATH = config["excel_load_path"]

wb = load_workbook(EXCEL_LOAD_PATH)
First_day = ""
Last_day = ""
ws = wb.active

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)

service = build("calendar", "v3", credentials=credentials)



# Aktuellen Monat berechnen
now = datetime.utcnow()

# now = datetime(2026,4,1) gibt jahr monat und Tag an
first_day = datetime(now.year, now.month, 1)
last_day = datetime(now.year, now.month, calendar.monthrange(now.year, now.month)[1])

# In RFC3339-Format umwandeln
time_min = first_day.isoformat() + "Z"
time_max = (last_day + timedelta(days=1)).isoformat() + "Z"

events_result = service.events().list(
    calendarId=CALENDAR_ID,
    timeMin=time_min,
    timeMax=time_max,
    singleEvents=True,
    orderBy="startTime",
).execute()

events = events_result.get("items", [])

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
            Last_day = start_date # 
            start_time = start_dt.strftime("%H:%M") 
            end_date = end_dt.strftime("%d.%m.%Y") 
            end_time = end_dt.strftime("%H:%M")

            #berechnet den letzten die Stundendiferenz
            dif =  end_dt - start_dt 
            decimal_hours = dif.total_seconds() / 3600
            total_minutes = int(dif.total_seconds() // 60) 
            hours = total_minutes // 60 
            minutes = total_minutes % 60
            month = datetime.now().strftime("%B")
            #month = datetime(2026,3,1).strftime("%B")
            #print(f"{decimal_hours} Stunden")
          #  print(i)
            #print(summary, start_date, end_date,start_time, end_time, dif, f"{decimal_hours} stunden", description)
            excel_setter(i,ws, datum=start_date, decimal_hours=decimal_hours, description=description,First_day=First_day, Last_day=Last_day) 
              
            wb.save(f"Muster_Honorarrechnung-Lehrkräfte_{month}.xlsx")
            i += 1 
          
