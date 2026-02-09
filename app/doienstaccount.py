from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timedelta
from dateutil import parser
import calendar
from Excel_eintrag import excel_setter

from openpyxl import load_workbook

First_day = ""
Last_day = ""
wb = load_workbook("app/Muster_Honorarrechnung-Lehrkräfte_pytest.xlsx")
ws = wb.active

TAGS = ["#Pro", "#Vor","#pro", "#vor"]
SERVICE_ACCOUNT_FILE = "credentials.json"
SCOPES = ["https://www.googleapis.com/auth/calendar"]

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)

service = build("calendar", "v3", credentials=credentials)

CALENDAR_ID = "dominikballentin@gmail.com"

# Aktuellen Monat berechnen
now = datetime.utcnow()
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
       

        if any(tag in summary for tag in TAGS):
            
            start_dt = parser.parse(start) 
            end_dt = parser.parse(end) # Datum und Uhrzeit getrennt formatieren
            start_date = start_dt.strftime("%d.%m.%Y") 
            if i == 32:
                First_day = start_date
            Last_day = start_date
            start_time = start_dt.strftime("%H:%M") 
            end_date = end_dt.strftime("%d.%m.%Y") 
            end_time = end_dt.strftime("%H:%M")

            dif =  end_dt - start_dt
            decimal_hours = dif.total_seconds() / 3600
            total_minutes = int(dif.total_seconds() // 60) 
            hours = total_minutes // 60 
            minutes = total_minutes % 60
            month = datetime.now().strftime("%B")
            #print(f"{decimal_hours} Stunden")
          #  print(i)
            #print(summary, start_date, end_date,start_time, end_time, dif, f"{decimal_hours} stunden", description)
            excel_setter(i,ws, datum=start_date, decimal_hours=decimal_hours, description=description,First_day=First_day, Last_day=Last_day) 
              
            wb.save(f"Muster_Honorarrechnung-Lehrkräfte_{month}.xlsx")
            i += 1 
          