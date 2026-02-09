from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timedelta
import calendar

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
    for event in events:
        start = event["start"].get("dateTime", event["start"].get("date"))
        summary = event.get("summary", "(kein Titel)")
        description=event.get("description", "")
        print(start, summary, description)
