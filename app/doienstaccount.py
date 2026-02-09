from google.oauth2 import service_account
from googleapiclient.discovery import build

# Pfad zu deiner JSON-Datei
SERVICE_ACCOUNT_FILE = "credentials.json"

# Die benötigten Berechtigungen (Scopes)
SCOPES = ["https://www.googleapis.com/auth/calendar"]

# Dienstkonto authentifizieren
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)

# Calendar API Client erstellen
service = build("calendar", "v3", credentials=credentials)

# WICHTIG: Kalender-ID deines Kalenders eintragen
# (zu finden in Google Kalender → Einstellungen → Kalender integrieren)
CALENDAR_ID = "dominikballentin@gmail.com"

# Beispiel: Nächste 10 Termine abrufen
events_result = service.events().list(
    calendarId=CALENDAR_ID,
    maxResults=10,
    singleEvents=True,
    orderBy="startTime"
).execute()

events = events_result.get("items", [])

if not events:
    print("Keine Termine gefunden.")
else:
    for event in events:
        start = event["start"].get("dateTime", event["start"].get("date"))
        print(start, event["summary"])
