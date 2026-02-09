from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/calendar.readonly"]

def get_service():
    try:
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    except:
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
        creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return build("calendar", "v3", credentials=creds)

service = get_service()
events = service.events().list(calendarId="primary").execute()

for event in events.get("items", []):
    print(event["summary"], event["start"])
