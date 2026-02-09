from abc import ABC, abstractmethod
from typing import List, Dict, Any
import os
from pathlib import Path

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


class AuthenticationManager:
    """Single Responsibility: Verwaltet die Authentifizierung mit Google Calendar API"""
    
    def __init__(self, credentials_file: str, token_file: str, scopes: List[str]):
        self.credentials_file = credentials_file
        self.token_file = token_file
        self.scopes = scopes
    
    def get_credentials(self) -> Credentials:
        """Holt oder erstellt Authentifizierungstoken"""
        if self._token_exists():
            return self._load_credentials_from_file()
        return self._create_new_credentials()
    
    def _token_exists(self) -> bool:
        """Überprüft, ob ein Token bereits existiert"""
        return Path(self.token_file).exists()
    
    def _load_credentials_from_file(self) -> Credentials:
        """Lädt Credentials aus einer bestehenden Token-Datei"""
        try:
            return Credentials.from_authorized_user_file(self.token_file, self.scopes)
        except Exception as e:
            print(f"Fehler beim Laden der Credentials: {e}")
            return self._create_new_credentials()
    
    def _create_new_credentials(self) -> Credentials:
        """Erstellt neue Credentials durch OAuth2-Flow"""
        flow = InstalledAppFlow.from_client_secrets_file(
            self.credentials_file, self.scopes
        )
        creds = flow.run_local_server(port=0)
        self._save_credentials(creds)
        return creds
    
    def _save_credentials(self, credentials: Credentials) -> None:
        """Speichert die Credentials in eine Datei"""
        with open(self.token_file, "w") as token_file:
            token_file.write(credentials.to_json())


class CalendarServiceInterface(ABC):
    """Interface für Kalender-Services (Interface Segregation Principle)"""
    
    @abstractmethod
    def get_events(self, calendar_id: str) -> List[Dict[str, Any]]:
        """Holt alle Events aus dem Kalender"""
        pass


class GoogleCalendarService(CalendarServiceInterface):
    """Single Responsibility: Interaktion mit Google Calendar API"""
    
    def __init__(self, authentication_manager: AuthenticationManager):
        self.auth_manager = authentication_manager
        self.service = self._build_service()
    
    def _build_service(self):
        """Erstellt den Google Calendar API Service"""
        credentials = self.auth_manager.get_credentials()
        return build("calendar", "v3", credentials=credentials)
    
    def get_events(self, calendar_id: str = "primary") -> List[Dict[str, Any]]:
        """Holt alle Events aus dem angegebenen Kalender"""
        try:
            result = self.service.events().list(calendarId=calendar_id).execute()
            return result.get("items", [])
        except Exception as e:
            print(f"Fehler beim Abrufen der Events: {e}")
            return []


class EventProcessor:
    """Single Responsibility: Verarbeitet und formatiert Events"""
    
    @staticmethod
    def process_events(events: List[Dict[str, Any]]) -> List[Dict[str, str]]:
        """Verarbeitet Events und extrahiert relevante Informationen"""
        processed_events = []
        for event in events:
            processed_event = EventProcessor._extract_event_info(event)
            if processed_event:
                processed_events.append(processed_event)
        return processed_events
    
    @staticmethod
    def _extract_event_info(event: Dict[str, Any]) -> Dict[str, str]:
        """Extrahiert relevante Informationen aus einem Event"""
        return {
            "summary": event.get("summary", "Kein Titel"),
            "start": event.get("start", {}).get("dateTime", event.get("start", {}).get("date", "N/A"))
        }


class EventPrinter:
    """Single Responsibility: Gibt Events aus (Open/Closed Principle - leicht erweiterbar)"""
    
    def print_events(self, events: List[Dict[str, str]]) -> None:
        """Gibt processierte Events aus"""
        if not events:
            print("Keine Events gefunden.")
            return
        
        for event in events:
            self._print_event(event)
    
    @staticmethod
    def _print_event(event: Dict[str, str]) -> None:
        """Gibt ein einzelnes Event aus"""
        print(f"{event['summary']} - {event['start']}")


class CalendarApplicationOrchestrator:
    """Orchestriert die Zusammenarbeit der verschiedenen Komponenten"""
    
    def __init__(
        self,
        credentials_file: str,
        token_file: str,
        scopes: List[str]
    ):
        self.auth_manager = AuthenticationManager(credentials_file, token_file, scopes)
        self.calendar_service = GoogleCalendarService(self.auth_manager)
        self.event_processor = EventProcessor()
        self.event_printer = EventPrinter()
    
    def run(self, calendar_id: str = "primary") -> None:
        """Führt die komplette Anwendungslogik aus"""
        events = self.calendar_service.get_events(calendar_id)
        processed_events = self.event_processor.process_events(events)
        self.event_printer.print_events(processed_events)


# Main-Anwendung
if __name__ == "__main__":
    SCOPES = ["https://www.googleapis.com/auth/calendar.readonly"]
    
    app = CalendarApplicationOrchestrator(
        credentials_file="credentials.json",
        token_file="token.json",
        scopes=SCOPES
    )
    app.run()

