import requests
import json
from pathlib import Path
from datetime import datetime, time, timedelta

class FastBill_invoice_creator:
    def __init__(self, project_root: Path):
        with open(project_root) as f:
            self.config_data = json.load(f)

        self.url = "https://my.fastbill.com/api/1.0/api.php"
        self.headers = {
            "key": self.config_data["Authorization"],
            "content_type": self.config_data["Content-Type"]
        }
        self.current_date = datetime.now().strftime("%d.%m.%Y")
        self.standard_text= f"\n die Abrechnung erfolgt ausschließlich auf Grundlage dieser Rechnung. Der vom Träger erstellte Leistungsnachweis vom {self.current_date} dient lediglich als Tätigkeitsnachweis und stellt keine Rechnung dar. Eine doppelte Abrechnung ist ausgeschlossen. Für die unten aufgeführten Leistungen für den Stütz- und Förderunterricht in der Ausbildungsassistenz im Projekt Assistierte Ausbildung AsA VIII / AsA IX berechne ich Ihnen wie folgt:"

    def create_invoice(self, current_date = "2026-03-30", start_date = "2026-03-01", end_date = "2026-03-31", quantity_preparing = 1, quantity_teaching = 1) -> int: 
        """Erstellt eine Rechnung bei FastBill und gibt die Rechnungs-ID zurück., leider noch nicht fertig da stunden noch anhand der tage errechnet werden und nicht nach kalender API"""
        #print(start_date, end_date, quantity_preparing, quantity_teaching)
        
        payload = json.dumps({
            "SERVICE": "invoice.create",
            "DATA": {
                "CUSTOMER_ID": 14375916,
                "CURRENCY_CODE": "EUR",
                "INVOICE_TITLE": "",
                "INTROTEXT": self.standard_text,
                "INVOICE_DATE": current_date,
                "SERVICE_PERIOD_START": start_date,
                "SERVICE_PERIOD_END": end_date,
                "VAT_CASE": 0,
                "IS_GROSS": 1,
                "ITEMS": [
                    {
                        "DESCRIPTION": "Vorbereitung Stütz-Unterricht\nStütz-Unterrichtsvorbereitung für Auszubildende (0,75 Stunden einer vollen Unterrichtsstunde)",
                        "QUANTITY": quantity_preparing,
                        "UNIT": "Stunde",
                        "UNIT_PRICE": 14.25,
                        "VAT_PERCENT": 0
                    },
                    {
                        "DESCRIPTION": "Stütz-Unterricht\nFörderunterricht und Unterstützung in der Ausbildung",
                        "QUANTITY": quantity_teaching, 
                        "UNIT": "Stunde",
                        "UNIT_PRICE": 19,
                        "VAT_PERCENT": 0
                    }
                ]
            }
        })
        response = requests.post(self.url,auth=(self.config_data["user"], self.config_data["Authorization"]), headers=self.headers, data=payload)
        response_data = response.json()
        #print(response_data)  # Zum Debuggen die vollständige Antwort anzeigen
        return response_data["RESPONSE"]["INVOICE_ID"]
    
    def delete_invoice(self, invoice_id: int):
        payload = json.dumps({
            "SERVICE": "invoice.delete", 
            "DATA": {
                "INVOICE_ID": invoice_id
            }
        })
        response = requests.post(self.url,auth=(self.config_data["user"], self.config_data["Authorization"]), headers=self.headers, data=payload)
        return response.json()
#delete_response = requests.request("POST", url, headers=headers, data=delete_post_request(invoice_id))


