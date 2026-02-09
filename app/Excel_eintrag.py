from openpyxl import load_workbook

# Excel-Datei laden
wb = load_workbook("app/Muster_Honorarrechnung-Lehrkräfte_pytest.xlsx")
ws = wb.active

# Zeile 31 befüllen
ws["A31"] = "02/04/2026"          # Datum
ws["B31"] = 0.75                  # Stunden
ws["C31"] = "Vorbereitung"        # Leistung
ws["D31"] = 19                    # Stundensatz
           
#asd
# Datei speichern
wb.save("Test_Eintrag.xlsx")
