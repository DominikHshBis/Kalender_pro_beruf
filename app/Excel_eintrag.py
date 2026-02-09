from openpyxl import load_workbook

# Excel-Datei laden
wb = load_workbook("app/Muster_Honorarrechnung-Lehrkräfte_pytest.xlsx")
ws = wb.active

for i in range(32,33):
    
# Zeile 31 befüllen
    ws[f"A{i}"] = "02/04/2026"          # Datum
    ws[f"B{i}"] = 0.75                  # Stunden
    ws[f"C{i}"] = "Vorbereitung wie gewünscht"        # Leistung
    ws[f"D{i}"] = 19    
                   # Stundensatz
            
#asd
# Datei speichernd
wb.save("Test_Eintrag.xlsx")
