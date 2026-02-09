from openpyxl import load_workbook

# Excel-Datei laden
wb = load_workbook("app/Muster_Honorarrechnung-Lehrkräfte_pytest.xlsx",data_only=False)
ws = wb.active

for i in range(32,35):
    
# Zeile 31 befüllen
    ws[f"A{i}"] = "02/04/2026"          # Datum
    ws[f"B{i}"] = 0.75                  # Stunden
    ws[f"C{i}"] = "Vorbereitung wie gewünscht"        # Leistung
    ws[f"D{i}"] = 19
    ws[f"E{i}"] = f"=ROUND(B{i}*$D${i},2)"         # Berechnung der Kosten
                   # Stundensatz
            
#asd
# Datei speichernd
wb.save("Test_Eintrag.xlsx")
