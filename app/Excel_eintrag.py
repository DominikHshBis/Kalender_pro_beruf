from openpyxl import load_workbook


# Excel-Datei laden

def excel_setter(i,ws, datum, decimal_hours, description,First_day, Last_day, stundensatz=19):
    
    # Zeile 31 befüllen
   
    ws[f"A{i}"] = datum     # Datum
    ws[f"B{i}"] = decimal_hours                # Stunden
    ws[f"C{i}"] = description        # Leistung
    ws[f"D{i}"] = stundensatz
    ws[f"E{i}"] = f"=ROUND(B{i}*$D$32,2)"  
    ws["C24"] = f"{First_day} bis {Last_day}"
    ws["E23"] = f"{Last_day}"
    ws["A29"] = f"für die unten aufgeführten Leistungen als Honorardozent:in im Projekt Assistierte Ausbildung AsA VIII berechne ich Ihnen wie folgt: "
#asd
# Datei speichernd

