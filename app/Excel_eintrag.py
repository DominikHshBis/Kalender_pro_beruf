from openpyxl import load_workbook


# Excel-Datei laden

def excel_setter(i,ws, datum, decimal_hours, description):
    
    
        
    # Zeile 31 befüllen
    ws[f"A{i}"] = datum     # Datum
    ws[f"B{i}"] = decimal_hours                # Stunden
    ws[f"C{i}"] = description        # Leistung
    ws[f"D{i}"] = 19
    ws[f"E{i}"] = f"=ROUND(B{i}*$D$32,2)"  #f"=RUNDEN(B{i}*$D$32;2)"          # Berechnung der Kosten
    #wb.save("Test_Eintrag.xlsx")           # Stundensatzf
            #asdnn
#asd
# Datei speichernd

