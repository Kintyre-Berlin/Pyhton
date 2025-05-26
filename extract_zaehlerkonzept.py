import camelot
import pandas as pd
from pathlib import Path

# 1. Pfade definieren
pdf_path = Path(r"C:\Users\Tino\Downloads\zwischen\test_versuch\MA-001.00-TA-Zählerkonzept Heizung, Sanitär und Kälte.pdf")
output_path = pdf_path.with_name("Zählerkonzept_Export.xlsx")

# 2. Tabellen aus den jeweiligen Seitenbereichen extrahieren
tables_heizung  = camelot.read_pdf(str(pdf_path), pages="3-5", flavor="lattice")
tables_sanitaer = camelot.read_pdf(str(pdf_path), pages="6-8", flavor="lattice")
tables_kaelte   = camelot.read_pdf(str(pdf_path), pages="9-10", flavor="lattice")

# 3. Teil-Tabellen je Bereich zu einem DataFrame zusammenführen
df_heizung  = pd.concat([t.df for t in tables_heizung],  ignore_index=True)
df_sanitaer = pd.concat([t.df for t in tables_sanitaer], ignore_index=True)
df_kaelte   = pd.concat([t.df for t in tables_kaelte],   ignore_index=True)

# 4. Erste Zeile als Header setzen (falls nötig)
for df in (df_heizung, df_sanitaer, df_kaelte):
    df.columns = df.iloc[0]   # erste Zeile zu Spaltennamen
    df.drop(0, inplace=True)  # erste Zeile entfernen
    df.reset_index(drop=True, inplace=True)

# 5. In eine Excel-Datei schreiben mit drei Sheets
with pd.ExcelWriter(str(output_path), engine="openpyxl") as writer:
    df_heizung.to_excel(writer, sheet_name="Heizung", index=False)
    df_sanitaer.to_excel(writer, sheet_name="Sanitär", index=False)
    df_kaelte.to_excel(writer, sheet_name="Kälte", index=False)

print(f"✅ Export fertig: {output_path}")
