import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from io import BytesIO
import re

st.set_page_config(page_title="Janmar WZ Stokrotka Diagnostyka", layout="centered")

st.title("JANMAR WZ-Stokrotka Web v7.1 (PODGLĄD SPECJALNY)")
st.write("Wgraj surowy plik tekstowy (.TXT), aby zobaczyć co blokuje odczyt.")

uploaded_file = st.file_uploader("Wgraj plik zamówienia Stokrotki (.TXT)", type=["txt"])

if uploaded_file is not None:
    nazwa_pliku = uploaded_file.name
    file_bytes = uploaded_file.read()
    
    tekst = None
    for enc in ['cp1250', 'windows-1250', 'utf-8', 'latin1']:
        try:
            tekst = file_bytes.decode(enc)
            if "ZAMÓWIENIE" in tekst or "Stokrotka" in tekst or len(tekst) > 50:
                break
        except:
            continue
            
    if tekst is None:
        tekst = file_bytes.decode('utf-8', errors='ignore')
    
    linie = tekst.splitlines()
    
    # WYŚWIETLAMY PODGLĄD NA STRONIE DLA MARCINA
    st.warning("📊 Poniżej widzisz pierwsze 15 linii Twojego pliku, które widzi serwer:")
    for i, l in enumerate(linie[:15]):
        st.code(f"Linia {i+1}: {l}")

    towary = []
    for line_raw in linie:
        match_kod = re.search(r'\b\d{6}\b', line_raw)
        if match_kod:
            kod_part = match_kod.group(0)
            parts = line_raw.split()
            if len(parts) >= 3:
                try:
                    ilosc_koncowa = float(parts[-1].replace(',', '.'))
                    if ilosc_koncowa > 0:
                        towary.append({"kod": kod_part, "nazwa": line_raw[:30].strip(), "jm": "szt", "w_opak": 10, "ilosc_op": 1, "ilosc_koncowa": ilosc_koncowa, "kraj": "POLSKA"})
                except: pass

    if towary:
        st.success(f"Opa! Jednak znaleziono {len(towary)} towarów! Przycisk pobierania odblokowany.")
        # Prosty testowy plik excel
        wb = openpyxl.Workbook()
        output = BytesIO()
        wb.save(output)
        st.download_button(label="📥 Pobierz WZ TEST", data=output.getvalue(), file_name="TEST.xlsx")
    else:
        st.error("Błąd odczytu. System przeskanował te linie powyżej i żadna nie pasowała do kodu 6 cyfr.")
