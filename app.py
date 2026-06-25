import streamlit as st
import openpyxl
from io import BytesIO
import os

st.set_page_config(page_title="Janmar WZ Stokrotka Excel", layout="centered")

st.title("JANMAR WZ-Stokrotka Web v6.0")
st.subheader("Nowy silnik oparty na Zbiorczym Raporcie Zakupowym (Excel)")
st.write("Wgraj plik 'ZBIORCZY_RAPORT_ZAKUPOWY.xlsx', aby wygenerować oficjalną WZ Stokrotki.")

PRZELICZNIKI_SIECI = {
    "KALAFIOR": 6.0, "KAPUSTA PEKIŃSKA": 10.0, "KAPUSTA BIAŁA": 10.0,
    "KAPUSTA CZERWONA": 10.0, "KAPUSTA WŁOSKA": 10.0, "KAPUSTA WCZESNA": 6.0,
    "CUKINIA": 5.0, "KALAREPA": 8.0, "KUKURYDZA": 30.0, "RZODKIEWKA": 10.0,
    "SAŁATA LODOWA": 10.0, "SAŁATA MASŁOWA": 12.0, "SELER NACIOWY": 16.0, "MARCHEW": 10.0
}

def ustal_kraj_pochodzenia(nazwa_towaru):
    nazwa_upper = nazwa_towaru.upper()
    if any(x in nazwa_upper for x in ["POMARAŃCZE", "MANDARYNKI"]): return "HISZPANIA / TURCJA"
    if "ZIEMNIAKI IMPORTOWE" in nazwa_upper: return "CYPR / GRECJA"
    if "ARBUZ" in nazwa_upper: return "WŁOCHY / HISZPANIA"
    return "POLSKA"

szablon_path = "szablon_wz.xlsx"

if not os.path.exists(szablon_path):
    st.error("Brak pliku 'szablon_wz.xlsx' w folderze GitHub! Upewnij się, że został wgrany.")
else:
    uploaded_file = st.file_uploader("Wgraj ZBIORCZY RAPORT ZAKUPOWY (.xlsx)", type=["xlsx"])

    if uploaded_file is not None:
        with st.spinner("Przetwarzanie raportu excel..."):
            try:
                # Otwieramy wgrany raport zakupowy
                wb_raport = openpyxl.load_workbook(uploaded_file, data_only=True)
                ws_raport = wb_raport.active
                
                towary = []
                
                # Czytamy wiersze raportu (zakładamy standardowy układ kolumn z bota lokalnego)
                for row in range(2, ws_raport.max_row + 1):
                    siec = str(ws_raport.cell(row=row, column=1).value).strip().upper() if ws_raport.cell(row=row, column=1).value else ""
                    
                    # Interesuje nas tylko i wyłącznie Stokrotka
                    if "STOKROTKA" in siec:
                        nazwa = str(ws_raport.cell(row=row, column=2).value).strip().upper() if ws_raport.cell(row=row, column=2).value else ""
                        ilosc_val = ws_raport.cell(row=row, column=3).value
                        
                        try:
                            ilosc_koncowa = float(ilosc_val) if ilosc_val else 0.0
                        except:
                            ilosc_koncowa = 0.0
                            
                        if ilosc_koncowa > 0 and len(nazwa) > 2:
                            # Próba dopasowania kodu (jeśli jest w kolumnie lub domyślny)
                            kod = "000000"
                            # Szukamy czy w nazwie lub innej kolumnie nie ma kodu, jeśli nie - zostaje domyślny
                            
                            w_opak = 10.0
                            for klucz, waga in PRZELICZNIKI_SIECI.items():
                                if klucz in nazwa:
                                    w_opak = waga
                                    break
                                    
                            ilosc_op = round(ilosc_koncowa / w_opak, 1) if w_opak > 0 else 0
                            kraj = ustal_kraj_pochodzenia(nazwa)
                            
                            towary.append({
                                "kod": kod, "nazwa": nazwa, "jm": "szt" if "SZT" in nazwa else "kg",
                                "w_opak": w_opak, "ilosc_op": ilosc_op, "ilosc_koncowa": ilosc_koncowa,
                                "kraj": kraj
                            })
                
                if towary:
                    # Ładujemy oryginalny szablon WZ
                    wb_wz = openpyxl.load_workbook(szablon_path)
                    ws_wz = wb_wz["Dokument WZ"]
                    
                    # Wypełniamy nagłówki standardowymi danymi
                    ws_wz['C7'] = "ZBIORCZE"
                    ws_wz['C8'] = "................"
                    ws_wz['C9'] = "................"
                    ws_wz['A11'] = " MIEJSCE DOSTAWY: STOKROTKA MAGAZYN CENTRALNY"
                    
                    # Wklejamy odfiltrowane towary Stokrotki od 15 wiersza
                    start_row = 15
                    for idx, t in enumerate(towary):
                        r = start_row + idx
                        if r <= 60:
                            ws_wz.cell(row=r, column=2, value=t['kod'])
                            ws_wz.cell(row=r, column=3, value=t['nazwa'])
                            ws_wz.cell(row=r, column=4, value=t['kraj'])
                            ws_wz.cell(row=r, column=5, value=t['jm'])
                            ws_wz.cell(row=r, column=6, value=t['w_opak'])
                            ws_wz.cell(row=r, column=7, value=t['ilosc_op'])
                            ws_wz.cell(row=r, column=8, value=t['ilosc_koncowa'])
                    
                    output = BytesIO()
                    wb_wz.save(output)
                    wz_excel = output.getvalue()
                    
                    st.success(f"Sukces! Wyciągnięto {len(towary)} pozycji dla Stokrotki i naniesiono na szablon.")
                    st.download_button(
                        label="📥 Pobierz WZ Stokrotka z Raportu (Excel)",
                        data=wz_excel,
                        file_name="WZ_STOKROTKA_Z_RAPORTU.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("W raporcie nie znaleziono żadnych pozycji oznaczonych jako STOKROTKA.")
            except Exception as e:
                st.error(f"Błąd przetwarzania pliku Excel: {str(e)}")
