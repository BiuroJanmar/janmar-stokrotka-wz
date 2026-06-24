import streamlit as st
import openpyxl
from io import BytesIO
import os
import re

st.set_page_config(page_title="Janmar WZ Stokrotka", layout="centered")

st.title("JANMAR WZ-Stokrotka Web v4.4")
st.subheader("Oryginalny i czysty silnik z Twojego komputera Mac")
st.write("Wgraj surowy plik tekstowy (.TXT) zamówienia ze Stokrotki.")

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
    uploaded_file = st.file_uploader("Wgraj plik zamówienia Stokrotki (.TXT)", type=["txt"])

    if uploaded_file is not None:
        nazwa_pliku = uploaded_file.name
        file_bytes = uploaded_file.read()
        
        try: tekst = file_bytes.decode('cp1250', errors='ignore')
        except: tekst = file_bytes.decode('utf-8', errors='ignore')
        
        linie = tekst.splitlines(keepends=True)
        
        nr_zam = "Nieznany"
        nazwa_pliku_lower = nazwa_pliku.lower()
        if "nr" in nazwa_pliku_lower:
            try: nr_zam = nazwa_pliku_lower.split("nr")[1].split()[0].replace(".txt","")
            except: nr_zam = nazwa_pliku.replace(".txt","").replace(".TXT","")
        elif "zam." in nazwa_pliku_lower:
            try: nr_zam = nazwa_pliku_lower.split("zam.")[1].split()[0].replace(".txt","")
            except: nr_zam = nazwa_pliku.replace(".txt","").replace(".TXT","")
        else:
            nr_zam = nazwa_pliku.replace(".txt","").replace(".TXT","")
            
        nr_zam = "".join([c for c in nr_zam if c.isdigit() or c in ['/', '_', '-']])

        towary = []
        for line_raw in linie:
            if len(line_raw) > 45:
                kod_part = line_raw[0:15].strip()
                if '/' in kod_part:
                    kod_part = kod_part.split('/')[0].strip()
                    
                if kod_part.isdigit() and len(kod_part) == 6:
                    kod_towaru = kod_part
                    nazwa_towaru = line_raw[12:40].strip().upper()
                    jm = line_raw[40:46].strip().lower().replace('.', '')
                    ilosc_str = line_raw[46:57].strip().replace(',', '.')
                    
                    try:
                        ilosc_koncowa = float(ilosc_str)
                        if ilosc_koncowa > 0:
                            w_opak = 10.0
                            for klucz, waga in PRZELICZNIKI_SIECI.items():
                                if klucz in nazwa_towaru:
                                    w_opak = waga
                                    break
                                    
                            ilosc_op = round(ilosc_koncowa / w_opak, 1) if w_opak > 0 else 0
                            kraj = ustal_kraj_pochodzenia(nazwa_towaru)
                            
                            towary.append({
                                "kod": kod_towaru, "nazwa": nazwa_towaru, "jm": jm,
                                "w_opak": w_opak, "ilosc_op": ilosc_op, "ilosc_koncowa": ilosc_koncowa,
                                "kraj": kraj
                            })
                    except:
                        pass

        if towary:
            wb = openpyxl.load_workbook(szablon_path)
            ws = wb["Dokument WZ"]
            
            data_zamowienia = "................"
            data_dostawy = "................"
            miejsce_dostawy = "STOKROTKA CD"
            
            for line in linie:
                if "Termin dostawy" in line:
                    match_dost = re.search(r'Termin dostawy\s+([\d\.]+)', line)
                    match_wyst = re.search(r'Data wystawienia\s+([\d\.]+)', line)
                    if match_dost: data_dostawy = match_dost.group(1)
                    if match_wyst: data_zamowienia = match_wyst.group(1)
                if "Towar należy dostarczyć:" in line:
                    try:
                        idx = linie.index(line)
                        if idx + 1 < len(linie): miejsce_dostawy = linie[idx+1].strip()
                    except: pass

            ws['C7'] = nr_zam
            ws['C8'] = data_zamowienia
            ws['C9'] = data_dostawy
            ws['A11'] = f" MIEJSCE DOSTAWY: {miejsce_dostawy.upper()}"
            
            start_row = 15
            for idx, t in enumerate(towary):
                r = start_row + idx
                if r <= 60:
                    ws.cell(row=r, column=2, value=t['kod'])
                    ws.cell(row=r, column=3, value=t['nazwa'])
                    ws.cell(row=r, column=4, value=t['kraj'])
                    ws.cell(row=r, column=5, value=t['jm'])
                    ws.cell(row=r, column=6, value=t['w_opak'])
                    ws.cell(row=r, column=7, value=t['ilosc_op'])
                    ws.cell(row=r, column=8, value=t['ilosc_koncowa'])
            
            output = BytesIO()
            wb.save(output)
            wz_excel = output.getvalue()
            
            st.success("Sukces! Dane zostały naniesione na Twój szablon WZ.")
            st.download_button(
                label="📥 Pobierz oficjalną WZ Stokrotka (Excel)",
                data=wz_excel,
                file_name=f"WZ_STOKROTKA_{nr_zam}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Błąd odczytu. System nie odnalazł linii produktowych.")
