import streamlit as st
import openpyxl
from io import BytesIO
import os

st.set_page_config(page_title="Janmar WZ Stokrotka", page_icon="🥦", layout="centered")

st.title("🥦 JANMAR WZ-Stokrotka Web v3.0")
st.subheader("Oficjalny, oryginalny silnik z komputera Mac")
st.write("Wgraj surowy plik tekstowy (.TXT), aby nanieść dane na oryginalny szablon WZ.")

PRZELICZNIKI_STOKROTKA = {
    "BROKUŁ": 10.0, "KALAFIOR": 6.0, "KAPUSTA PEKIŃSKA": 10.0, "KAPUSTA BIAŁA": 10.0,
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

def parsuj_stokrotka_txt(file_bytes):
    try: tekst = file_bytes.decode('windows-1250')
    except: tekst = file_bytes.decode('utf-8', errors='ignore')
        
    lines = tekst.split('\n')
    
    nr_zamowienia = "........"
    data_zamowienia = "................"
    data_dostawy = "................"
    miejsce_dostawy = "STOKROTKA CD"
    
    for line in lines:
        if "ZAMÓWIENIE TOWARU" in line:
            import re
            match = re.search(r'NR\s*[:]?\s*([\d/]+)', line, re.IGNORECASE)
            if match: nr_zamowienia = match.group(1)
        if "Termin dostawy" in line:
            match_dost = re.search(r'Termin dostawy\s+([\d\.]+)', line)
            match_wyst = re.search(r'Data wystawienia\s+([\d\.]+)', line)
            if match_dost: data_dostawy = match_dost.group(1)
            if match_wyst: data_zamowienia = match_wyst.group(1)
        if "Towar należy dostarczyć:" in line:
            try:
                idx = lines.index(line)
                if idx + 1 < len(lines): miejsce_dostawy = lines[idx+1].strip()
            except: pass

    pozycje = []
    for line_raw in lines:
        if len(line_raw) < 50:
            continue
            
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
                    for klucz, waga in PRZELICZNIKI_STOKROTKA.items():
                        if klucz in nazwa_towaru:
                            w_opak = waga
                            break
                            
                    ilosc_op = round(ilosc_koncowa / w_opak, 1) if w_opak > 0 else 0
                    kraj = ustal_kraj_pochodzenia(nazwa_towaru)
                    
                    pozycje.append({
                        "kod": kod_towaru, "nazwa": nazwa_towaru, "jm": jm,
                        "w_opak": w_opak, "ilosc_op": ilosc_op, "ilosc_koncowa": ilosc_koncowa,
                        "kraj": kraj
                    })
            except:
                pass
                
    return nr_zamowienia, data_zamowienia, data_dostawy, miejsce_dostawy, pozycje

# SPRAWDZAMY CZY PLIK SZABLONU ISTNIEJE NA SERWERZE
szablon_path = "szablon_wz.xlsx"

if not os.path.exists(szablon_path):
    st.error("Brak pliku 'szablon_wz.xlsx' w folderze GitHub! Wgraj go najpierw.")
else:
    uploaded_file = st.file_uploader("Wgraj plik zamówienia Stokrotki (.TXT)", type=["txt"])

    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        with st.spinner("Naniesienie danych na oryginalny szablon..."):
            nr_zam, data_zam, data_dost, miejsce, towary = parsuj_stokrotka_txt(file_bytes)
            
            if towary:
                # Otwieramy oryginalny plik szablonu z Maca
                wb = openpyxl.load_workbook(szablon_path)
                ws = wb["Dokument WZ"]
                
                # Wpisywanie danych nagłówka dokładnie w kratki z szablonu
                ws['C7'] = nr_zam
                ws['C8'] = data_zam
                ws['C9'] = data_dost
                ws['A11'] = f" MIEJSCE DOSTAWY: {miejsce.upper()}"
                
                # Wpisywanie towarów w puste linie (od wiersza 15)
                start_row = 15
                for idx, t in enumerate(towary):
                    r = start_row + idx
                    if r <= 60: # Zabezpieczenie siatki
                        ws.cell(row=r, column=2, value=t['kod'])
                        ws.cell(row=r, column=3, value=t['nazwa'])
                        ws.cell(row=r, column=4, value=t['kraj'])
                        ws.cell(row=r, column=5, value=t['jm'])
                        ws.cell(row=r, column=6, value=t['w_opak'])
                        ws.cell(row=r, column=7, value=t['ilosc_op'])
                        ws.cell(row=r, column=8, value=t['ilosc_koncowa'])
                
                # Zapis do pamięci podręcznej i udostępnienie do pobrania
                output = BytesIO()
                wb.save(output)
                wz_excel = output.getvalue()
                
                st.success(f"Wygenerowano WZ na oryginalnym szablonie dla zamówienia nr: {nr_zam}")
                st.download_button(
                    label="📥 Pobierz idealną WZ Stokrotka (Excel)",
                    data=wz_excel,
                    file_name=f"WZ_STOKROTKA_{nr_zam.replace('/', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Błąd odczytu. System nie odnalazł linii produktowych.")
