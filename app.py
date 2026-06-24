import streamlit as st
import openpyxl
from io import BytesIO
import os
import re

st.set_page_config(page_title="Janmar WZ Stokrotka", page_icon="🥦", layout="centered")

st.title("🥦 JANMAR WZ-Stokrotka Web v3.1")
st.subheader("Oficjalny, pancerny silnik chmurowy Janmar")
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

def parsuj_stokrotka_txt_pancerny(file_bytes):
    try: tekst = file_bytes.decode('windows-1250')
    except: tekst = file_bytes.decode('utf-8', errors='ignore')
        
    tekst = tekst.replace('\r', '')
    lines = tekst.split('\n')
    
    nr_zamowienia = "........"
    data_zamowienia = "................"
    data_dostawy = "................"
    miejsce_dostawy = "STOKROTKA CD"
    
    for line in lines:
        if "ZAMÓWIENIE TOWARU" in line:
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
        parts = line_raw.split()
        if len(parts) >= 5:
            surowy_kod = parts[0]
            kod_bazowy = surowy_kod.split('/')[0].strip()
            
            # Całkowita odporność na spacje – sprawdzamy czy pierwszy wyraz to 6 cyfr kodu
            if kod_bazowy.isdigit() and len(kod_bazowy) == 6:
                try:
                    # Ilość końcowa to zawsze ostatnia wartość w linii Stokrotki
                    ilosc_koncowa = float(parts[-1].replace(',', '.'))
                    
                    jm = "szt"
                    for p in parts:
                        if p.lower() in ['szt', 'szt.', 'kg', 'kg.']:
                            jm = p.lower().replace('.', '')
                            break
                    
                    # Wyciągamy nazwę ze środka wiersza
                    nazwa_parts = []
                    for p in parts[1:]:
                        if p.lower() in ['szt', 'szt.', 'kg', 'kg.'] or p.replace(',', '').replace('.', '').isdigit():
                            break
                        nazwa_parts.append(p)
                    nazwa_towaru = " ".join(nazwa_parts).upper()
                    
                    if ilosc_koncowa > 0 and len(nazwa_towaru) > 2:
                        w_opak = 10.0
                        for klucz, waga in PRZELICZNIKI_STOKROTKA.items():
                            if klucz in nazwa_towaru:
                                w_opak = waga
                                break
                                
                        ilosc_op = round(ilosc_koncowa / w_opak, 1) if w_opak > 0 else 0
                        kraj = ustal_kraj_pochodzenia(nazwa_towaru)
                        
                        pozycje.append({
                            "kod": kod_bazowy, "nazwa": nazwa_towaru, "jm": jm,
                            "w_opak": w_opak, "ilosc_op": ilosc_op, "ilosc_koncowa": ilosc_koncowa,
                            "kraj": kraj
                        })
                except:
                    pass
                
    return nr_zamowienia, data_zamowienia, data_dostawy, miejsce_dostawy, pozycje

szablon_path = "szablon_wz.xlsx"

if not os.path.exists(szablon_path):
    st.error("Brak pliku 'szablon_wz.xlsx' w folderze GitHub! Wgraj go najpierw.")
else:
    uploaded_file = st.file_uploader("Wgraj plik zamówienia Stokrotki (.TXT)", type=["txt"])

    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        with st.spinner("Generowanie dokumentu na oryginalnym szablonie..."):
            nr_zam, data_zam, data_dost, miejsce, towary = parsuj_stokrotka_txt_pancerny(file_bytes)
            
            if towary:
                wb = openpyxl.load_workbook(szablon_path)
                ws = wb["Dokument WZ"]
                
                ws['C7'] = nr_zam
                ws['C8'] = data_zam
                ws['C9'] = data_dost
                ws['A11'] = f" MIEJSCE DOSTAWY: {miejsce.upper()}"
                
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
                
                st.success(f"Wygenerowano WZ na oryginalnym szablonie dla zamówienia nr: {nr_zam}")
                st.download_button(
                    label="📥 Pobierz oficjalną WZ Stokrotka (Excel)",
                    data=wz_excel,
                    file_name=f"WZ_STOKROTKA_{nr_zam.replace('/', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Błąd odczytu. System nie odnalazł linii produktowych.")
