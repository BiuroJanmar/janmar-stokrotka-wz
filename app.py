import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from io import BytesIO
import re

st.set_page_config(page_title="Janmar WZ Stokrotka", layout="centered")

st.title("JANMAR WZ-Stokrotka Web v8.3")
st.subheader("Oficjalny, w 100% dopasowany generator dokumentów WZ")
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

uploaded_file = st.file_uploader("Wgraj plik zamówienia Stokrotki (.TXT)", type=["txt"])

if uploaded_file is not None:
    nazwa_pliku = uploaded_file.name
    file_bytes = uploaded_file.read()
    
    tekst = None
    for enc in ['cp1250', 'windows-1250', 'utf-8', 'latin1']:
        try:
            tekst = file_bytes.decode(enc)
            if "ZAMÓWIENIE" in tekst or "Stokrotka" in tekst or "Termin" in tekst:
                break
        except:
            continue
            
    if tekst is None:
        tekst = file_bytes.decode('utf-8', errors='ignore')
    
    linie = tekst.splitlines()
    
    nr_zam = "Nieznany"
    for line in linie:
        if "ZAMÓWIENIE TOWARU NR" in line.upper():
            try:
                nr_zam = line.upper().split("NR")[1].strip().split()[0]
            except:
                pass
                
    if nr_zam == "Nieznany" or not nr_zam:
        nazwa_pliku_lower = nazwa_pliku.lower()
        if "nr" in nazwa_pliku_lower:
            try: nr_zam = nazwa_pliku_lower.split("nr")[1].split()[0].replace(".txt","")
            except: nr_zam = nazwa_pliku.replace(".txt","").replace(".TXT","")
        else:
            nr_zam = nazwa_pliku.replace(".txt","").replace(".TXT","")
            
    nr_zam = "".join([c for c in nr_zam if c.isdigit() or c in ['/', '_', '-']])

    towary = []
    for line_raw in linie:
        parts = line_raw.split()
        if len(parts) < 2:
            continue
            
        kod_part = parts[0].split('/')[0].strip()
        
        # Jeśli linia zaczyna się od 6 cyfr - bierzemy bez zbędnych pytań
        if kod_part.isdigit() and len(kod_part) == 6:
            try:
                # Ostatni element to ZAWSZE ilość
                ilosc_str = parts[-1].replace(',', '.')
                ilosc_koncowa = float(ilosc_str)
                
                jm = "szt"
                for p in parts:
                    if p.lower() in ['szt', 'szt.', 'kg', 'kg.']:
                        jm = p.lower().replace('.', '')
                        break
                
                # Zbieramy nazwę ze wszystkiego co jest pomiędzy kodem a ilościami/EANami
                nazwa_parts = []
                for p in parts[1:]:
                    if (p.isdigit() and len(p) >= 10) or p.lower() in ['szt', 'szt.', 'kg', 'kg.'] or p == parts[-1]:
                        break
                    nazwa_parts.append(p)
                
                nazwa_towaru = " ".join(nazwa_parts).upper()
                
                if ilosc_koncowa > 0:
                    w_opak = 10.0
                    for klucz, waga in PRZELICZNIKI_SIECI.items():
                        if klucz in nazwa_towaru:
                            w_opak = waga
                            break
                            
                    ilosc_op = round(ilosc_koncowa / w_opak, 1) if w_opak > 0 else 0
                    kraj = ustal_kraj_pochodzenia(nazwa_towaru)
                    
                    towary.append({
                        "kod": kod_part, "nazwa": nazwa_towaru if len(nazwa_towaru) > 2 else "TOWAR " + kod_part, 
                        "jm": jm, "w_opak": w_opak, "ilosc_op": ilosc_op, "ilosc_koncowa": ilosc_koncowa, "kraj": kraj
                    })
            except:
                pass

    if towary:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Dokument WZ"
        ws.views.sheetView[0].showGridLines = True
        
        font_title = Font(name="Arial", size=15, bold=True, color="1F497D")
        font_section = Font(name="Arial", size=10, bold=True, color="000000")
        font_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        font_body = Font(name="Arial", size=10, bold=False, color="000000")
        font_body_bold = Font(name="Arial", size=10, bold=True, color="000000")
        
        header_fill = PatternFill(start_color="1F497D", fill_type="solid")
        info_bar_fill = PatternFill(start_color="E9EDF4", fill_type="solid")
        zebra_fill = PatternFill(start_color="F2F5F8", fill_type="solid")
        
        thin = Side(style='thin', color='D9D9D9')
        border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
        double_bottom = Border(top=Side(style='thin', color='D9D9D9'), bottom=Side(style='double', color='1F497D'))
        
        ws['A1'] = "DOKUMENT WZ"
        ws['A1'].font = font_title
        ws.merge_cells('G1:H1')
        ws['G1'] = f"Nr WZ: STOK/{nr_zam}"
        ws['G1'].font = font_body_bold
        ws['G1'].alignment = Alignment(horizontal="right")
        
        ws['D3'] = "DOSTAWCA:"
        ws['D3'].font = font_section
        ws['D4'] = "GPW JANMAR SP. Z O.O. SP. K."
        ws['D4'].font = font_body_bold
        ws['D5'] = "ul. Gołaśka 3/58, 30-619 Kraków"
        ws['D5'].font = font_body
        
        ws['G3'] = "ODBIORCA:"
        ws['G3'].font = font_section
        ws['G4'] = "STOKROTKA SP. Z O.O."
        ws['G4'].font = font_body_bold
        ws['G5'] = "ul. Projektowa 1, 20-209 Lublin"
        ws['G5'].font = font_body
        
        data_zamowienia = "................"
        data_dostawy = "................"
        miejsce_dostawy = "STOKROTKA CD"
        
        for line in linie:
            if "Termin dostawy" in line:
                match_dost = re.search(r'Termin dostawy\s+([\d\.]+)', line)
                match_wyst = re.
