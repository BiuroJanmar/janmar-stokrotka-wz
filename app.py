import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from io import BytesIO
from datetime import datetime
import re

st.set_page_config(page_title="Janmar WZ Stokrotka", page_icon="🥦", layout="centered")

st.title("🥦 JANMAR WZ-Stokrotka Web v1.0")
st.subheader("Dedykowany dekoder zamówień TXT dla sieci Stokrotka")
st.write("Wgraj surowy plik tekstowy (.TXT) zamówienia ze Stokrotki, aby wygenerować oficjalny arkusz WZ.")

DANE_DOSTAWCY = {
    "nazwa": "GPW JANMAR SP. Z O.O. SP. K.",
    "adres_1": "ul. Gołaśka 3/58",
    "adres_2": "30-619 Kraków",
    "nip": "6793087742"
}

# Twarde mnożniki przeliczników opakowań wyciągnięte z Twojego agenta z Maca
PRZELICZNIKI_STOKROTKA = {
    "BROKUŁ": 10.0,
    "KALAFIOR": 6.0,
    "KAPUSTA PEKIŃSKA": 10.0,
    "KAPUSTA BIAŁA": 10.0,
    "KAPUSTA CZERWONA": 10.0,
    "KAPUSTA WŁOSKA": 10.0,
    "KAPUSTA WCZESNA": 6.0,
    "KAPUSTA CZERWONA WCZESNA": 6.0,
    "CUKINIA": 5.0,
    "KALAREPA": 8.0,
    "KUKURYDZA": 30.0,
    "RZODKIEWKA": 10.0,
    "SAŁATA LODOWA": 10.0,
    "SAŁATA MASŁOWA": 12.0,
    "SELER NACIOWY": 16.0,
    "WŁOSZCZYZNA": 10.0,
    "ARBUZ": 20.0,
    "BRZOSKWINIE": 10.0,
    "MORELE": 5.0,
    "NEKTARYNY": 10.0,
    "CEBULA CZERWONA": 5.0,
    "CEBULA ŻÓŁTA": 10.0,
    "MARCHEW": 10.0,
    "PIETRUSZKA": 5.0,
    "ZIEMNIAK": 15.0
}

def ustal_kraj_pochodzenia(nazwa_towaru):
    nazwa_upper = nazwa_towaru.upper()
    if any(x in nazwa_upper for x in ["POMARAŃCZE", "MANDARYNKI"]): return "HISZPANIA / TURCJA"
    if "ZIEMNIAKI IMPORTOWE" in nazwa_upper: return "CYPR / GRECJA"
    if "ARBUZ" in nazwa_upper: return "WŁOCHY / HISZPANIA"
    return "POLSKA"

def dekoduj_txt_stokrotka(file_bytes):
    try: tekst = file_bytes.decode('windows-1250')
    except: tekst = file_bytes.decode('utf-8', errors='ignore')
    
    lines = tekst.split('\n')
    
    nr_zamowienia = "Nieznany"
    data_zamowienia = datetime.today().strftime('%d.%m.%Y')
    data_dostawy = "................"
    miejsce_dostawy = "STOKROTKA CD"
    
    for line in lines:
        if "ZAMÓWIENIE TOWARU NR" in line:
            match = re.search(r'NR\s+([\d/]+)', line)
            if match: nr_zamowienia = match.group(1)
        if "Termin dostawy" in line:
            match_dost = re.search(r'Termin dostawy\s+([\d\.]+)', line)
            match_wyst = re.search(r'Data wystawienia\s+([\d\.]+)', line)
            if match_dost: data_dostawy = match_dost.group(1)
            if match_wyst: data_zamowienia = match_wyst.group(1)
        if "Towar należy dostarczyć:" in line:
            idx = lines.index(line)
            if idx + 1 < len(lines): miejsce_dostawy = lines[idx+1].strip()

    towary = []
    for line in lines:
        parts = line.split()
        if len(parts) >= 6 and parts[0].isdigit() and len(parts[0]) == 6:
            kod = parts[0]
            nazwa_parts = []
            jm_idx = -1
            for idx, p in enumerate(parts[1:], start=1):
                if p.lower() in ['szt.', 'kg', 'szt']:
                    jm_idx = idx
                    break
                if not (p.isdigit() and len(p) > 7): nazwa_parts.append(p)
                
            if jm_idx != -1:
                nazwa = " ".join(nazwa_parts).upper()
                jm = parts[jm_idx]
                try:
                    ilosc_koncowa = float(parts[jm_idx - 1].replace(',', '.'))
                    
                    w_opak = 10.0
                    for klucz, waga in PRZELICZNIKI_STOKROTKA.items():
                        if klucz in nazwa:
                            w_opak = waga
                            break
                            
                    ilosc_op = round(ilosc_koncowa / w_opak, 1) if w_opak > 0 else 0
                    if ilosc_koncowa > 0 and not nazwa.isdigit():
                        towary.append({
                            'kod': kod, 'nazwa': nazwa, 'jm': jm, 
                            'w_opak': w_opak, 'ilosc_op': ilosc_op, 'ilosc_koncowa': ilosc_koncowa,
                            'kraj': ustal_kraj_pochodzenia(nazwa)
                        })
                except: pass
                
    return nr_zamowienia, data_zamowienia, data_dostawy, miejsce_dostawy, towary

def buduj_excel_stokrotka(nr_zam, data_zam, data_dost, miejsce, towary):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dokument WZ"
    ws.views.sheetView[0].showGridLines = True
    
    font_title = Font(name="Arial", size=15, bold=True, color="1F497D")
    font_section = Font(name="Arial", size=10, bold=True, color="000000")
    font_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_body = Font(name="Arial", size=10, bold=False, color="000000")
    font_body_bold = Font(name="Arial", size=10, bold=True, color="000000")
    
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    info_bar_fill = PatternFill(start_color="E9EDF4", end_color="E9EDF4", fill_type="solid")
    zebra_fill = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")
    
    thin = Side(style='thin', color='D9D9D9')
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    double_bottom = Border(top=Side(style='thin', color='D9D9D9'), bottom=Side(style='double', color='1F497D'))
    
    ws['A1'] = "DOKUMENT WZ"
    ws['A1'].font = font_title
    ws.merge_cells('G1:H1')
    ws['G1'] = "Nr WZ: ......................."
    ws['G1'].font = font_body_bold
    ws['G1'].alignment = Alignment(horizontal="right")
    
    ws['A3'] = "DOSTAWCA:"
    ws['A3'].font = font_section
    ws['A4'] = "GPW JANMAR SP. Z O.O. SP. K."
    ws['A4'].font = font_body_bold
    ws['A5'] = "ul. Gołaśka 3/58, 30-619 Kraków"
    ws['A5'].font = font_body
    
    ws['E3'] = "ODBIORCA:"
    ws['E3'].font = font_section
    ws['E4'] = "STOKROTKA SP. Z O.O."
    ws['E4'].font = font_body_bold
    ws['E5'] = "ul. Projektowa 1, 20-209 Lublin"
    ws['E5'].font = font_body
    
    ws.cell(row=7, column=1, value="Nr zamówienia Stokrotka:").font = font_body_bold
    ws.cell(row=7, column=4, value=nr_zam).font = font_body
    ws.cell(row=8, column=1, value="Data zamówienia:").font = font_body_bold
    ws.cell(row=8, column=3, value=data_zam).font = font_body
    ws.cell(row=9, column=1, value="Data dostawy:").font = font_body_bold
    ws.cell(row=9, column=3, value=data_dost).font = font_body
    
    ws.merge_cells('A10:H10')
    ws['A10'] = f" MIEJSCE DOSTAWY: {miejsce.upper()}"
    ws['A10'].font = font_body_bold
    ws['A10'].fill = info_bar_fill
    ws['A10'].alignment = Alignment(vertical="center")
    ws.row_dimensions[10].height = 24
    
    naglowki = ["Lp.", "Kod towaru", "Nazwa asortymentu", "Kraj pochodzenia", "Jm.", "W opak.", "Ilość op.", "Ilość szt./kg"]
    for col_idx, text in enumerate(naglowki, start=1):
        cell = ws.cell(row=12, column=col_idx, value=text)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[12].height = 24
        
    start_row = 13
    for idx, t in enumerate(towary):
        r = start_row + idx
        ws.row_dimensions[r].height = 22
        wiersz = [idx + 1, t['kod'], t['nazwa'], t['kraj'], t['jm'], t['w_opak'], t['ilosc_op'], t['ilosc_koncowa']]
        for col_idx, val in enumerate(wiersz, start=1):
            cell = ws.cell(row=r, column=col_idx, value=val)
            cell.font = font_body
            cell.border = border_all
            if idx % 2 == 1: cell.fill = zebra_fill
            if col_idx in [1, 2, 4, 5]: cell.alignment = Alignment(horizontal="center")
            elif col_idx in [6, 7, 8]: cell.alignment = Alignment(horizontal="right")
            if col_idx in [6, 7, 8]: cell.number_format = '#,##0.0' if col_idx == 7 else '#,##0'

    # Podsumowania logistyczne Stokrotki (Wiersz Razem Netto)
    sum_row = start_row + len(towary)
    ws.row_dimensions[sum_row].height = 24
    ws.cell(row=sum_row, column=6, value="RAZEM NETTO:").font = font_body_bold
    ws.cell(row=sum_row, column=6).alignment = Alignment(horizontal="right")
    
    sum_cell = ws.cell(row=sum_row, column=8, value=f"=SUM(H13:H{sum_row-1})")
    sum_cell.font = font_body_bold
    sum_cell.alignment = Alignment(horizontal="right")
    sum_cell.border = double_bottom
    sum_cell.number_format = '#,##0'
    
    # Stopka dokumentu Stokrotki
    footer_row = sum_row + 2
    ws.row_dimensions[footer_row].height = 22
    ws.cell(row=footer_row, column=1, value="DOKUMENT SPORZĄDZIŁ: .............................").font = font_body_bold
    ws.cell(row=footer_row, column=6, value="ILOŚĆ PALET EURO: ..................").font = font_body_bold
    ws.cell(row=footer_row, column=6).alignment = Alignment(horizontal="right")

    kraje_lista = '"POLSKA, HISZPANIA, HOLANDIA, PORTUGALIA, WŁOCHY, GRECJA, NIEMCY, FRANCJA, TURCJA, MAROKO"'
    dv = DataValidation(type="list", formula1=kraje_lista, allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"D13:D{sum_row+15}")
    
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 13
    ws.column_dimensions['C'].width = 35
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 8
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 10
    ws.column_dimensions['H'].width = 14
    
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

uploaded_file = st.file_uploader("Załaduj oryginalny plik tekstowy Stokrotki (.TXT)", type=["txt"])

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    with st.spinner("Przetwarzanie dokumentu Stokrotki..."):
        nr_zam, data_zam, data_dost, miejsce, towary = dekoduj_txt_stokrotka(file_bytes)
        if towary:
            wz_excel = buduj_excel_stokrotka(nr_zam, data_zam, data_dost, miejsce, towary)
            st.success(f"Sukces! Wygenerowano WZ Stokrotka dla zamówienia nr: {nr_zam}")
            st.download_button(
                label="📥 Pobierz oficjalną WZ Stokrotka (Excel)",
                data=wz_excel,
                file_name=f"WZ_STOKROTKA_{nr_zam.replace('/', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Błąd odczytu pliku TXT. System nie odnalazł linii produktowych.")
