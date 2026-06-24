import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from io import BytesIO
from datetime import datetime
import re

st.set_page_config(page_title="Janmar WZ Stokrotka", page_icon="🥦", layout="centered")

st.title("🥦 JANMAR WZ-Stokrotka Web v1.9")
st.subheader("Oficjalny dekoder zamówień z silnika komputerowego Mac")
st.write("Wgraj surowy plik tekstowy (.TXT) zamówienia ze Stokrotki, aby wygenerować oficjalny arkusz WZ.")

DANE_DOSTAWCY = {
    "nazwa": "GPW JANMAR SP. Z O.O. SP. K.",
    "adres_1": "ul. Gołaśka 3/58",
    "adres_2": "30-619 Kraków",
    "nip": "6793087742"
}

PRZELICZNIKI_STOKROTKA = {
    "BROKUŁ": 10.0, "KALAFIOR": 6.0, "KAPUSTA PEKIŃSKA": 10.0, "KAPUSTA BIAŁA": 10.0,
    "KAPUSTA CZERWONA": 10.0, "KAPUSTA WŁOSKA": 10.0, "KAPUSTA WCZESNA": 6.0,
    "KAPUSTA CZERWONA WCZESNA": 6.0, "CUKINIA": 5.0, "KALAREPA": 8.0, "KUKURYDZA": 30.0,
    "RZODKIEWKA": 10.0, "SAŁATA LODOWA": 10.0, "SAŁATA MASŁOWA": 12.0, "SELER NACIOWY": 16.0,
    "WŁOSZCZYZNA": 10.0, "ARBUZ": 20.0, "BRZOSKWINIE": 10.0, "MORELE": 5.0, "NEKTARYNY": 10.0,
    "CEBULA CZERWONA": 5.0, "CEBULA ŻÓTL": 10.0, "CEBULA ŻÓŁTA": 10.0, "MARCHEW": 10.0,
    "PIETRUSZKA": 5.0, "ZIEMNIAK": 15.0
}

def ustal_kraj_pochodzenia(nazwa_towaru):
    nazwa_upper = nazwa_towaru.upper()
    if any(x in nazwa_upper for x in ["POMARAŃCZE", "MANDARYNKI"]): return "HISZPANIA / TURCJA"
    if "ZIEMNIAKI IMPORTOWE" in nazwa_upper: return "CYPR / GRECJA"
    if "ARBUZ" in nazwa_upper: return "WŁOCHY / HISZPANIA"
    return "POLSKA"

def parsuj_zamowienie_stokrotka_oryginal(file_bytes):
    try: tekst = file_bytes.decode('windows-1250')
    except:
        try: tekst = file_bytes.decode('cp1250')
        except: tekst = file_bytes.decode('utf-8', errors='ignore')
        
    # KLUCZOWA POPRAWKA: Usuwamy ukryte znaki powrotu karetki \r dla systemu Linux
    tekst = tekst.replace('\r', '')
    lines = tekst.split('\n')
    
    nr_zamowienia = "Nieznany"
    data_zamowienia = "................"
    data_dostawy = "................"
    miejsce_dostawy = "STOKROTKA CD"
    
    for line in lines:
        if "ZAMÓWIENIE TOWARU NR" in line:
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
        if len(line_raw) < 55:  # Zabezpieczenie przed krótkimi liniami
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
    
    header_fill = PatternFill(start_color="1F497D", fill_type="solid")
    info_bar_fill = PatternFill(start_color="E9EDF4", fill_type="solid")
    zebra_fill = PatternFill(start_color="F2F5F8", fill_type="solid")
    
    thin = Side(style='thin', color='D9D9D9')
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    double_bottom = Border(top=Side(style='thin', color='D9D9D9'), bottom=Side(style='double', color='1F497D'))
    
    ws['A1'] = "DOKUMENT WZ"
    ws['A1'].font = font_title
    ws.merge_cells('G1:H1')
    ws['G1'] = "Nr WZ: ......................."
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
    
    ws.cell(row=7, column=1, value="Nr zamówienia Stokrotka:").font = font_body_bold
    ws.cell(row=7, column=3, value=nr_zam).font = font_body
    ws.cell(row=8, column=1, value="Data zamówienia:").font = font_body_bold
    ws.cell(row=8, column=3, value=data_zam).font = font_body
    ws.cell(row=9, column=1, value="Data dostawy:").font = font_body_bold
    ws.cell(row=9, column=3, value=data_dost).font = font_body
    
    ws.merge_cells('A11:H11')
    ws['A11'] = f" MIEJSCE DOSTAWY: {miejsce.upper()}"
    ws['A11'].font = font_body_bold
    ws['A11'].fill = info_bar_fill
    ws['A11'].alignment = Alignment(vertical="center")
    ws.row_dimensions[11].height = 24
    
    naglowki = ["Lp.", "Kod towaru", "Nazwa asortymentu", "Kraj pochodzenia", "Jm.", "W opak.", "Ilość op.", "Ilość szt./kg"]
    for col_idx, text in enumerate(naglowki, start=1):
        cell = ws.cell(row=14, column=col_idx, value=text)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[14].height = 24
        
    start_row = 15
    max_wiersz_siatki = 60
    
    for r in range(start_row, max_wiersz_siatki + 1):
        idx_towaru = r - start_row
        ws.row_dimensions[r].height = 22
        
        ws.cell(row=r, column=1, value=idx_towaru + 1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=4, value="POLSKA").alignment = Alignment(horizontal="center")
        
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.border = border_all
            cell.font = font_body
            if idx_towaru % 2 == 1: cell.fill = zebra_fill
                
        if idx_towaru < len(towary):
            t = towary[idx_towaru]
            ws.cell(row=r, column=2, value=t['kod']).alignment = Alignment(horizontal="center")
            ws.cell(row=r, column=3, value=t['nazwa']).alignment = Alignment(horizontal="left")
            ws.cell(row=r, column=4, value=t['kraj']).alignment = Alignment(horizontal="center")
            ws.cell(row=r, column=5, value=t['jm']).alignment = Alignment(horizontal="center")
            
            c_w_opak = ws.cell(row=r, column=6, value=t['w_opak'])
            c_w_opak.alignment = Alignment(horizontal="right")
            c_w_opak.number_format = '#,##0.0'
            
            c_op = ws.cell(row=r, column=7, value=t['ilosc_op'])
            c_op.alignment = Alignment(horizontal="right")
            c_op.number_format = '#,##0.0'
            
            c_koncowa = ws.cell(row=r, column=8, value=t['ilosc_koncowa'])
            c_koncowa.alignment = Alignment(horizontal="right")
            c_koncowa.number_format = '#,##0'

    sum_row = max_wiersz_siatki + 2
    ws.row_dimensions[sum_row].height = 24
    ws.cell(row=sum_row, column=6, value="RAZEM NETTO:").font = font_body_bold
    ws.cell(row=sum_row, column=6).alignment = Alignment(horizontal="right")
    
    sum_cell = ws.cell(row=sum_row, column=8, value=f"=SUM(H15:H{max_wiersz_siatki})")
    sum_cell.font = font_body_bold
    sum_cell.alignment = Alignment(horizontal="right")
    sum_cell.border = double_bottom
    sum_cell.number_format = '#,##0'
    
    footer_row = sum_row + 2
    ws.row_dimensions[footer_row].height = 22
    ws.cell(row=footer_row, column=1, value="DOKUMENT SPORZĄDZIŁ: .............................").font = font_body_bold
    ws.cell(row=footer_row, column=6, value="ILOŚĆ PALET EURO: ..................").font = font_body_bold
    ws.cell(row=footer_row, column=6).alignment = Alignment(horizontal="right")

    kraje_lista = '"POLSKA, HISZPANIA, HOLANDIA, PORTUGALIA, WŁOCHY, GRECJA, NIEMCY, FRANCJA, TURCJA, MAROKO"'
    dv = DataValidation(type="list", formula1=kraje_lista, allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"D15:D{max_wiersz_siatki}")
    
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 13
    ws.column_dimensions['C'].width = 38
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 8
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 10
    ws.column_dimensions['H'].width = 14
    
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

uploaded_file = st.file_uploader("Wgraj plik zamówienia Stokrotki (.TXT)", type=["txt"])

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    with st.spinner("Przetwarzanie pliku oryginalnym algorytmem..."):
        nr_zam, data_zam, data_dost, miejsce, towary = parsuj_zamowienie_stokrotka_oryginal(file_bytes)
        if towary:
            wz_excel = buduj_excel_stokrotka(nr_zam, data_zam, data_dost, miejsce, towary)
            st.success(f"Wygenerowano WZ Stokrotka nr: {nr_zam}")
            st.download_button(
                label="📥 Pobierz oficjalną WZ Stokrotka (Excel)",
                data=wz_excel,
                file_name=f"WZ_STOKROTKA_{nr_zam.replace('/', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Błąd odczytu. System nie odnalazł linii produktowych.")
