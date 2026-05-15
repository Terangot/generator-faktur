"""
Generator raportów faktur według księgowych
Загрузи CSV (lista faktur) и XLSX (lista klientów) — получи готовый Excel по бухгалтерам
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import re
import os
from rapidfuzz import fuzz
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────
#  MATCHING
# ─────────────────────────────────────────

def extract_nip(k):
    m = re.search(r'NIP/PESEL:\s*(\d+)', str(k))
    return m.group(1) if m else ''

def extract_name(k):
    return re.sub(r'\s*NIP/PESEL:\s*\d+', '', str(k)).strip()

def normalize(name):
    name = str(name).upper().strip()
    name = re.sub(r'[\n\r]', ' ', name)
    for pat in [r'^\(JDG\s*IT\)\s*', r'^\(JDG\)\s*', r'^JDG\s+']:
        name = re.sub(pat, '', name)
    name = re.sub(r'[.,\-]', '', name)
    name = re.sub(r'SP[ÓO][Ł]KA\s+Z\s+OGRANICZON[AĄ]\s+ODPOWIEDZIALNO[SŚ]CI[AĄ]', 'SPZOO', name)
    name = re.sub(r'SP\s+Z\s+O\s+O\b', 'SPZOO', name)
    name = re.sub(r'SP\s+Z\s+OO\b', 'SPZOO', name)
    name = re.sub(r'\bPSA\b', 'SPZOO', name)
    return re.sub(r'\s+', ' ', name).strip()

def find_accountant(csv_name, xl_norm_list):
    csv_core = normalize(csv_name).replace('SPZOO', '').strip()
    best_score, best_rec = 0, None
    for xl_norm, rec in xl_norm_list:
        score = fuzz.WRatio(csv_core, xl_norm.replace('SPZOO', '').strip())
        if score > best_score:
            best_score, best_rec = score, rec
    if best_score >= 88:
        return best_rec['Buchhalter'], best_rec['Pomochnik']
    return '', ''


# ─────────────────────────────────────────
#  EXCEL GENERATION
# ─────────────────────────────────────────

def parse_amount(val):
    try:
        return float(re.sub(r'[^\d,]', '', str(val)).replace(',', '.'))
    except:
        return 0.0

ACCOUNTANT_COLORS = {
    'Aneta':                 {'header': '1F497D', 'subheader': '2E75B6'},
    'Dominika':              {'header': '375623', 'subheader': '70AD47'},
    'Angelika Olejniczak':   {'header': 'BE4B05', 'subheader': 'ED7D31'},
    'Jadwiga':               {'header': '7F6000', 'subheader': 'FFC000'},
    'Iwona':                 {'header': '3D1F6B', 'subheader': '7030A0'},
    'Agnieszka Orczykowska': {'header': '820000', 'subheader': 'C00000'},
    'Anieszka Młynarska':    {'header': '005E7C', 'subheader': '00B0F0'},
}
DEFAULT_COLOR = {'header': '404040', 'subheader': '808080'}

thin = Side(style='thin', color='CCCCCC')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

INV_HEADERS = [
    'Lp', 'Klient (pełna nazwa)', 'NIP', 'Numer faktury',
    'Data wystawienia', 'Termin płatności',
    'Kwota netto', 'Kwota brutto', 'Waluta',
    'Zapłacona', 'Zatwierdzona', 'KSeF', 'Pomochnik/Asystent'
]
COL_WIDTHS = [5, 42, 13, 16, 17, 17, 15, 15, 8, 12, 14, 14, 22]


def set_col_widths(ws, widths):
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w

def write_header_row(ws, row_num, values, bg_color, font_color='FFFFFF', font_size=10, height=20):
    fill = PatternFill("solid", fgColor=bg_color)
    font = Font(bold=True, color=font_color, size=font_size)
    ws.row_dimensions[row_num].height = height
    for col, val in enumerate(values, 1):
        c = ws.cell(row=row_num, column=col, value=val)
        c.fill = fill
        c.font = font
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = BORDER

def write_invoice_rows(ws, grp, start_row):
    alt_fill = PatternFill("solid", fgColor="EEF2F7")
    for i, (_, row_data) in enumerate(grp.iterrows()):
        r = start_row + i
        netto_str = str(row_data.get('Kwota netto', ''))
        curr_match = re.search(r'[A-Z]{3}', netto_str)
        currency = curr_match.group() if curr_match else 'PLN'
        vals = [
            row_data.get('Lp', ''),
            row_data.get('Klient', ''),
            str(row_data.get('NIP', '')),
            row_data.get('Numer', ''),
            row_data.get('Data wystawienia', ''),
            row_data.get('Termin płatności', ''),
            re.sub(r'[A-Z]{3}', '', netto_str).strip(),
            re.sub(r'[A-Z]{3}', '', str(row_data.get('Kwota brutto', ''))).strip(),
            currency,
            row_data.get('Zapłacona', ''),
            row_data.get('Zatwierdzona', ''),
            row_data.get('KSeF', ''),
            row_data.get('Pomochnik', ''),
        ]
        ws.row_dimensions[r].height = 16
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=r, column=col, value=str(val) if val is not None else '')
            if i % 2 == 1:
                c.fill = alt_fill
            c.alignment = Alignment(vertical='center', wrap_text=(col == 2))
            c.border = BORDER
            if col in (1, 9, 10, 11, 12):
                c.alignment = Alignment(horizontal='center', vertical='center')


def generate_excel(df_inv, output_path):
    wb = Workbook()
    wb.remove(wb.active)

    matched_df  = df_inv[df_inv['Buchhalter'] != ''].copy()
    not_found_df = df_inv[df_inv['Buchhalter'] == ''].copy()
    accountants = sorted(matched_df['Buchhalter'].unique())
    ncols = len(INV_HEADERS)

    # ── PODSUMOWANIE ──
    ws_sum = wb.create_sheet("PODSUMOWANIE")
    ws_sum.sheet_view.showGridLines = False
    ws_sum.merge_cells('A1:F1')
    c = ws_sum['A1']
    c.value = 'PODSUMOWANIE — Faktury według księgowych'
    c.font = Font(bold=True, size=14, color='FFFFFF')
    c.fill = PatternFill("solid", fgColor='1F3864')
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws_sum.row_dimensions[1].height = 28
    write_header_row(ws_sum, 2,
        ['Księgowy', 'Liczba faktur', 'Kwota netto (PLN)', 'Kwota brutto (PLN)',
         'Kwota netto (EUR)', 'Kwota brutto (EUR)'], '2E4057', height=22)

    row = 3
    for buch in accountants:
        grp = matched_df[matched_df['Buchhalter'] == buch]
        pln = grp[grp['Kwota netto'].str.contains('PLN', na=False)]
        eur = grp[grp['Kwota netto'].str.contains('EUR', na=False)]
        clr = ACCOUNTANT_COLORS.get(buch, DEFAULT_COLOR)
        fill = PatternFill("solid", fgColor=clr['subheader'])
        vals = [buch, len(grp),
                f"{sum(parse_amount(v) for v in pln['Kwota netto']):,.2f}" if len(pln) else '-',
                f"{sum(parse_amount(v) for v in pln['Kwota brutto']):,.2f}" if len(pln) else '-',
                f"{sum(parse_amount(v) for v in eur['Kwota netto']):,.2f}" if len(eur) else '-',
                f"{sum(parse_amount(v) for v in eur['Kwota brutto']):,.2f}" if len(eur) else '-']
        for col, val in enumerate(vals, 1):
            c = ws_sum.cell(row=row, column=col, value=val)
            c.font = Font(bold=(col == 1), color='FFFFFF', size=10)
            c.fill = fill
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.border = BORDER
        ws_sum.row_dimensions[row].height = 18
        row += 1

    for col, val in enumerate(['NIE ZNALEZIONO', len(not_found_df), '-', '-', '-', '-'], 1):
        c = ws_sum.cell(row=row, column=col, value=val)
        c.font = Font(bold=True, color='FFFFFF', size=10)
        c.fill = PatternFill("solid", fgColor='C00000')
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = BORDER
    ws_sum.row_dimensions[row].height = 18
    set_col_widths(ws_sum, [28, 16, 20, 20, 18, 18])

    # ── Листы по бухгалтерам ──
    for buch in accountants:
        grp = matched_df[matched_df['Buchhalter'] == buch].copy()
        clr = ACCOUNTANT_COLORS.get(buch, DEFAULT_COLOR)
        ws = wb.create_sheet(buch[:31])
        ws.sheet_view.showGridLines = False

        pln = grp[grp['Kwota netto'].str.contains('PLN', na=False)]
        eur = grp[grp['Kwota netto'].str.contains('EUR', na=False)]
        netto_pln  = sum(parse_amount(v) for v in pln['Kwota netto'])
        brutto_pln = sum(parse_amount(v) for v in pln['Kwota brutto'])
        netto_eur  = sum(parse_amount(v) for v in eur['Kwota netto'])
        brutto_eur = sum(parse_amount(v) for v in eur['Kwota brutto'])
        summary = f"PLN: netto {netto_pln:,.2f} / brutto {brutto_pln:,.2f}"
        if netto_eur:
            summary += f"   |   EUR: netto {netto_eur:,.2f} / brutto {brutto_eur:,.2f}"

        ws.merge_cells(f'A1:{get_column_letter(ncols)}1')
        c = ws['A1']
        c.value = f"Księgowy: {buch}   |   Liczba faktur: {len(grp)}   |   {summary}"
        c.font = Font(bold=True, size=12, color='FFFFFF')
        c.fill = PatternFill("solid", fgColor=clr['header'])
        c.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 26

        write_header_row(ws, 2, INV_HEADERS, clr['subheader'], height=22)
        set_col_widths(ws, COL_WIDTHS)
        write_invoice_rows(ws, grp, 3)

        total_row = 3 + len(grp)
        ws.row_dimensions[total_row].height = 18
        total_fill = PatternFill("solid", fgColor=clr['subheader'])
        for col in range(1, ncols + 1):
            c = ws.cell(row=total_row, column=col)
            c.fill = total_fill
            c.font = Font(bold=True, color='FFFFFF')
            c.border = BORDER
            c.alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=total_row, column=1, value='RAZEM')
        ws.cell(row=total_row, column=7, value=f"{netto_pln:,.2f}" if netto_pln else (f"{netto_eur:,.2f}" if netto_eur else ''))
        ws.cell(row=total_row, column=8, value=f"{brutto_pln:,.2f}" if brutto_pln else (f"{brutto_eur:,.2f}" if brutto_eur else ''))

    # ── NIE ZNALEZIONO ──
    ws_nf = wb.create_sheet("NIE ZNALEZIONO")
    ws_nf.sheet_view.showGridLines = False
    ws_nf.merge_cells(f'A1:{get_column_letter(ncols)}1')
    c = ws_nf['A1']
    c.value = f"Faktury bez przypisanego księgowego — {len(not_found_df)} pozycji"
    c.font = Font(bold=True, size=12, color='FFFFFF')
    c.fill = PatternFill("solid", fgColor='7B0000')
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws_nf.row_dimensions[1].height = 26
    write_header_row(ws_nf, 2, INV_HEADERS, 'C00000', height=22)
    set_col_widths(ws_nf, COL_WIDTHS)

    alt_fill = PatternFill("solid", fgColor="FFF0F0")
    for i, (_, row_data) in enumerate(not_found_df.iterrows()):
        r = 3 + i
        netto_str = str(row_data.get('Kwota netto', ''))
        curr_match = re.search(r'[A-Z]{3}', netto_str)
        currency = curr_match.group() if curr_match else 'PLN'
        vals = [
            row_data.get('Lp', ''), row_data.get('Klient', ''), str(row_data.get('NIP', '')),
            row_data.get('Numer', ''), row_data.get('Data wystawienia', ''),
            row_data.get('Termin płatności', ''),
            re.sub(r'[A-Z]{3}', '', netto_str).strip(),
            re.sub(r'[A-Z]{3}', '', str(row_data.get('Kwota brutto', ''))).strip(),
            currency, row_data.get('Zapłacona', ''), row_data.get('Zatwierdzona', ''),
            row_data.get('KSeF', ''), '',
        ]
        ws_nf.row_dimensions[r].height = 16
        for col, val in enumerate(vals, 1):
            c = ws_nf.cell(row=r, column=col, value=str(val) if val is not None else '')
            if i % 2 == 1:
                c.fill = alt_fill
            c.alignment = Alignment(vertical='center', wrap_text=(col == 2))
            c.border = BORDER
            if col in (1, 9, 10, 11, 12):
                c.alignment = Alignment(horizontal='center', vertical='center')

    wb.save(output_path)
    return len(matched_df), len(not_found_df), accountants


# ─────────────────────────────────────────
#  GUI
# ─────────────────────────────────────────

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Generator raportów faktur")
        self.root.geometry("520x340")
        self.root.resizable(False, False)
        self.root.configure(bg='#F0F4F8')

        self.csv_path  = tk.StringVar(value='')
        self.xlsx_path = tk.StringVar(value='')

        self._build_ui()

    def _build_ui(self):
        bg = '#F0F4F8'
        btn_bg   = '#1A5CA8'   # тёмно-синий фон кнопок
        btn_fg   = '#FFFFFF'   # белый текст
        btn_active = '#0D3E7A' # ещё темнее при наведении
        run_bg   = '#0D3E7A'
        run_active = '#092D5C'

        # Title
        tk.Label(self.root, text="Generator raportów faktur", font=('Helvetica', 14, 'bold'),
                 bg=bg, fg='#1A5CA8').pack(pady=(18, 4))
        tk.Label(self.root, text="Załaduj dwa pliki — otrzymasz gotowy raport Excel",
                 font=('Helvetica', 9), bg=bg, fg='#444').pack(pady=(0, 16))

        frame = tk.Frame(self.root, bg=bg)
        frame.pack(padx=30, fill='x')

        # CSV row
        tk.Label(frame, text="1.  Lista faktur (.csv):", font=('Helvetica', 10, 'bold'),
                 bg=bg, fg='#222', anchor='w').grid(row=0, column=0, sticky='w', pady=4)
        tk.Entry(frame, textvariable=self.csv_path, width=36,
                 font=('Helvetica', 9), bg='#FFFFFF', fg='#111',
                 relief='solid', bd=1).grid(row=1, column=0, sticky='ew', pady=(0, 4))
        tk.Button(frame, text="Wybierz plik", command=self._pick_csv,
                  bg=btn_bg, fg=btn_fg, activebackground=btn_active, activeforeground=btn_fg,
                  relief='raised', bd=2, padx=10, pady=4,
                  font=('Helvetica', 9, 'bold'), cursor='hand2').grid(row=1, column=1, padx=(8, 0))

        # XLSX row
        tk.Label(frame, text="2.  Lista klientów z księgowymi (.xlsx):", font=('Helvetica', 10, 'bold'),
                 bg=bg, fg='#222', anchor='w').grid(row=2, column=0, sticky='w', pady=(12, 4))
        tk.Entry(frame, textvariable=self.xlsx_path, width=36,
                 font=('Helvetica', 9), bg='#FFFFFF', fg='#111',
                 relief='solid', bd=1).grid(row=3, column=0, sticky='ew', pady=(0, 4))
        tk.Button(frame, text="Wybierz plik", command=self._pick_xlsx,
                  bg=btn_bg, fg=btn_fg, activebackground=btn_active, activeforeground=btn_fg,
                  relief='raised', bd=2, padx=10, pady=4,
                  font=('Helvetica', 9, 'bold'), cursor='hand2').grid(row=3, column=1, padx=(8, 0))

        frame.columnconfigure(0, weight=1)

        # Progress bar (hidden until run)
        self.progress = ttk.Progressbar(self.root, mode='indeterminate', length=300)

        # Status label
        self.status_var = tk.StringVar(value='')
        tk.Label(self.root, textvariable=self.status_var, font=('Helvetica', 9),
                 bg=bg, fg='#333').pack(pady=(10, 0))

        # Generate button
        tk.Button(self.root, text="▶  Generuj raport",
                  command=self._run,
                  bg=run_bg, fg=btn_fg, activebackground=run_active, activeforeground=btn_fg,
                  font=('Helvetica', 11, 'bold'), relief='raised', bd=3,
                  padx=24, pady=10, cursor='hand2').pack(pady=(12, 0))

    def _pick_csv(self):
        path = filedialog.askopenfilename(
            title="Wybierz plik CSV z fakturami",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if path:
            self.csv_path.set(path)

    def _pick_xlsx(self):
        path = filedialog.askopenfilename(
            title="Wybierz plik Excel z klientami",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if path:
            self.xlsx_path.set(path)

    def _run(self):
        csv_path  = self.csv_path.get().strip()
        xlsx_path = self.xlsx_path.get().strip()

        if not csv_path or not os.path.exists(csv_path):
            messagebox.showerror("Błąd", "Wskaż plik CSV z fakturami!")
            return
        if not xlsx_path or not os.path.exists(xlsx_path):
            messagebox.showerror("Błąd", "Wskaż plik Excel z klientami!")
            return

        self.status_var.set("Przetwarzanie...")
        self.progress.pack(pady=4)
        self.progress.start(10)
        self.root.update()

        try:
            # Load CSV
            df_inv = pd.read_csv(csv_path, sep=';', encoding='cp1250')
            df_inv = df_inv.drop(columns=[c for c in df_inv.columns if 'Unnamed' in c], errors='ignore')
            df_inv['NIP']    = df_inv['Kontrahent'].apply(extract_nip)
            df_inv['Klient'] = df_inv['Kontrahent'].apply(extract_name)

            # Load Excel
            df_xl = pd.read_excel(xlsx_path, header=None)
            df_xl.columns = ['Klient', 'Buchhalter', 'Pomochnik']
            df_xl = df_xl[df_xl['Klient'].notna()]
            df_xl = df_xl[~df_xl['Klient'].astype(str).str.contains('бухг', na=False)]
            df_xl['Klient']     = df_xl['Klient'].astype(str).str.strip().str.lstrip('\n').str.strip()
            df_xl['Buchhalter'] = df_xl['Buchhalter'].astype(str).str.strip().str.title()
            df_xl = df_xl[df_xl['Klient'].str.len() > 1].reset_index(drop=True)

            xl_norm_list = [(normalize(r['Klient']), r) for r in df_xl.to_dict('records')]

            # Match
            df_inv['Buchhalter'] = ''
            df_inv['Pomochnik']  = ''
            for idx, row in df_inv.iterrows():
                b, p = find_accountant(row['Klient'], xl_norm_list)
                df_inv.at[idx, 'Buchhalter'] = b
                df_inv.at[idx, 'Pomochnik']  = p

            # Output path — same folder as CSV
            out_dir  = os.path.dirname(csv_path)
            out_path = os.path.join(out_dir, 'Faktury_po_buchhalterach.xlsx')

            matched, not_found, accountants = generate_excel(df_inv, out_path)

            self.progress.stop()
            self.progress.pack_forget()
            self.status_var.set(f"✓ Gotowe!  Dopasowano: {matched}  |  Nie znaleziono: {not_found}")

            messagebox.showinfo("Gotowe!",
                f"Raport zapisany:\n{out_path}\n\n"
                f"Faktury dopasowane: {matched}\n"
                f"Bez księgowego: {not_found}\n\n"
                f"Księgowi: {', '.join(accountants)}")

        except Exception as e:
            self.progress.stop()
            self.progress.pack_forget()
            self.status_var.set("Błąd!")
            messagebox.showerror("Błąd", f"Wystąpił błąd:\n\n{e}")


if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
