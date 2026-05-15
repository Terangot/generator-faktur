import os, re, tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from rapidfuzz import fuzz
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Цвета бухгалтеров ──────────────────────────────────────────
COLORS = {
    'Aneta':                 ('1F497D', '2E75B6'),
    'Dominika':              ('375623', '70AD47'),
    'Angelika Olejniczak':   ('BE4B05', 'ED7D31'),
    'Jadwiga':               ('7F6000', 'FFC000'),
    'Iwona':                 ('3D1F6B', '7030A0'),
    'Agnieszka Orczykowska': ('820000', 'C00000'),
    'Anieszka Młynarska':    ('005E7C', '00B0F0'),
}
DEF_COLOR = ('404040', '808080')

THIN  = Side(style='thin', color='CCCCCC')
BRD   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

HEADERS = ['Lp', 'Klient', 'NIP', 'Numer faktury', 'Data wystawienia',
           'Termin płatności', 'Kwota netto', 'Kwota brutto', 'Waluta',
           'Zapłacona', 'Zatwierdzona', 'KSeF', 'Pomochnik']
WIDTHS  = [5, 40, 13, 16, 17, 17, 14, 14, 8, 12, 14, 14, 20]


# ── Утилиты ─────────────────────────────────────────────────────
def read_file(path, **kw):
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.xlsx', '.xls'):
        return pd.read_excel(path, **kw)
    for enc in ('cp1250', 'utf-8', 'utf-8-sig', 'latin-1'):
        try:
            return pd.read_csv(path, sep=';', encoding=enc, **kw)
        except (UnicodeDecodeError, ValueError):
            continue
    raise ValueError(f'Не удалось прочитать: {path}')

def extract_nip(k):
    m = re.search(r'NIP/PESEL:\s*(\d+)', str(k))
    return m.group(1) if m else ''

def extract_name(k):
    return re.sub(r'\s*NIP/PESEL:\s*\d+', '', str(k)).strip()

def normalize(name):
    n = str(name).upper().strip()
    n = re.sub(r'[\n\r]', ' ', n)
    for pat in [r'^\(JDG\s*IT\)\s*', r'^\(JDG\)\s*', r'^JDG\s+']:
        n = re.sub(pat, '', n)
    n = re.sub(r'[.,\-]', '', n)
    n = re.sub(r'SP[ÓO][Ł]KA\s+Z\s+OGRANICZON[AĄ]\s+ODPOWIEDZIALNO[SŚ]CI[AĄ]', 'SPZOO', n)
    n = re.sub(r'SP\s+Z\s+O\s+O\b', 'SPZOO', n)
    n = re.sub(r'SP\s+Z\s+OO\b', 'SPZOO', n)
    n = re.sub(r'\bPSA\b', 'SPZOO', n)
    return re.sub(r'\s+', ' ', n).strip()

def find_accountant(name, xl_list):
    core = normalize(name).replace('SPZOO', '').strip()
    best, rec = 0, None
    for xl_norm, r in xl_list:
        s = fuzz.WRatio(core, xl_norm.replace('SPZOO', '').strip())
        if s > best:
            best, rec = s, r
    if best >= 88:
        return rec['Buchhalter'], rec['Pomochnik']
    return '', ''

def parse_amount(v):
    try:
        return float(re.sub(r'[^\d,]', '', str(v)).replace(',', '.'))
    except:
        return 0.0


# ── Стили ───────────────────────────────────────────────────────
def hdr(ws, row, vals, bg, fg='FFFFFF', sz=10, h=20):
    fill = PatternFill('solid', fgColor=bg)
    font = Font(bold=True, color=fg, size=sz)
    ws.row_dimensions[row].height = h
    for c, v in enumerate(vals, 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.fill, cell.font, cell.border = fill, font, BRD
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ── Генерация Excel ──────────────────────────────────────────────
def build_excel(df, path):
    wb  = Workbook()
    wb.remove(wb.active)

    matched   = df[df['Buchhalter'] != ''].copy()
    not_found = df[df['Buchhalter'] == ''].copy()
    buchhalters = sorted(matched['Buchhalter'].unique())
    ncols = len(HEADERS)
    last_col = get_column_letter(ncols)

    # запоминаем строку RAZEM каждого листа для PODSUMOWANIE
    razem = {}

    # ── листы бухгалтеров ────────────────────────────────────────
    for buch in buchhalters:
        grp  = matched[matched['Buchhalter'] == buch].copy()
        dark, light = COLORS.get(buch, DEF_COLOR)
        ws   = wb.create_sheet(buch[:31])
        ws.sheet_view.showGridLines = False

        # баннер
        ws.merge_cells(f'A1:{last_col}1')
        c = ws['A1']
        c.value = f"Księgowy: {buch}  |  Faktur: {len(grp)}  |  Filtruj ▼ — RAZEM przeliczy się automatycznie"
        c.font  = Font(bold=True, size=12, color='FFFFFF')
        c.fill  = PatternFill('solid', fgColor=dark)
        c.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 26

        # заголовки + автофильтр
        hdr(ws, 2, HEADERS, light, h=22)
        data_start = 3
        data_end   = 2 + len(grp)

        ws.auto_filter.ref = f'A2:{last_col}{data_end}'
        set_widths(ws, WIDTHS)

        # данные
        alt = PatternFill('solid', fgColor='EEF2F7')
        for i, (_, row) in enumerate(grp.iterrows()):
            r    = data_start + i
            nstr = str(row.get('Kwota netto', ''))
            bstr = str(row.get('Kwota brutto', ''))
            curr = (re.search(r'[A-Z]{3}', nstr) or re.search(r'[A-Z]{3}', bstr))
            curr = curr.group() if curr else 'PLN'
            vals = [
                row.get('Lp', ''),
                row.get('Klient', ''),
                str(row.get('NIP', '')),
                row.get('Numer', ''),
                row.get('Data wystawienia', ''),
                row.get('Termin płatności', ''),
                parse_amount(nstr),
                parse_amount(bstr),
                curr,
                row.get('Zapłacona', ''),
                row.get('Zatwierdzona', ''),
                row.get('KSeF', ''),
                row.get('Pomochnik', ''),
            ]
            ws.row_dimensions[r].height = 16
            for col, val in enumerate(vals, 1):
                cell = ws.cell(row=r, column=col, value=val)
                if i % 2 == 1:
                    cell.fill = alt
                cell.border = BRD
                if col in (7, 8):
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                elif col in (1, 9, 10, 11, 12):
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    cell.alignment = Alignment(vertical='center', wrap_text=(col == 2))

        # RAZEM с SUBTOTAL
        tr = data_end + 1
        razem[buch] = (ws.title, tr)
        ws.row_dimensions[tr].height = 20
        rfill = PatternFill('solid', fgColor=light)
        for col in range(1, ncols + 1):
            cell = ws.cell(row=tr, column=col)
            cell.fill = rfill
            cell.font = Font(bold=True, color='FFFFFF', size=11)
            cell.border = BRD
            cell.alignment = Alignment(horizontal='center', vertical='center')

        ws['A' + str(tr)].value = 'RAZEM'
        g = get_column_letter(7)
        h = get_column_letter(8)
        a = get_column_letter(1)
        c7 = ws.cell(row=tr, column=7)
        c7.value = f'=SUBTOTAL(9,{g}{data_start}:{g}{data_end})'
        c7.number_format = '#,##0.00'
        c7.alignment = Alignment(horizontal='right', vertical='center')
        c8 = ws.cell(row=tr, column=8)
        c8.value = f'=SUBTOTAL(9,{h}{data_start}:{h}{data_end})'
        c8.number_format = '#,##0.00'
        c8.alignment = Alignment(horizontal='right', vertical='center')
        cb = ws.cell(row=tr, column=2)
        cb.value = f'=SUBTOTAL(103,{a}{data_start}:{a}{data_end})'

    # ── PODSUMOWANIE ─────────────────────────────────────────────
    ws_s = wb.create_sheet('PODSUMOWANIE', 0)
    ws_s.sheet_view.showGridLines = False

    ws_s.merge_cells('A1:E1')
    c = ws_s['A1']
    c.value = 'PODSUMOWANIE — aktualizuje się po filtrach'
    c.font  = Font(bold=True, size=13, color='FFFFFF')
    c.fill  = PatternFill('solid', fgColor='1F3864')
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws_s.row_dimensions[1].height = 28

    hdr(ws_s, 2, ['Księgowy', 'Widocznych faktur', 'Kwota netto', 'Kwota brutto', 'Arkusz'], '2E4057', h=22)

    for i, buch in enumerate(buchhalters):
        r = 3 + i
        dark, light = COLORS.get(buch, DEF_COLOR)
        fill = PatternFill('solid', fgColor=light)
        sh, tr = razem[buch]
        safe = f"'{sh}'" if (' ' in sh or '.' in sh) else sh

        row_vals = [
            (1, buch),
            (2, f'={safe}!B{tr}'),
            (3, f'={safe}!G{tr}'),
            (4, f'={safe}!H{tr}'),
            (5, sh),
        ]
        ws_s.row_dimensions[r].height = 18
        for col, val in row_vals:
            cell = ws_s.cell(row=r, column=col, value=val)
            cell.font   = Font(bold=(col == 1), color='FFFFFF', size=10)
            cell.fill   = fill
            cell.border = BRD
            cell.alignment = Alignment(horizontal='center', vertical='center')
            if col in (3, 4):
                cell.number_format = '#,##0.00'

    # NIE ZNALEZIONO строка в сводке
    nf_row = 3 + len(buchhalters)
    ws_s.row_dimensions[nf_row].height = 18
    for col, val in enumerate(['NIE ZNALEZIONO', len(not_found), '-', '-', '-'], 1):
        cell = ws_s.cell(row=nf_row, column=col, value=val)
        cell.font   = Font(bold=True, color='FFFFFF', size=10)
        cell.fill   = PatternFill('solid', fgColor='C00000')
        cell.border = BRD
        cell.alignment = Alignment(horizontal='center', vertical='center')

    set_widths(ws_s, [28, 18, 16, 16, 22])

    # ── NIE ZNALEZIONO лист ──────────────────────────────────────
    ws_nf = wb.create_sheet('NIE ZNALEZIONO')
    ws_nf.sheet_view.showGridLines = False
    ws_nf.merge_cells(f'A1:{last_col}1')
    c = ws_nf['A1']
    c.value = f'Faktury bez księgowego — {len(not_found)} pozycji'
    c.font  = Font(bold=True, size=12, color='FFFFFF')
    c.fill  = PatternFill('solid', fgColor='7B0000')
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws_nf.row_dimensions[1].height = 26
    hdr(ws_nf, 2, HEADERS, 'C00000', h=22)
    set_widths(ws_nf, WIDTHS)

    alt = PatternFill('solid', fgColor='FFF0F0')
    for i, (_, row) in enumerate(not_found.iterrows()):
        r    = 3 + i
        nstr = str(row.get('Kwota netto', ''))
        bstr = str(row.get('Kwota brutto', ''))
        curr = (re.search(r'[A-Z]{3}', nstr) or re.search(r'[A-Z]{3}', bstr))
        curr = curr.group() if curr else 'PLN'
        vals = [
            row.get('Lp', ''), row.get('Klient', ''), str(row.get('NIP', '')),
            row.get('Numer', ''), row.get('Data wystawienia', ''), row.get('Termin płatności', ''),
            parse_amount(nstr), parse_amount(bstr), curr,
            row.get('Zapłacona', ''), row.get('Zatwierdzona', ''), row.get('KSeF', ''), '',
        ]
        ws_nf.row_dimensions[r].height = 16
        for col, val in enumerate(vals, 1):
            cell = ws_nf.cell(row=r, column=col, value=val)
            if i % 2 == 1:
                cell.fill = alt
            cell.border = BRD
            if col in (7, 8):
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif col in (1, 9, 10, 11, 12):
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(vertical='center', wrap_text=(col == 2))

    wb.save(path)
    return len(matched), len(not_found), buchhalters


# ── GUI ──────────────────────────────────────────────────────────
class App:
    def __init__(self, root):
        self.root = root
        self.root.title('Generator raportów faktur')
        self.root.geometry('520x350')
        self.root.resizable(False, False)
        self.root.configure(bg='#E8EDF2')
        self.f1 = tk.StringVar()
        self.f2 = tk.StringVar()
        self._ui()

    def _ui(self):
        bg, btn, btn_act = '#E8EDF2', '#1A5CA8', '#0D3E7A'
        lbl = '#0A1F3C'

        tk.Label(self.root, text='Generator raportów faktur',
                 font=('Helvetica', 14, 'bold'), bg=bg, fg='#1A5CA8').pack(pady=(18, 4))
        tk.Label(self.root, text='Załaduj dwa pliki — otrzymasz gotowy raport Excel',
                 font=('Helvetica', 10), bg=bg, fg='#333').pack(pady=(0, 16))

        frm = tk.Frame(self.root, bg=bg)
        frm.pack(padx=30, fill='x')

        tk.Label(frm, text='1.  Lista faktur (.csv / .xlsx):',
                 font=('Helvetica', 11, 'bold'), bg=bg, fg=lbl,
                 anchor='w').grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 4))
        tk.Entry(frm, textvariable=self.f1, width=34, font=('Helvetica', 9),
                 bg='#FFF', fg='#111', relief='solid', bd=1).grid(row=1, column=0, sticky='ew', pady=(0, 8))
        tk.Button(frm, text='  Wybierz plik  ', command=self._pick1,
                  bg=btn, fg='white', activebackground=btn_act, activeforeground='white',
                  relief='raised', bd=2, pady=4, font=('Helvetica', 10, 'bold'),
                  cursor='hand2').grid(row=1, column=1, padx=(8, 0), pady=(0, 8))

        tk.Label(frm, text='2.  Lista klientów z księgowymi (.csv / .xlsx):',
                 font=('Helvetica', 11, 'bold'), bg=bg, fg=lbl,
                 anchor='w').grid(row=2, column=0, columnspan=2, sticky='w', pady=(4, 4))
        tk.Entry(frm, textvariable=self.f2, width=34, font=('Helvetica', 9),
                 bg='#FFF', fg='#111', relief='solid', bd=1).grid(row=3, column=0, sticky='ew', pady=(0, 4))
        tk.Button(frm, text='  Wybierz plik  ', command=self._pick2,
                  bg=btn, fg='white', activebackground=btn_act, activeforeground='white',
                  relief='raised', bd=2, pady=4, font=('Helvetica', 10, 'bold'),
                  cursor='hand2').grid(row=3, column=1, padx=(8, 0))

        frm.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(self.root, mode='indeterminate', length=300)
        self.status   = tk.StringVar()
        tk.Label(self.root, textvariable=self.status,
                 font=('Helvetica', 9), bg=bg, fg='#222').pack(pady=(10, 0))
        tk.Button(self.root, text='▶   Generuj raport', command=self._run,
                  bg='#0D3E7A', fg='white', activebackground='#092D5C', activeforeground='white',
                  font=('Helvetica', 12, 'bold'), relief='raised', bd=3,
                  padx=28, pady=12, cursor='hand2').pack(pady=(14, 0))

    def _pick(self, var):
        p = filedialog.askopenfilename(
            filetypes=[('Obsługiwane pliki', '*.csv *.xlsx *.xls'), ('All files', '*.*')])
        if p:
            var.set(p)

    def _pick1(self): self._pick(self.f1)
    def _pick2(self): self._pick(self.f2)

    def _run(self):
        p1, p2 = self.f1.get().strip(), self.f2.get().strip()
        if not p1 or not os.path.exists(p1):
            messagebox.showerror('Błąd', 'Wskaż plik z fakturami!')
            return
        if not p2 or not os.path.exists(p2):
            messagebox.showerror('Błąd', 'Wskaż plik z klientami!')
            return

        self.status.set('Przetwarzanie...')
        self.progress.pack(pady=4)
        self.progress.start(10)
        self.root.update()

        try:
            df_inv = read_file(p1)
            df_inv = df_inv.drop(columns=[c for c in df_inv.columns if 'Unnamed' in c], errors='ignore')
            df_inv['NIP']    = df_inv['Kontrahent'].apply(extract_nip)
            df_inv['Klient'] = df_inv['Kontrahent'].apply(extract_name)

            df_xl = read_file(p2, header=None)
            df_xl.columns = ['Klient', 'Buchhalter', 'Pomochnik']
            df_xl = df_xl[df_xl['Klient'].notna()]
            df_xl = df_xl[~df_xl['Klient'].astype(str).str.contains('бухг', na=False)]
            df_xl['Klient']     = df_xl['Klient'].astype(str).str.strip().str.lstrip('\n').str.strip()
            df_xl['Buchhalter'] = df_xl['Buchhalter'].astype(str).str.strip().str.title()
            df_xl = df_xl[df_xl['Klient'].str.len() > 1].reset_index(drop=True)
            xl_list = [(normalize(r['Klient']), r) for r in df_xl.to_dict('records')]

            df_inv['Buchhalter'] = ''
            df_inv['Pomochnik']  = ''
            for idx, row in df_inv.iterrows():
                b, p = find_accountant(row['Klient'], xl_list)
                df_inv.at[idx, 'Buchhalter'] = b
                df_inv.at[idx, 'Pomochnik']  = p

            desktop   = os.path.join(os.path.expanduser('~'), 'Desktop')
            out_path  = os.path.join(desktop, 'Faktury_po_buchhalterach.xlsx')
            matched, nf, accs = build_excel(df_inv, out_path)

            self.progress.stop()
            self.progress.pack_forget()
            self.status.set(f'✓ Gotowe!  Dopasowano: {matched}  |  Nie znaleziono: {nf}')
            messagebox.showinfo('Gotowe!',
                f'Raport zapisany na Pulpicie:\n{out_path}\n\n'
                f'Faktur dopasowanych: {matched}\n'
                f'Bez księgowego: {nf}\n\n'
                f'Księgowi: {", ".join(accs)}')

        except Exception as e:
            self.progress.stop()
            self.progress.pack_forget()
            self.status.set('Błąd!')
            import traceback
            messagebox.showerror('Błąd', f'{e}\n\n{traceback.format_exc()[-500:]}')


if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
