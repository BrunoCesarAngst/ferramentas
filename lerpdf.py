import os
import subprocess
import time
import re
import pdfplumber
import pandas as pd

# caminhos
pdf_path = r'C:\TKE\meutester\M8001339629_SLDDRW.pdf'
xls_path = r'C:\TKE\meutester\bom_completo.xlsx'

# cabeçalhos esperados e nomes finais
desired = {
    'ITEM':                      'item',
    'QTY':                       'qty',
    'TKE PART  NO.':             'part_no',
    'NAME':                      'name',
    'MFG DATA  SHEET  NOTES':    'mfg_data_sheet_notes',
    'ENG  DESIGN  SHEET  NOTES': 'eng_design_sheet_notes'
}

all_rows = []

with pdfplumber.open(pdf_path) as pdf:
    for page_no in (0, 1):    # só sheet 1 e 2
        tbl = pdf.pages[page_no].extract_table({
            'vertical_strategy':   'lines',
            'horizontal_strategy': 'lines'
        })
        if not tbl:
            continue
        # limpa células: sem None, newlines->espaço, strip
        clean = [
            [ (cell or '').replace('\n',' ').strip() for cell in row ]
            for row in tbl
        ]
        # localiza índice do cabeçalho (row que contém ITEM e QTY)
        header_idx = None
        for i, row in enumerate(clean):
            up = [c.upper() for c in row]
            if 'ITEM' in up and 'QTY' in up:
                header_idx = i
                break
        if header_idx is None:
            continue

        header = clean[header_idx]
        # normaliza nomes: uppercase e collapse spaces
        norm = [' '.join(h.upper().split()) for h in header]
        # mapeia cada coluna desejada ao seu índice
        col_idx = {}
        for raw, final in desired.items():
            key = ' '.join(raw.split())
            if key in norm:
                col_idx[final] = norm.index(key)
            else:
                raise ValueError(f'coluna "{raw}" não encontrada na página {page_no+1}')

        # varre linhas até chegar na data (p.ex. 2024-08-29)
        for row in clean[header_idx+1:]:
            if re.match(r'\d{4}-\d{2}-\d{2}', row[0].strip()):
                break
            # extrai só as colunas desejadas
            entry = { final: row[idx] for final, idx in col_idx.items() }
            all_rows.append(entry)

if not all_rows:
    raise RuntimeError('nenhuma linha extraída da BOM')

# DataFrame com colunas na ordem desejada
df = pd.DataFrame(all_rows, columns=desired.values())

# deleta xlsx anterior (mata Excel se estiver aberto)
if os.path.exists(xls_path):
    try:
        os.remove(xls_path)
    except PermissionError:
        subprocess.call(['taskkill','/f','/im','EXCEL.EXE'])
        time.sleep(1)
        os.remove(xls_path)

# salva e abre
df.to_excel(xls_path, index=False)
os.startfile(xls_path)
