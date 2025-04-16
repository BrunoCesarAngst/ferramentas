import xml.etree.ElementTree as ET
import csv
import os
import re
import pandas as pd
import fitz

# caminhos
xml_file = r'C:\TKE\meutester\8000849648.pmx'
pdf_file = r"C:\TKE\meutester\Markup D8000849648-E.02.pdf"
csv_wc = r"C:\TKE\meutester\8000849648_BOX_SF_BACK_ASSY_LH_FULL_COP_F6_Design.xlsx"
saida_xlsx = r"C:\TKE\meutester\comparativo_bom_vs_modelo.xlsx"

# --- 1) parse xml e gera CSV ordenado ---
root = ET.parse(xml_file).getroot()
base = os.path.splitext(xml_file)[0]
csv_file = base + '_ordenado.csv'

dados = []
for grp in root.findall('.//group'):
    if grp.find('ref_pm') is not None:
        continue
    gname = grp.get('name', '')
    if not gname:
        continue


    item_num = revision = show = assembly_file = None
    for pv in grp.findall('./property_values/property_value'):
        name = pv.get('name', '')
        val = pv.text or ''
        if name == 'Item_Num':
            item_num = val
        elif name == 'Revision':
            revision = val
        elif name == 'Show':
            show = val
        elif name == 'AssemblyFile':
            assembly_file = val

    desc = partnum = qty = None
    for var in grp.findall('variable'):
        vname = var.get('name', '')
        calc_val = var.find('./property_values/property_value[@name="CalculatedValue"]')
        val = calc_val.text if calc_val is not None else ''
        if vname == 'DESC':
            desc = val
        elif vname == 'PARTNUM':
            partnum = val
        elif vname == 'QTY':
            qty = val

    dados.append({
        'item_num': item_num,
        'group_name': gname,
        'revision': revision,
        'show': show,
        'desc': desc,
        'partnum': partnum,
        'qty': qty,
        'assembly_file': assembly_file
    })

def conv_itemnum(num_str):
    try:
        return int(num_str)
    except:
        return 9999999999

dados.sort(key=lambda d: conv_itemnum(d['item_num']))

with open(csv_file, 'w', newline='', encoding='utf-8') as f:
    w = csv.writer(f, delimiter=';')
    w.writerow(['Item Num', 'Group Name', 'Revision', 'Show', 'Desc', 'Partnum', 'QTY', 'Assembly File'])
    for row in dados:
        w.writerow([row[k] or '' for k in ['item_num', 'group_name', 'revision', 'show', 'desc', 'partnum', 'qty', 'assembly_file']])

# --- 2) extrai texto do PDF ---
doc = fitz.open(pdf_file)
pdf_text = "\n".join([pg.get_text() for pg in doc])
lines = [line.strip() for line in pdf_text.split('\n') if line.strip()]

# --- 3) parser tolerante pro BOM do PDF ---
pdf_rows = []
i = 0
while i < len(lines) - 3:
    if re.fullmatch(r"\d{1,3}", lines[i]):
        try:
            item = int(lines[i])
            qty_raw = lines[i+1].strip()
            part_no = lines[i+2].strip()
            description = lines[i+3].strip()

            aviso = ''
            if part_no.strip().upper() == 'V' or qty_raw.strip().upper() == 'V':
                aviso = 'ver design sheet'

            qty = int(qty_raw) if qty_raw.isdigit() else None

            if re.match(r"(8\d{9}|V.*)", part_no):
                pdf_rows.append({
                    "item": item,
                    "qty": qty,
                    "part_no": part_no,
                    "description": description,
                    "aviso": aviso
                })
                i += 4
                continue
        except:
            pass
    i += 1

df_pdf = pd.DataFrame(pdf_rows)


# --- 4) modelo do xml ---
df_modelo = pd.read_csv(csv_file, sep=";", encoding="utf-8")
df_modelo.columns = df_modelo.columns.str.lower().str.strip()
df_modelo = df_modelo.rename(columns={"partnum": "part_no", "desc": "description"})
df_modelo["qty"] = df_modelo["qty"].astype(str).str.extract(r"(\d+)").astype(float)
df_modelo["part_no"] = df_modelo["part_no"].apply(lambda x: str(int(float(x))) if pd.notna(x) and str(x).replace('.', '', 1).isdigit() else str(x).strip())
df_modelo["description"] = df_modelo["description"].astype(str).str.strip()

# --- 5) modelo do WC ---
df_wc = pd.read_excel(csv_wc)
df_wc = df_wc.rename(columns=lambda x: x.strip())
df_wc = df_wc.rename(columns={
    "Find Number": "item",
    "Number": "part_no",
    "Name": "description",
    "Quantity": "qty",
    "Revision": "revision"
})
df_wc["qty"] = df_wc["qty"].astype(str).str.extract(r"(\d+)").astype(float)
df_wc["part_no"] = df_wc["part_no"].apply(lambda x: str(int(float(x)))if pd.notna(x) and str(x).replace('.', '', 1).isdigit() else str(x).strip())
df_wc["description"] = df_wc["description"].astype(str).str.strip()
df_wc["revision"] = df_wc["revision"].astype(str).str.strip()

# --- 6) função comparação ---
def compara_bom(df_a, df_b, suf_a, suf_b, check_revision=False):
    df = df_a.merge(df_b, on="part_no", how="outer", suffixes=(f"_{suf_a}", f"_{suf_b}"))

    df["match_desc"] = df.get(f"description_{suf_a}") == df.get(f"description_{suf_b}")
    df["match_qty"] = df.get(f"qty_{suf_a}") == df.get(f"qty_{suf_b}")

    if check_revision:
        df["revision_" + suf_a] = df.get("revision_" + suf_a, pd.NA)
        df["revision_" + suf_b] = df.get("revision_" + suf_b, pd.NA)
        df["match_revision"] = df[f"revision_{suf_a}"] == df[f"revision_{suf_b}"]
    else:
        df["match_revision"] = pd.NA

    def status(row):
        aviso_txt = row.get("aviso", "")
        if isinstance(aviso_txt, str) and "design" in aviso_txt.lower():
            return "ver design sheet"
        if pd.isna(row.get(f"description_{suf_b}")) or pd.isna(row.get(f"qty_{suf_b}")):
            return "faltando no " + suf_b
        elif not row["match_desc"] or not row["match_qty"]:
            return "divergente"
        else:
            return "ok"


    df["status"] = df.apply(status, axis=1)
    return df


df_pdf_vs_modelo = compara_bom(df_pdf, df_modelo, "pdf", "modelo")
df_pdf_vs_wc = compara_bom(df_pdf, df_wc, "pdf", "wc")
df_modelo_vs_wc = compara_bom(df_modelo, df_wc, "modelo", "wc", check_revision=True)

# --- 7) exporta resultados ---
with pd.ExcelWriter(saida_xlsx) as w:
    df_pdf_vs_modelo.to_excel(w, sheet_name="PDF vs Modelo", index=False)
    df_pdf_vs_wc.to_excel(w, sheet_name="PDF vs WC", index=False)
    df_modelo_vs_wc.to_excel(w, sheet_name="Modelo vs WC", index=False)

print("comparativo salvo em:", saida_xlsx)
