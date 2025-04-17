import pdfplumber,pandas as pd
from collections import defaultdict
cols=["item","qty","tke part no.","name","mfg data sheet notes","eng design sheet notes"]
d=defaultdict(lambda:{c:""for c in cols})
with pdfplumber.open("M8001339629_SLDDRW.pdf")as pdf:
 for i,pg in enumerate(pdf.pages):
  words=pg.extract_words()
  if i==0:hpos={w["text"].lower():w["x0"]for w in words if w["text"].lower()in cols}
  for w in words:
   y=round(w["top"]);t=w["text"]
   for c,x in hpos.items():
    if abs(w["x0"]-x)<20:d[y][c]+=t+" "
df=pd.DataFrame(d.values())
df.to_excel("saida.xlsx",index=False)
print("salvo")
