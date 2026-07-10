"""Consolidador BeefBR cloud: estende formulas sem Excel."""
from __future__ import annotations
import datetime as dt,re,sys
from copy import copy
from pathlib import Path
import openpyxl
sys.path.insert(0,str(Path(__file__).resolve().parent.parent));import drive_io
XLSX_NAME='BeefBR.xlsm';REF=re.compile(r'(?<![A-Za-z0-9_])(\$?[A-Za-z]{1,3})(\$?)(\d+)')
def shift(v,d):
 if not isinstance(v,str) or not v.startswith('='):return v
 return REF.sub(lambda m:m.group(0) if m.group(2)=='$' else f'{m.group(1)}{m.group(2)}{int(m.group(3))+d}',v)
def copy_cell(s,t,d):
 t.value=shift(s.value,d);t._style=copy(s._style);t.number_format=s.number_format;t.font=copy(s.font);t.fill=copy(s.fill);t.border=copy(s.border);t.alignment=copy(s.alignment)
def last_formula(ws,col,start=5):
 for r in range(ws.max_row,start-1,-1):
  if isinstance(ws.cell(r,col).value,str) and ws.cell(r,col).value.startswith('='):return r
 return None
def extend(ws,src,end,ranges):
 if not src or end<=src:return 0
 for r in range(src+1,end+1):
  for a,b in ranges:
   for c in range(a,b+1):copy_cell(ws.cell(src,c),ws.cell(r,c),r-src)
 return end-src
def row_of(y,m):return 5+(y-2000)*12+m-1
def main():
 fid=drive_io.find_file_id(XLSX_NAME);path=drive_io.download_to_temp(fid);wb=openpyxl.load_workbook(path,keep_vba=True,data_only=False);changes=0
 try:
  beef,abates,bbg=wb['BeefBR'],wb['Abates'],wb['BBG_Dados']
  # BBG_Dados I:J acompanha a ultima data com FX em G.
  data=max((r for r in range(2,bbg.max_row+1) if bbg.cell(r,7).value not in (None,'')),default=0);changes+=extend(bbg,last_formula(bbg,9,2),data,((9,10),))
  # Spreads e FX ate o mes corrente.
  today=dt.date.today();target=row_of(today.year,today.month);last=min(x for x in (last_formula(beef,7),last_formula(beef,40)) if x);changes+=extend(beef,last,target,((4,4),(7,9),(11,13),(40,40)))
  # Preco/spread externo acompanha a ultima observacao em E.
  e_last=max((r for r in range(5,beef.max_row+1) if isinstance(beef.cell(r,5).value,(int,float)) and beef.cell(r,5).value>0),default=0);last=min(x for x in (last_formula(beef,6),last_formula(beef,14)) if x);changes+=extend(beef,last,e_last,((6,6),(14,19)))
  # IBGE: mesmo offset historico da planilha.
  ibge=max((r for r in range(39,abates.max_row+1) if isinstance(abates.cell(r,3).value,(int,float)) and abates.cell(r,3).value>0),default=0);changes+=extend(beef,last_formula(beef,22),ibge-34,((22,29),))
  # SIF: dois meses de defasagem.
  sif=max((r for r in range(38,abates.max_row+1) if isinstance(abates.cell(r,22).value,(int,float)) and abates.cell(r,22).value>0),default=0)
  if sif:
   d=abates.cell(sif,2).value
   if isinstance(d,dt.datetime):d=d.date()
   if isinstance(d,dt.date):
    d=(d.replace(day=1)-dt.timedelta(days=1)).replace(day=1)-dt.timedelta(days=1);changes+=extend(beef,last_formula(beef,32),row_of(d.year,d.month),((32,38),))
  if changes:wb.save(path);drive_io.update_file(fid,path)
  print(f'Consolidador BeefBR: {changes} linha(s) estendida(s).')
 finally:wb.close()
if __name__=='__main__':main()
