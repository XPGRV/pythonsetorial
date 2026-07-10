"""Atualiza SECEX (Brasil) de BeefBR.xlsm pelo Comex Stat, via Drive."""
from __future__ import annotations
import datetime as dt
import sys,time
from copy import copy
from pathlib import Path
import openpyxl,requests
sys.path.insert(0,str(Path(__file__).resolve().parent.parent));import drive_io
XLSX_NAME='BeefBR.xlsm';SHEET_NAME='SECEX';URL='https://api-comexstat.mdic.gov.br/general';CODES=('0203','0201','0202','0207')
def period(v):
 if isinstance(v,dt.datetime):v=v.date()
 return (v.year,v.month) if isinstance(v,dt.date) else None
def api(code,start):
 payload={'flow':'export','monthDetail':True,'period':{'from':f'{start}-01','to':f'{dt.date.today().year}-12'},'filters':[{'filter':'heading','values':[code]}],'details':['heading'],'metrics':['metricFOB','metricKG']}
 for _ in range(4):
  r=requests.post(URL,json=payload,headers={'Content-Type':'application/json'},timeout=90)
  if r.status_code==429:time.sleep(25);continue
  r.raise_for_status();return r.json().get('data',{}).get('list',[])
 return []
def main():
 fid=drive_io.find_file_id(XLSX_NAME);path=drive_io.download_to_temp(fid);wb=openpyxl.load_workbook(path,keep_vba=True,data_only=False)
 try:
  ws=wb[SHEET_NAME];rows={period(ws.cell(r,2).value):r for r in range(3,ws.max_row+1) if period(ws.cell(r,2).value)}
  target=[p for p,r in rows.items() if p<=(dt.date.today().year,dt.date.today().month) and ws.cell(r,4).value in (None,'')]
  if not target:print('Comex Stat: sem lacunas.');return
  raw={c:{} for c in CODES}
  for c in CODES:
   for x in api(c,min(y for y,m in target)):raw[c][(int(x['year']),int(x['monthNumber']))]=(float(x.get('metricKG') or 0),float(x.get('metricFOB') or 0))
   time.sleep(10)
  changes=0
  for p in target:
   r=rows[p];a,b,d=raw['0203'].get(p,(0,0)),raw['0201'].get(p,(0,0)),raw['0202'].get(p,(0,0));f=raw['0207'].get(p,(0,0));beef=(b[0]+d[0],b[1]+d[1])
   vals={4:a[0]/1e6 if a[0] else None,5:beef[0]/1e6 if beef[0] else None,6:f[0]/1e6 if f[0] else None,8:a[1]/a[0] if a[0] else None,10:beef[1]/beef[0] if beef[0] else None,12:f[1]/f[0] if f[0] else None}
   for col,val in vals.items():
    if val is None:continue
    cell=ws.cell(r,col)
    if cell.value not in (None,''):continue
    cell.value=val
    if r>3:cell._style=copy(ws.cell(r-1,col)._style)
    changes+=1
   for col,src in ((9,8),(11,10),(13,12)):
    if ws.cell(r,col).value not in (None,''):continue
    ws.cell(r,col).value=f'={openpyxl.utils.get_column_letter(src)}{r}*$O{r}'
    changes+=1
  if changes:wb.save(path);drive_io.update_file(fid,path)
  print(f'Comex Stat BeefBR: {changes} celula(s) atualizada(s).')
 finally:wb.close()
if __name__=='__main__':main()
