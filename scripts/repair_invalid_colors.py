"""Audit and repair obvious invalid BI color values. Keeps uncertain values untouched."""
from pathlib import Path
from zipfile import ZipFile
from fill_colors import read_sheet, patch_row, normalized
import re, os

def invalid(value):
 v=normalized(value)
 if not v:return False
 return bool(
    re.search(r'^\s*\(?\s*\d+(?:[.,]\d+)?\s*[*×xх]\s*\d',v,re.I)
    or re.search(r'\b(?:mm|мм)\s*\)?\s*$',v,re.I) and re.search(r'\d',v)
    or v.startswith('--')
    or re.search(r'\b\d+\s*[*×xх]\s*\d+',v,re.I)
 )
def candidate(row):
 name=normalized(row.get('E',''))
 group=row.get('A','').lower()
 if not name:return ''
 # Product naming with -- separators: PRAKTIKA -- NUBIA 5 -- ширина...
 m=re.search(r'\s--\s*(.*?)\s*--(?:\s|$)',name,re.I)
 if m:
  dec=normalized(m.group(1)).strip(' -"')
  if dec and not invalid(dec):return dec
 # Norland / other LVT: design and numeric pattern inside quotes,
 # unlike collection names where the quoted text is followed by decor.
 if any(s in group for s in ('плитка','ламинат','spc','lvt')):
  for q in re.findall(r'["«]([^"»]+)["»]',name):
   q=normalized(q)
   if re.fullmatch(r'[A-Za-zА-Яа-яЁё][\wА-Яа-яЁё -]*\s+\d{2,5}-\d{1,3}',q):
    return q
 # PVC panels: dimensions BEFORE actual marketing decor.
 if 'стеновые панели' in group:
  m=re.search(r'(?:\b\d{3,5}\s*[*×xх]\s*\d{2,4}\s*[*×xх]\s*\d{1,3}\s*мм)\s+(.*)',name,re.I)
  if m:
   s=m.group(1).strip(' ,.')
   s=re.sub(r'^(?:цветная|ламинированная)\s+','',s,flags=re.I)
   s=re.sub(r',?\s*ламин\.?\s*$','',s,flags=re.I)
   s=re.sub(r'\s+\([^)]*\)\s*$','',s)
   s=normalized(s)
   if s and not invalid(s) and len(s)<75:return s
 return ''

def main():
 p=next(x for x in Path('.').glob('*.xlsx') if not x.name.startswith('~$') and not x.name.endswith('_colors.xlsx'))
 with ZipFile(p) as z:
  sheet,data=read_sheet(z)
  assert data[1].get('BI')=='Доп. поле: Цвет'
  updates={}
  suspicious=[]
  fixed=[]
  for i,row in data.items():
   if i==1:continue
   old=normalized(row.get('BI',''))
   want=candidate(row)
   if want and (invalid(old) or not old):
    if want!=old:
     updates[i]=want
     fixed.append((i,row.get('E','')[:125],old,want))
   if invalid(old) and i not in updates:
    suspicious.append((i,row.get('E','')[:100],old))
  oldxml=z.read(sheet).decode('utf-8')
  newxml=re.sub(r'(<row\b[^>]*\br="(\d+)"[^>]*>)(.*?)(</row>)',lambda m:patch_row(m,updates),oldxml,flags=re.S)
  if updates:
   tmp=p.with_name(p.stem+'_colors.xlsx')
   with ZipFile(tmp,'w') as out:
    for info in z.infolist():
     out.writestr(info,newxml.encode('utf-8') if info.filename==sheet else z.read(info.filename))
 if updates:
  with ZipFile(tmp) as z:
   _,after=read_sheet(z)
   for i,v in updates.items():assert after[i].get('BI')==v,(i,v,after[i].get('BI'))
  os.replace(tmp,p)
 print('CORRECTED',len(updates))
 for e in fixed[:90]:print('FIX',repr(e))
 print('STILL_SUSPICIOUS',len(suspicious))
 for e in suspicious[:130]:print('REVIEW',repr(e))
 assert candidate({'E':'Бытовой линолеум | PRAKTIKA -- NUBIA 5 -- ширина 3м','A':'Публикуемые товары/Линолеум/PRAKTIKA'})=='NUBIA 5'
 assert candidate({'E':'Плитка ПВХ клеевая NORLAND Sigrid LVT "Baldr 1003-4" (2х184х1219mm)','A':'Публикуемые товары/Плитка ПВХ клеевая LVT/Norland'})=='Baldr 1003-4'
 assert candidate({'E':'Панель ПВХ ламинированная 2700*250*8мм Лен фисташковый, ламин.','A':'Публикуемые товары/Стеновые панели'})=='Лен фисташковый'

if __name__=='__main__':main()
