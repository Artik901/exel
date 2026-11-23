"""Fill marketing decor names from naming schemes, leaving nondecorative goods untouched."""
from pathlib import Path
from zipfile import ZipFile
from collections import Counter
import re,os
from fill_colors import read_sheet, patch_row, normalized

def clean(s):
 s=normalized(s).strip(' ,;."')
 s=re.split(r'\s+(?:ECO|ЕСО)\s*\d',s,flags=re.I)[0]
 s=re.split(r'\s+(?:Ширина|шир\.?)\s+\d',s,flags=re.I)[0]
 s=re.split(r'\s+\d+(?:[.,]\d+)?\s*(?:мм|mm|см|метр|м2|м²)\b',s,flags=re.I)[0]
 s=re.split(r'\s+\(\s*\d',s)[0]
 s=re.split(r'\s+(?:1уп|1/|защ\.слой|защитный слой)\b',s,flags=re.I)[0]
 s=re.sub(r'\s+(?:Polished|Glossy|High Glossy|Carving|Matt|матовый|лаппатированный|структурная)\s*$','',s,flags=re.I)
 return normalized(s).strip(' ,;."')

def normalize_item(s):
 s=clean(s)
 s=re.sub(r'^(?:планка|плитка|декор)\s+','',s,flags=re.I)
 s=re.sub(r'^(?:[A-Z]{1,5}[\d-]{2,}[A-Z0-9-]*\s+)','',s,flags=re.I)
 s=re.sub(r'^\d{2,5}[A-ZА-Я]?-?\d*\s+','',s,flags=re.I)
 s=clean(s)
 if not s or len(s)<3 or len(s)>75:return ''
 if re.search(r'\b(?:купить|упаковк|шт|м2|класс|толщин|метра|покрытие)\b',s,re.I):return ''
 if not re.search(r'[а-яёa-z]',s,re.I):return ''
 return s

def choose(row):
 name=normalized(row.get('E',''))
 group=row.get('A','')
 g=group.lower()
 if not name:return ''
 floors=any(x in g for x in ['ламинат','линолеум','плитка замковая','плитка пвх','паркетная доска','керамогранит','стеновые панели','террасные системы'])
 if not floors:return ''
 # Explicit color key, especially decking/facade parts.
 x=re.search(r'\bцвет\s*:\s*([^,()]+)',name,re.I)
 if x:
  return normalize_item(x.group(1))
 # All flooring products with quoted collection after manufacturer.
 quoted=list(re.finditer(r'["«]([^"»]+)["»]',name))
 for q in reversed(quoted):
  candidate=name[q.end():]
  # don't interpret a quoted manufacturer as a collection
  if q.group(1).lower() in ['tarkett','синтерос','juteкs','juteks','ivc','ideal','grunholz','alpine floor']:continue
  if 'линолеум' in g:
   candidate=re.sub(r'^\s*'+re.escape(group.split('/')[-1])+r'\s+','',candidate,flags=re.I)
  val=normalize_item(candidate)
  if val and not re.match(r'^(?:\d+\s*кл\b|[A-Z]{0,3}\d+\s*$)',val,re.I):
   return val
 # Other linoleum brand/collection patterns e.g quoted "Tarkett" SPRINT PRO ARIZONA 1
 if 'линолеум' in g:
  last=group.split('/')[-1].strip()
  if last:
   m=re.search(re.escape(last)+r'\s+(.+?)\s+Ширина\b',name,re.I)
   if m:return normalize_item(m.group(1))
 # SPC, LVT, laminate group specific patterns around code and label.
 if any(x in g for x in ['плитка замковая','плитка пвх','ламинат']):
  # typically  "Wonderful LX 763-9 Винтер ("
  m=re.search(r'\b(?:[A-Z]{1,8}\s*)?\d{2,6}[A-ZА-Я]?(?:[-./]\w+)*\s+([A-Za-zА-Яа-яЁё][^()]{2,65})\s*\(',name)
  if m:
   v=normalize_item(m.group(1))
   if v:return v
  # Code following decor in end-position: "Avers Boden SPC Бирманский тик ABSPC824..."
  m=re.search(r'\b(?:SPC|LVT)\s+([А-ЯЁа-яё][А-ЯЁа-яё\s-]{2,40})\s+[A-Z]{2,}\d{2,}',name,re.I)
  if m:return normalize_item(m.group(1))
  # Zeta La casa Парма (...
  m=re.search(r'\b(?:La casa|Zeta La casa)\s+([А-ЯЁа-яё][А-ЯЁа-яё\s-]{2,40})\s*\(',name,re.I)
  if m:return normalize_item(m.group(1))
 # Ceramic design name after model SKU, avoid generic texture and specs
 if 'керамогранит' in g:
  m=re.search(r'\b(?:Primavera|Гранитея)\s+([A-ZА-Я]{1,4}\d{2,6}[A-ZА-Я]?)\s+(.+?)\s+\d{2,4}\s*[*×х]\s*\d',name,re.I)
  if m:return normalize_item(m.group(2))
  m=re.search(r'\b(?:Керамогранит)\s+[А-ЯЁа-яёA-Za-z]+\s+([А-ЯЁа-яё][А-ЯЁа-яё\s-]{2,40})\s+\d{2,4}\s*[*×х]\s*\d',name,re.I)
  if m:return normalize_item(m.group(1))
 # Panel decor after dimensions and color type.
 if 'стеновые панели' in g:
  m=re.search(r'\b(?:цветная|ламинированная)\s+(.+?)(?:\s*\(|,\s*ламин)',name,re.I)
  if m:return normalize_item(m.group(1))
 return ''

def main():
 p=next(x for x in Path('.').glob('*.xlsx') if not x.name.startswith('~$') and not x.name.endswith('_colors.xlsx'))
 with ZipFile(p) as z:
  sheet,data=read_sheet(z)
  assert data[1].get('BI')=='Доп. поле: Цвет'
  updates={};groups=Counter()
  for i,row in data.items():
   if i==1 or normalized(row.get('BI','')):continue
   c=choose(row)
   if c:
    updates[i]=c
    groups[row.get('A','').split('/')[1] if '/' in row.get('A','') else 'other']+=1
  old=z.read(sheet).decode('utf-8')
  new=re.sub(r'(<row\b[^>]*\br="(\d+)"[^>]*>)(.*?)(</row>)',lambda m:patch_row(m,updates),old,flags=re.S)
  tmp=p.with_name(p.stem+'_colors.xlsx')
  with ZipFile(tmp,'w') as out:
   for info in z.infolist():
    out.writestr(info,new.encode('utf-8') if info.filename==sheet else z.read(info.filename))
 with ZipFile(tmp) as check:
  _,data2=read_sheet(check)
  for i,c in updates.items():
   assert data2[i].get('BI')==c,(i,c,data2[i].get('BI'))
 os.replace(tmp,p)
 print('MARKETING_DECOR_NEW',len(updates),'BY_CATEGORY',dict(groups))
 print('SAMPLES',[(i,data[i].get('E','')[:115],c) for i,c in list(updates.items())[:85]])
 print('REMAINING',sum(not r.get('BI') for i,r in data2.items() if i>1))
if __name__=='__main__':main()
