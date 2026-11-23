"""Fill «Доп. поле: Цвет» conservatively from existing product data.
Only previously empty BI cells are changed. Requires only Python stdlib.
"""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET
from collections import Counter
from html import escape
import os
import re
import tempfile

S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
COLOR_WORDS = [
 ('без покрытия', 'Без покрытия'),
 ('светлое золото', 'Золото Светлое'),
 ('св. золото', 'Золото Светлое'),
 ('золото светлое', 'Золото Светлое'),
 ('золото', 'Золото'),
 ('шампань', 'Шампань'),
 ('серебрист', 'Серебристый'),
 ('серебро', 'Серебро'),
 ('графит', 'Графит'),
 ('антрацит', 'Антрацит'),
 ('черн', 'Черный'), ('чёрн', 'Черный'),
 ('бел', 'Белый'), ('бежев', 'Бежевый'),
 ('сер', 'Серый'), ('коричнев', 'Коричневый'),
 ('кремов', 'Кремовый'), ('песочн', 'Песочный'),
 ('терракот', 'Терракотовый'), ('розов', 'Розовый'),
 ('красн', 'Красный'), ('син', 'Синий'),
 ('голуб', 'Голубой'), ('зелен', 'Зеленый'), ('зелён', 'Зеленый'),
 ('желт', 'Желтый'), ('жёлт', 'Желтый'),
 ('бронз', 'Бронза'), ('медн', 'Медный'),
]
WOOD_OR_STONE = r'дуб|ясень|орех|венге|сосна|клен|клён|бук|береза|берёза|лиственница|гикори|акация|ольха|каштан|тик|секвойя|мрамор|бетон|сланец|гранит|травертин|оникс'
STOP = {'мм','м2','м²','плитка','ламинат','панель','плинтус','spc','lvt','паркет','м','см','класс','клеевой','замковый','декор','покрытие','напольный','полимерный','подложка','толщина'}
def normalized(s):
 return re.sub(r'\s+', ' ', str(s)).strip()

def decide(name, group, desc, bq, known):
 if bq.strip(): return (bq.strip(), 'characteristic')
 title=normalized(name)
 if not title: return ('','')
 # Most reliable: existing color/decor values in names of other products.
 matches=[x for x in known if len(x)>=5 and re.search(r'(?<!\w)'+re.escape(x)+r'(?!\w)',title,re.I)]
 if matches:
  return (max(matches,key=len),'existing-vocabulary')
 # Explicit finish/color labels and color words.
 for needle,out in COLOR_WORDS:
  # Terms with a stem intentionally allow suffixes; avoid matches inside larger words.
  if re.search(r'(?<![а-яa-z])'+re.escape(needle),title,re.I):
   return (out,'color-word')
 # Botanical/stone decor names with an unambiguous adjective/proper-name following.
 match=re.search(r'(?<!\w)('+WOOD_OR_STONE+r')\s+([А-ЯЁа-яё][А-ЯЁа-яё-]{2,})(?!\w)',title,re.I)
 if match:
  term=match.group(1)
  word=match.group(2)
  if word.lower() not in STOP and not word.lower().startswith(('ламинат','плитк','плинтус','паркет','натуральн')):
   return (term.capitalize()+' '+word.capitalize(),'decor')
 return ('','')

def xlcol(address):
 s=re.match(r'[A-Z]+',address).group()
 out=0
 for c in s:out=out*26+ord(c)-64
 return out

def read_sheet(z):
 strings=[]
 if 'xl/sharedStrings.xml' in z.namelist():
  root=ET.fromstring(z.read('xl/sharedStrings.xml'))
  strings=[''.join(x.text or '' for x in node.iter(S+'t')) for node in root.findall(S+'si')]
 sheet='xl/worksheets/sheet1.xml'
 root=ET.fromstring(z.read(sheet))
 rows=root.findall('.//'+S+'sheetData/'+S+'row')
 def cellval(c):
  if c.attrib.get('t')=='inlineStr': return ''.join(x.text or '' for x in c.iter(S+'t'))
  v=c.find(S+'v')
  if v is None: return ''
  if c.attrib.get('t')=='s': return strings[int(v.text)]
  return v.text or ''
 data={}
 for r in rows:
  fields={re.match(r'[A-Z]+',c.attrib['r']).group():cellval(c) for c in r.findall(S+'c')}
  data[int(r.attrib['r'])]=fields
 return sheet,data

def patch_row(m, updates):
 pre,row,body,post=m.group(1),int(m.group(2)),m.group(3),m.group(4)
 if row not in updates:return m.group(0)
 value=escape(updates[row],quote=False)
 cell=f'<c r="BI{row}" t="inlineStr"><is><t>{value}</t></is></c>'
 present=re.search(r'<c\b[^>]*\br="BI'+str(row)+r'"[^>]*(?:/>|>.*?</c>)',body,re.S)
 if present:
  return pre+body[:present.start()]+cell+body[present.end():]+post
 for cm in re.finditer(r'<c\b[^>]*\br="([A-Z]+)'+str(row)+r'"',body):
  if xlcol(cm.group(1))>xlcol('BI'):
   k=cm.start();return pre+body[:k]+cell+body[k:]+post
 return pre+body+cell+post

def main():
 p=next(x for x in Path('.').glob('*.xlsx') if not x.name.startswith('~$') and not x.name.endswith('_colors.xlsx'))
 with ZipFile(p) as z:
  sheet, data=read_sheet(z)
  assert data[1].get('BI')=='Доп. поле: Цвет', 'Expected BI color column'
  assert data[1].get('E')=='Наименование', 'Expected E product name column'
  known={normalized(r.get('BI','')) for i,r in data.items() if i>1 and r.get('BI')}
  known |= {normalized(r.get('BQ','')) for i,r in data.items() if i>1 and r.get('BQ')}
  updates={};stats=Counter()
  for i,row in data.items():
   if i==1:continue
   if normalized(row.get('BI','')):stats['already-filled']+=1;continue
   if row.get('C') not in ('Товар','Модификация',''):stats['not-product']+=1;continue
   value,reason=decide(row.get('E',''),row.get('A',''),row.get('X',''),row.get('BQ',''),known)
   if value:updates[i]=value;stats[reason]+=1
   else:stats['unresolved']+=1
  old=z.read(sheet).decode('utf-8')
  pattern=r'(<row\b[^>]*\br="(\d+)"[^>]*>)(.*?)(</row>)'
  revised=re.sub(pattern,lambda m:patch_row(m,updates),old,flags=re.S)
  if revised==old and updates:raise RuntimeError('No worksheet cells patched')
  if len(re.findall(r'<c\b[^>]*\br="BI\d+"',revised)) < len(updates):raise RuntimeError('Color cells missing after patch')
  tmp=p.with_name(p.stem+'_colors.xlsx')
  with ZipFile(tmp,'w') as out:
   for info in z.infolist():
    out.writestr(info,revised.encode('utf-8') if info.filename==sheet else z.read(info.filename))
 with ZipFile(tmp) as check:
  _,after=read_sheet(check)
  for i,color in updates.items():
   if after[i].get('BI')!=color:raise RuntimeError(f'Failed validation row {i}')
 os.replace(tmp,p)
 print('Workbook:',p.name,'Rows:',len(data)-1)
 print('New values:',len(updates),'Reason breakdown:',dict(stats))
 print('Examples:',[(i,data[i].get('E','')[:64],color) for i,color in list(updates.items())[:35]])

if __name__=='__main__':main()
