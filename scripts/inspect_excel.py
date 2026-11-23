from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from collections import Counter
import re
p=next(p for p in Path('.').glob('*.xlsx') if not p.name.startswith('~$'))
S='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
with ZipFile(p) as z:
 ss=[]
 if 'xl/sharedStrings.xml' in z.namelist():
  root=ET.fromstring(z.read('xl/sharedStrings.xml'))
  ss=[''.join(t.text or '' for t in x.iter(S+'t')) for x in root.findall(S+'si')]
 wb=ET.fromstring(z.read('xl/workbook.xml'))
 rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
 remap={e.attrib['Id']:e.attrib['Target'].lstrip('/') for e in rels}
 for sht in wb.find(S+'sheets'):
  target=remap[sht.attrib[R+'id']]
  if not target.startswith('xl/'): target='xl/'+target
  root=ET.fromstring(z.read(target))
  rows=root.findall('.//'+S+'sheetData/'+S+'row')
  def val(c):
   v=c.find(S+'v')
   if c.attrib.get('t')=='inlineStr': return ''.join(x.text or '' for x in c.iter(S+'t'))
   if v is None:return ''
   if c.attrib.get('t')=='s': return ss[int(v.text)]
   return v.text or ''
  print('SHEET',sht.attrib.get('name'),'ROWS',len(rows))
  for row in rows[:6]:
   print('ROW',row.attrib.get('r'),[(c.attrib['r'],str(val(c))[:110]) for c in row.findall(S+'c')][:80])
  for row in rows[:10]:
   for c in row.findall(S+'c'):
    if 'Цвет' in str(val(c)): print('COLOR_HEADER_CANDIDATE',row.attrib.get('r'),c.attrib.get('r'),repr(val(c)))
