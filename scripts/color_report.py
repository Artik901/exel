from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from collections import Counter
import re
p=next(p for p in Path('.').glob('*.xlsx') if not p.name.startswith('~$'))
S='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
with ZipFile(p) as z:
 ss=[]
 if 'xl/sharedStrings.xml' in z.namelist():
  root=ET.fromstring(z.read('xl/sharedStrings.xml'))
  ss=[''.join(t.text or '' for t in x.iter(S+'t')) for x in root.findall(S+'si')]
 root=ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
 rows=root.findall('.//'+S+'sheetData/'+S+'row')
 def val(c):
  v=c.find(S+'v')
  if c.attrib.get('t')=='inlineStr':return ''.join(x.text or '' for x in c.iter(S+'t'))
  if v is None:return ''
  if c.attrib.get('t')=='s':return ss[int(v.text)]
  return v.text or ''
 records=[]
 for row in rows[1:]:
  d={re.match(r'([A-Z]+)',c.attrib['r']).group(1):val(c) for c in row.findall(S+'c')}
  records.append(d)
 print('COUNT',len(records))
 print('BI',Counter(x.get('BI','') for x in records).most_common(40))
 print('BQ',Counter(x.get('BQ','') for x in records).most_common(80))
 print('GROUPS',Counter(x.get('A','') for x in records).most_common(30))
 print('EXAMPLE RECORDS WITH BQ',[(x.get('E','')[:85],x.get('BQ')) for x in records if x.get('BQ')][:35])
 print('EXAMPLES 60 NAMES',[(x.get('E','')[:105],x.get('A','')[:40]) for x in records[100:160]])
