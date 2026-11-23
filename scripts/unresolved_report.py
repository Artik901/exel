from pathlib import Path
from zipfile import ZipFile
from collections import defaultdict,Counter
from fill_colors import read_sheet,decide,normalized
p=next(x for x in Path('.').glob('*.xlsx') if not x.name.startswith('~$'))
with ZipFile(p) as z: _,data=read_sheet(z)
missing=defaultdict(list)
for i,row in data.items():
 if i==1 or normalized(row.get('BI','')):continue
 group=row.get('A','')
 top='/'.join(group.split('/')[:3])
 missing[top].append((i,row.get('E',''),group))
print('UNRESOLVED GROUPS',len(missing))
for name,items in sorted(missing.items(),key=lambda x:-len(x[1]))[:55]:
 print('GROUP',repr(name),'COUNT',len(items))
 for i,n,g in items[:8]:print('  ',i,repr(n[:150]))
