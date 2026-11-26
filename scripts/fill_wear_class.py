"""Populate the existing extra property 'Класс износостойкости' from product names.

Only previously blank cells are filled; no other property is modified.
Extracts a two-digit use/wear class explicitly accompanied by кл./класс,
not the abrasion rating AC4/AC5. Conflicting classes remain blank.
"""
from pathlib import Path
from zipfile import ZipFile
from html import escape
from collections import Counter
import re, os
from fill_colors import read_sheet, normalized, xlcol

VALID = {21,22,23,31,32,33,34,41,42,43}
PATTERNS = [
    re.compile(r'(?<![0-9A-Za-zА-Яа-яЁё])(\d{2})\s*(?:[-–]?\s*(?:й|го)\s*)?кл(?:асс)?\.?(?![A-Za-zА-Яа-яЁё])',re.I),
    re.compile(r'(?<![A-Za-zА-Яа-яЁё])класс(?:а)?\s*(?:износостойкости|применения|эксплуатации|помещений)?\s*[:№-]?\s*(\d{2})(?!\d)',re.I),
    re.compile(r'(?<![A-Za-zА-Яа-яЁё])кл\.?\s*[:№-]?\s*(\d{2})(?!\d)',re.I),
]

def extract(name):
    if re.search(r'(?<!\d)(?:2[1-3]|3[1-4]|4[1-3])\s*[/–-]\s*(?:2[1-3]|3[1-4]|4[1-3])\s*(?:кл\.?|класс)',name,re.I):
        return '', 'ambiguous'
    vals=set()
    for p in PATTERNS:
        for m in p.finditer(name):
            n=int(m.group(1))
            if n in VALID: vals.add(n)
    if len(vals)==1:return str(next(iter(vals))),'identified'
    if len(vals)>1:return '', 'ambiguous'
    return '', 'no-match'

def patch_generic(m, updates, col):
    pre,row,body,post=m.group(1),int(m.group(2)),m.group(3),m.group(4)
    if row not in updates:return m.group(0)
    address=f'{col}{row}'
    cell=f'<c r="{address}" t="inlineStr"><is><t>{escape(updates[row],quote=False)}</t></is></c>'
    existing=re.search(r'<c\b(?=[^>]*\br="'+re.escape(address)+r'")[^>]*(?:/>|>.*?</c>)',body,re.S)
    if existing:
        # Existing nonempty properties must never be replaced; caller checked this.
        return pre+body[:existing.start()]+cell+body[existing.end():]+post
    for found in re.finditer(r'<c\b[^>]*\br="([A-Z]+)'+str(row)+r'"',body):
        if xlcol(found.group(1))>xlcol(col):
            k=found.start()
            return pre+body[:k]+cell+body[k:]+post
    return pre+body+cell+post

def main():
    xlsx=next(p for p in Path('.').glob('*.xlsx') if not p.name.startswith('~$') and not p.name.endswith('_colors.xlsx'))
    with ZipFile(xlsx) as z:
        sheet,data=read_sheet(z)
        header=data[1]
        cols=[col for col,val in header.items() if normalized(val)=='Доп. поле: Класс износостойкости']
        assert len(cols)==1, f'Unable to uniquely locate wear class extra field: {cols}'
        col=cols[0]
        assert header.get('E')=='Наименование'
        changes={}
        stats=Counter()
        examples=[]
        for rownum,row in data.items():
            if rownum==1:continue
            if normalized(row.get(col,'')):
                stats['already-filled']+=1
                continue
            if row.get('C') not in ('Товар','Модификация',''):
                stats['not-product']+=1
                continue
            val,status=extract(normalized(row.get('E','')))
            stats[status]+=1
            if val:
                changes[rownum]=val
                if len(examples)<35:examples.append((rownum,row.get('D',''),row.get('E','')[:105],val))
        xml=z.read(sheet).decode('utf-8')
        revised=re.sub(r'(<row\b[^>]*\br="(\d+)"[^>]*>)(.*?)(</row>)',lambda m:patch_generic(m,changes,col),xml,flags=re.S)
        temp=xlsx.with_name(xlsx.stem+'_colors.xlsx')
        if changes:
            with ZipFile(temp,'w') as dst:
                for info in z.infolist():
                    dst.writestr(info,revised.encode('utf-8') if info.filename==sheet else z.read(info.filename))
    if changes:
        with ZipFile(temp) as z:
            _,check=read_sheet(z)
            for rownum,val in changes.items():
                assert check[rownum].get(col)==val, (rownum,col,val,check[rownum].get(col))
            for rownum,row in data.items():
                if rownum==1:continue
                if normalized(row.get(col,'')):
                    assert check[rownum].get(col)==row.get(col),(rownum,'existing cell changed')
        os.replace(temp,xlsx)
    testcases=[
      ('Ламинат GrunHolz Danholm 32 кл. 117B-11 Сона','32'),
      ('Ламинат Kastamonu 33кл. FP719 Дуб Маркес','33'),
      ('Ламинат 32 класс','32'),
      ('Ламинат 32кл.','32'),
      ('SPC 43 класс','43'),
      ('SPC класс 42','42'),
      ('Ламинат AC4',''),
      ('SPC 34/43 класс',''),
    ]
    for sample,expected in testcases:
        actual,_=extract(sample)
        assert actual==expected,(sample,actual,expected)
    print('CLASS_COLUMN',col)
    print('ADDED_WEAR_CLASSES',len(changes))
    print('STATS',dict(stats))
    print('EXAMPLES',examples)
    print('TESTS_PASSED',len(testcases))
if __name__=='__main__':main()
