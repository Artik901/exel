"""Fill 'Доп. поле: Метраж в упаковке' from explicit package-area text.

Preserves populated cells and all other columns. Only square-metre areas
associated with a package description are accepted. No dimensions or prices.
"""
from pathlib import Path
from zipfile import ZipFile
from collections import Counter
from html import escape
import re,os
from fill_colors import read_sheet,normalized,xlcol

AREA_UNIT=r'(?:м\s*(?:²|2|\^\s*2|кв\.?)|m\s*(?:²|2|\^\s*2|sq\.?\s*m))'
NUMBER=r'(?<![\d.,])(\d{1,2}(?:[,.]\d{1,5})?)(?![\d.,])'
AREA=re.compile(NUMBER+r'\s*'+AREA_UNIT+r'(?![A-Za-zА-Яа-яЁё\d])',re.I)
PACK=re.compile(r'(?:\b1\s*уп(?:ак(?:овка)?)?\.?\b|\b1\s*/\s*\d+\s*шт\b|\bуп(?:ак(?:овка)?)?\.?\s*[/:\-]\s*\d+\s*шт\b|\b\d+\s*шт(?:ук)?\s*(?:в\s+упаковке|/\s*уп))',re.I)

def area_from_name(name):
    s=normalized(name)
    if not s:return '', 'absent'
    hits=[]
    for m in AREA.finditer(s):
        raw=m.group(1).replace('.',',')
        try: value=float(raw.replace(',','.'))
        except ValueError:continue
        if not 0.1<=value<=30:continue
        if re.search(r'(?:цена|за)\s*$',s[max(0,m.start()-12):m.start()],re.I):continue
        # Strong pack relation: either explicit 1уп/area, or a pack quantity nearby
        before=s[max(0,m.start()-45):m.start()]
        after=s[m.end():m.end()+45]
        if re.search(r'(?:1\s*уп(?:аковка)?\.?\s*[/:\-]?\s*|в\s+упаковке\s*|упаковка\s*[:\-]?\s*)$',before,re.I):
            hits.append(raw);continue
        if PACK.search(before+' '+after):
            hits.append(raw)
    unique=set(hits)
    if len(unique)==1:return next(iter(unique)),'identified'
    if len(unique)>1:return '', 'conflicting'
    return '', 'absent'

def patch(m,updates,col):
    pre,row,body,post=m.group(1),int(m.group(2)),m.group(3),m.group(4)
    if row not in updates:return m.group(0)
    addr=f'{col}{row}'
    cell=f'<c r="{addr}" t="inlineStr"><is><t>{escape(updates[row],quote=False)}</t></is></c>'
    existing=re.search(r'<c\b(?=[^>]*\br="'+re.escape(addr)+r'")[^>]*(?:/>|>.*?</c>)',body,re.S)
    if existing:
        body=body[:existing.start()]+cell+body[existing.end():]
    else:
        inserted=False
        for match in re.finditer(r'<c\b[^>]*\br="([A-Z]+)'+str(row)+r'"',body):
            if xlcol(match.group(1))>xlcol(col):
                body=body[:match.start()]+cell+body[match.start():]
                inserted=True;break
        if not inserted:body+=cell
    return pre+body+post

def main():
    wb=next(p for p in Path('.').glob('*.xlsx') if not p.name.startswith('~$') and not p.name.endswith('_colors.xlsx'))
    with ZipFile(wb) as z:
        sheet,data=read_sheet(z)
        cols=[c for c,v in data[1].items() if normalized(v)=='Доп. поле: Метраж в упаковке']
        assert len(cols)==1,cols
        col=cols[0]
        edits={}
        stats=Counter()
        examples=[]
        for n,row in data.items():
            if n==1:continue
            if normalized(row.get(col,'')):
                stats['already-filled']+=1;continue
            if row.get('C') not in ('Товар','Модификация',''):
                stats['not-product']+=1;continue
            area,reason=area_from_name(row.get('E',''))
            stats[reason]+=1
            if area:
                edits[n]=area
                if len(examples)<35:examples.append((n,row.get('D',''),row.get('E','')[:110],area))
        old=z.read(sheet).decode('utf-8')
        new=re.sub(r'(<row\b[^>]*\br="(\d+)"[^>]*>)(.*?)(</row>)',lambda m:patch(m,edits,col),old,flags=re.S)
        output=wb.with_name(wb.stem+'_area.xlsx')
        if edits:
            with ZipFile(output,'w') as dst:
                for info in z.infolist():
                    dst.writestr(info,new.encode('utf-8') if info.filename==sheet else z.read(info.filename))
    if edits:
        with ZipFile(output) as z:
            _,check=read_sheet(z)
            for n,a in edits.items():assert check[n].get(col)==a,(n,a,check[n].get(col))
            for n,row in data.items():
                if n==1:continue
                if normalized(row.get(col,'')):assert check[n].get(col)==row.get(col),(n,'existing altered')
                for c,v in row.items():
                    if c!=col:assert check[n].get(c,'')==v,(n,c,'other value changed')
        os.replace(output,wb)
    tests=[
      ('Ламинат GrunHolz Danholm (1219*198*12мм), 1/8шт. 1уп/1,931 м2','1,931'),
      ('Плитка ПВХ 2,2304м2 1/16шт','2,2304'),
      ('SPC 2.74 м2. 1уп/10шт','2,74'),
      ('Плитка 1уп/8шт 2,131 m²','2,131'),
      ('Линолеум ширина 3,5м',''),
      ('Ламинат 1219*198*12мм',''),
      ('Ламинат цена 1200 руб за м2',''),
    ]
    for name,expected in tests:
        actual,_=area_from_name(name)
        assert actual==expected,(name,actual,expected)
    print('AREA_COLUMN',col)
    print('AREA_NEW',len(edits))
    print('AREA_STATS',dict(stats))
    print('AREA_EXAMPLES',examples)
    print('AREA_TESTS_PASSED',len(tests))

if __name__=='__main__':main()
