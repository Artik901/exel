"""Populate three existing extra properties from product names and category hierarchy.

Fields: Бренд, Фаска, Количество в упаковке.
Existing values are preserved; missing/uncertain information is not invented.
"""
from pathlib import Path
from zipfile import ZipFile
from html import escape
from collections import Counter,defaultdict
import re,os
from fill_colors import read_sheet, normalized, xlcol

BRANDS = [
 ("Alpine Floor",r"\balpine\s*floor\b"),
 ("Art East",r"\bart\s*east\b"),
 ("Tarkett",r"\btarkett\b"),
 ("Синтерос",r"\bсинтерос\b|\bsinteros\b"),
 ("Kastamonu",r"\bkastamonu\b"),
 ("Classen",r"\bclassen\b"),
 ("GrunHolz",r"\bgrun\s*holz\b|\bgrunholz\b"),
 ("Aberhof",r"\baberhof\b"),
 ("Norland",r"\bnorland\b"),
 ("Primavera",r"\bprimavera\b"),
 ("Noventis",r"\bnoventis\b"),
 ("Avers Boden",r"\bavers\s*boden\b"),
 ("PRAKTIKA",r"\bpraktika\b"),
 ("Juteks",r"\bjuteks\b"),
 ("IVC",r"(?<![a-z])ivc(?![a-z])"),
 ("Ideal",r"\bideal\b"),
 ("Salag",r"\bsalag\b"),
 ("Polivan",r"\bpolivan\b"),
 ("Hilst",r"\bhilst\b"),
 ("Wonderful",r"\bwonderful\b"),
 ("Zeta",r"\bzeta\b"),
 ("Art Stone",r"\bart\s*stone\b"),
 ("AquaFloor",r"\baqua\s*floor\b"),
 ("FineFloor",r"\bfine\s*floor\b"),
 ("Quick-Step",r"\bquick[\s-]*step\b"),
 ("Egger",r"\begger\b"),
 ("Kronospan",r"\bkronospan\b"),
 ("Kronostar",r"\bkronostar\b"),
 ("Kronotex",r"\bkronotex\b"),
 ("Kronopol",r"\bkronopol\b"),
 ("Swiss Krono",r"\bswiss\s*krono\b"),
 ("Woodstyle",r"\bwoodstyle\b"),
 ("Floorwood",r"\bfloorwood\b"),
 ("AGT",r"(?<![a-z])agt(?![a-z])"),
 ("Tritty",r"\btritty\b"),
 ("Haro",r"\bharo\b"),
 ("BerryAlloc",r"\bberry\s*alloc\b"),
 ("Pergo",r"\bpergo\b"),
 ("Camsan",r"\bcamsan\b"),
 ("Komitex",r"\bkomitex\b|\bкомитекс\b"),
 ("Sofit",r"\bsofit\b"),
 ("Fargo",r"\bfargo\b"),
 ("Stone Floor",r"\bstone\s*floor\b"),
 ("Vinilam",r"\bvinilam\b"),
 ("Decoria",r"\bdecoria\b"),
 ("Artens",r"\bartens\b"),
 ("My Step",r"\bmy\s*step\b"),
]
BRAND_REGEX=[(name,re.compile(regex,re.I)) for name,regex in BRANDS]

def brand(row):
    title=normalized(row.get('E',''))
    group=normalized(row.get('A',''))
    in_title=[(name,match.start()) for name,rx in BRAND_REGEX if (match:=rx.search(title))]
    if in_title:return min(in_title,key=lambda t:t[1])[0],'title'
    parts=group.split('/')
    # Category is not necessarily the brand; recognize a named segment only.
    for segment in parts[2:]:
        for name,rx in BRAND_REGEX:
            if rx.search(segment):return name,'group'
    return '','unknown'

def bevel(name):
    s=normalized(name)
    if not s:return '','unknown'
    if re.search(r'\bбез\s+(?:v[- ]?образной\s+)?(?:микро)?фаски\b|\bбесфасочн',s,re.I):
        return 'Без фаски','explicit'
    # 4V/2V is a bevel indicator when near the word bevel, or a standalone V notation
    m=re.search(r'\b(?:фаск[а-яё]*\s*[:,-]?\s*)([24])\s*V\b',s,re.I)
    if m:return m.group(1)+'V','explicit'
    m=re.search(r'\b([24])V\s*(?:[-—,]?\s*)?фаск',s,re.I)
    if m:return m.group(1)+'V','explicit'
    if re.search(r'\bфаск[а-яё]*\s+(?:с\s+)?(?:четыр[её]х|4)(?:[-\s]*(?:х|ёх))?[-\s]*сторон',s,re.I) or re.search(r'\b(?:четыр[её]х|4)(?:[-\s]?х)?сторонн[а-яё]*\s+фаск',s,re.I):
        return '4V','explicit'
    if re.search(r'\bфаск[а-яё]*\s+(?:с\s+)?(?:двух|2)\s*сторон',s,re.I):
        return '2V','explicit'
    if re.search(r'\bмикрофаск[а-яё]*',s,re.I):
        return 'Микрофаска','explicit'
    if re.search(r'\bфаск[а-яё]*',s,re.I):
        return 'С фаской','explicit'
    return '','unknown'

COUNT_PATTERNS=[
  re.compile(r'\b1\s*уп\.?\s*[/:\-]\s*(\d{1,3})\s*(?:шт(?:ук)?|план(?:ок|ки)|дос(?:ок|ки))\b',re.I),
  re.compile(r'\b1\s*уп(?:аковка)?\.?\s*[/:\-]?\s*(\d{1,3})\s*(?:шт(?:ук)?|план(?:ок|ки)|дос(?:ок|ки))\b',re.I),
  re.compile(r'(?<!\d)\b1\s*[/]\s*(\d{1,3})\s*шт\b',re.I),
  re.compile(r'\b(?:в\s+упаковке|в\s+пачке|упаковка\s*[:\-]?|кол[- ]?во\s+в\s+упаковке)\s*(\d{1,3})\s*(?:шт(?:ук)?|план(?:ок|ки)|дос(?:ок|ки))\b',re.I),
  re.compile(r'\b(?:по)\s+(\d{1,3})\s*шт\s*(?:в\s+упаковке|в\s+пачке)\b',re.I),
  re.compile(r'\b1\s*уп\.?\s*[/:\-]\s*(\d{1,3})\s*шт\b',re.I),
]
def pack_count(name):
    found=set()
    for p in COUNT_PATTERNS:
        for m in p.finditer(name):
            number=int(m.group(1))
            if 1<=number<=100:found.add(number)
    if len(found)==1:return str(next(iter(found))),'explicit'
    if len(found)>1:return '','ambiguous'
    return '','unknown'

def patch_row(m,updates):
    pre,row,body,post=m.group(1),int(m.group(2)),m.group(3),m.group(4)
    if row not in updates:return m.group(0)
    for col,value in sorted(updates[row].items(),key=lambda e:xlcol(e[0])):
        addr=f'{col}{row}'
        cell=f'<c r="{addr}" t="inlineStr"><is><t>{escape(value,quote=False)}</t></is></c>'
        match=re.search(r'<c\b(?=[^>]*\br="'+re.escape(addr)+r'")[^>]*(?:/>|>.*?</c>)',body,re.S)
        if match:
            body=body[:match.start()]+cell+body[match.end():]
            continue
        inserted=False
        for match in re.finditer(r'<c\b[^>]*\br="([A-Z]+)'+str(row)+r'"',body):
            if xlcol(match.group(1))>xlcol(col):
                body=body[:match.start()]+cell+body[match.start():]
                inserted=True
                break
        if not inserted:body+=cell
    return pre+body+post

def main():
    wb=next(x for x in Path('.').glob('*.xlsx') if not x.name.startswith('~$') and not x.name.endswith('_colors.xlsx'))
    with ZipFile(wb) as z:
        sheet,data=read_sheet(z)
        header=data[1]
        labels={'brand':'Доп. поле: Бренд','bevel':'Доп. поле: Фаска','pack':'Доп. поле: Количество в упаковке'}
        cols={}
        for key,label in labels.items():
            options=[col for col,v in header.items() if normalized(v)==label]
            assert len(options)==1,(label,options)
            cols[key]=options[0]
        changes=defaultdict(dict)
        stats=Counter()
        examples=defaultdict(list)
        for n,row in data.items():
            if n==1:continue
            if row.get('C') not in ('Товар','Модификация',''):
                stats['not-product']+=1
                continue
            values={'brand':brand(row),'bevel':bevel(row.get('E','')),'pack':pack_count(row.get('E',''))}
            for kind,(new,why) in values.items():
                col=cols[kind]
                if normalized(row.get(col,'')):
                    stats[kind+'_preserved']+=1
                elif new:
                    changes[n][col]=new
                    stats[kind+'_new']+=1
                    if len(examples[kind])<23:examples[kind].append((n,row.get('D',''),row.get('E','')[:98],new))
                else: stats[kind+'_'+why]+=1
        old=z.read(sheet).decode('utf-8')
        newxml=re.sub(r'(<row\b[^>]*\br="(\d+)"[^>]*>)(.*?)(</row>)',lambda m:patch_row(m,changes),old,flags=re.S)
        temp=wb.with_name(wb.stem+'_colors.xlsx')
        if changes:
            with ZipFile(temp,'w') as dst:
                for info in z.infolist():
                    dst.writestr(info,newxml.encode('utf-8') if info.filename==sheet else z.read(info.filename))
    if changes:
        with ZipFile(temp) as z:
            _,after=read_sheet(z)
            for n,fields in changes.items():
                for col,v in fields.items():assert after[n].get(col)==v,(n,col,v,after[n].get(col))
            for n,row in data.items():
                if n==1:continue
                for col in cols.values():
                    if normalized(row.get(col,'')):assert after[n].get(col)==row.get(col),(n,col,'overwritten')
            for n,row in data.items():
                if n==1:continue
                for col,val in row.items():
                    if col not in cols.values():assert after[n].get(col,'')==val,(n,col,'other field changed')
        os.replace(temp,wb)
    tests=[
      (pack_count('Ламинат, 1уп/8шт. 1,931 м2')[0],'8'),
      (pack_count('Ламинат 1/10шт')[0],'10'),
      (pack_count('Ламинат 1уп/1,931м2')[0],''),
      (pack_count('Линолеум ширина 3м')[0],''),
      (bevel('Ламинат, фаска 4V')[0],'4V'),
      (bevel('Ламинат без фаски')[0],'Без фаски'),
      (bevel('Ламинат Фаска с 4- сторон')[0],'4V'),
      (brand({'E':'Ламинат GrunHolz "Danholm"','A':'Публикуемые товары/Ламинат/GrunHolz'})[0],'GrunHolz'),
      (brand({'E':'SPC Alpine floor GRAND SEQUOIA','A':'Публикуемые товары/Плитка замковая SPC'})[0],'Alpine Floor'),
    ]
    for actual,expected in tests:assert actual==expected,(actual,expected)
    print('EXTRA_COLUMNS',cols)
    print('CHANGED_ROWS',len(changes))
    print('STATS',dict(stats))
    for kind,x in examples.items():print('EXAMPLES',kind,x)
    print('TESTS_PASSED',len(tests))

if __name__=='__main__':main()
