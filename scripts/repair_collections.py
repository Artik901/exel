"""Normalize commercial decors in previously imported products.

This script updates ONLY BI ('Доп. поле: Цвет') for strictly recognized
collections. Other cells, columns, workbook metadata, and unknown colors stay intact.
"""
from pathlib import Path
from zipfile import ZipFile
from collections import Counter
from fill_colors import read_sheet, patch_row, normalized
import re
import os

def expected(row):
    name = normalized(row.get('E', ''))
    group = row.get('A', '').lower()
    low = name.lower()
    if not name:
        return '', ''

    # GrunHolz Danholm: '32 кл. 117B-11 Сона ( ... )'
    if 'grunholz' in low and re.search(r'["«]danholm["»]', name, re.I):
        m = re.search(r'["«]danholm["»]\s*\d{2}\s*кл\.?\s+[A-ZА-Я0-9-]+\s+(.+?)\s*\(', name, re.I)
        if m:
            return normalized(m.group(1)), 'GrunHolz Danholm'

    # GrunHolz Madeira: '32 кл. MAD172-13 Тауни 606*101*8мм'
    if 'grunholz' in low and re.search(r'["«]madeira["»]',name,re.I):
        m = re.search(r'\bMAD\d{2,4}-\d{1,3}\s+([А-ЯЁа-яё-]+)\s+\d{2,4}\s*[*xх×]',name,re.I)
        if m:
            return normalized(m.group(1)), 'GrunHolz Madeira'

    # Kastamonu Floorpan: FP719 is product code, following name is decor
    if 'kastamonu' in low and ('ламинат' in low or 'floorpan' in low):
        m = re.search(r'\bFP\d{3,5}[A-ZА-Я]?\s+(.+?)\s*\(\s*\d{3,5}\s*[*xх×]', name, re.I)
        if m:
            return normalized(m.group(1)), 'Kastamonu'

    # Avers Boden: 'ЦЕДР ABTE41. ( ... )', decor comes before SKU
    if 'avers boden' in low:
        m = re.search(r'\bAvers\s+Boden\s+([А-ЯЁа-яё][А-ЯЁа-яё\s-]{1,42}?)\s+AB[A-Z]{2,5}\d+(?=\b|\.)',name,re.I)
        if m:
            return normalized(m.group(1)), 'Avers Boden'

    # Aberhof Rhein: 'Bering (Беринг) 7771'; Russian display name is the decor
    if 'aberhof' in low and 'rhein' in low:
        m = re.search(r'["«]Rhein[^"»]*["»]\s+[A-Za-z][A-Za-z -]*\s*\(([А-ЯЁа-яё\s-]+)\)\s+\d{3,5}\s*\(',name,re.I)
        if m:
            return normalized(m.group(1)), 'Aberhof Rhein'

    # Noventis Granit: 'Айс 6007 ( ... )'
    if 'noventis' in low and 'granit' in low:
        m = re.search(r'["«]Granit["»]\s+([А-ЯЁа-яё][А-ЯЁа-яё\s-]{1,35}?)\s+\d{3,5}\s*\(',name,re.I)
        if m:
            return normalized(m.group(1)), 'Noventis Granit'

    # Tarkett LOUNGE/NEW AGE: 'ламинатная планка AMUSE'
    if 'tarkett' in low and re.search(r'["«](?:LOUNGE|NEW AGE)["»]',name,re.I):
        m = re.search(r'["«](?:LOUNGE|NEW AGE)["»]\s+(?:ламинатная\s+)?(?:планка|плитка)\s+([A-Z][A-Z0-9_-]+)\s*\(',name,re.I)
        if m:
            return normalized(m.group(1)), 'Tarkett LVT'

    # Primavera Torino: T1874 is code; 'Воздушный/Aire' is decor
    if 'primavera' in low and re.search(r'["«]Torino["»]',name,re.I):
        m = re.search(r'["«]Torino["»]\s*[TТ]\d{3,5}\s+(.+?)\s*\(',name,re.I)
        if m:
            return normalized(m.group(1)), 'Primavera Torino'

    # Linoleum PRAKTIKA: 'PRAKTIKA -- NUBIA 5 -- ширина 3м'
    if 'linoleum' in low or 'линолеум' in low:
        m = re.search(r'\s--\s*(.+?)\s*--\s*ширина',name,re.I)
        if m:
            return normalized(m.group(1)), 'PRAKTIKA'

    # Art East: 'AT 752 Бархат Ампир' / numeric code only '763 Тераццо Наполи'
    if 'art east' in low:
        for q in re.findall(r'["«]([^"»]+)["»]',name):
            m = re.fullmatch(r'\s*(?:(?:[A-ZА-Я]{1,5}\s*)?\d{2,6}(?:-\d+)?)\s+\*?\s*(.+?)\s*',q,re.I)
            if m and m.group(1).strip():
                return normalized(m.group(1)), 'Art East'

    return '', ''

def main():
    p = next(x for x in Path('.').glob('*.xlsx') if not x.name.startswith('~$') and not x.name.endswith('_colors.xlsx'))
    with ZipFile(p) as z:
        sheet, data = read_sheet(z)
        assert data[1].get('BI') == 'Доп. поле: Цвет'
        corrections = {}
        count = Counter()
        examples = []
        for i,row in data.items():
            if i == 1: continue
            desired, family = expected(row)
            old = normalized(row.get('BI',''))
            if desired and desired != old:
                corrections[i] = desired
                count[family] += 1
                examples.append((i, row.get('D',''), row.get('E','')[:115], old, desired))
        src = z.read(sheet).decode('utf-8')
        outxml = re.sub(r'(<row\b[^>]*\br="(\d+)"[^>]*>)(.*?)(</row>)',
                        lambda m: patch_row(m, corrections), src, flags=re.S)
        if corrections:
            output = p.with_name(p.stem + '_colors.xlsx')
            with ZipFile(output,'w') as dst:
                for item in z.infolist():
                    dst.writestr(item, outxml.encode('utf-8') if item.filename == sheet else z.read(item.filename))
    if corrections:
        with ZipFile(output) as z:
            _,verify = read_sheet(z)
            for i,val in corrections.items():
                assert verify[i].get('BI') == val, (i,val,verify[i].get('BI'))
        os.replace(output,p)
    print('COLLECTION_CORRECTED_TOTAL',len(corrections))
    print('COLLECTION_CORRECTED_BY_FAMILY',dict(count))
    for r in examples[:120]: print('CHANGE',repr(r))

    tests = [
        ('Ламинат GrunHolz "Danholm" 32 кл. 117B-11 Сона ( 1219*198*12мм)', 'Сона'),
        ('Ламинат GrunHolz "Danholm" 32 кл. L001 Валли ( 1219*198*12мм)', 'Валли'),
        ('Ламинат GrunHolz "Danholm" 32 кл. 187-1 Дюне ( 1219*198*12мм)', 'Дюне'),
        ('Ламинат Kastamonu Floorpan 33 кл. "Лагуна/Lagoon" FP713 Сосна Марсала (1380*193*8мм)', 'Сосна Марсала'),
        ('Ламинат Kastamonu Floorpan Lagoon 33кл. FP711 Вяз Марсель ( 1380x193x8мм)', 'Вяз Марсель'),
        ('Ламинированный паркет Avers Boden ЦЕДР ABTE41. (1218*168*10мм)', 'ЦЕДР'),
        ('Плитка кварц-виниловая Aberhof "Rhein CL" Bering (Беринг) 7771 (1220*151*4 мм)', 'Беринг'),
        ('Кварц-виниловая плитка (SPC) Noventis "Granit" Айс 6007 (1200*180*5мм)', 'Айс'),
        ('Плитка ПВХ LVT клеевая Tarkett "LOUNGE" ламинатная планка AMUSE (1280,2*182.9*3mm)', 'AMUSE'),
        ('Плитка ПВХ LVT клеевая Tarkett "NEW AGE" ламинатная планка ISIS (1280,2*182.9*2,1mm)', 'ISIS'),
        ('Плитка ПВХ LVT клеевая Primavera "Torino" Т1874 Воздушный/Aire (1227*187*2,5мм)', 'Воздушный/Aire'),
        ('Плитка ПВХ LVT клеевая Art East "AT 752 Бархат Ампир" (2.5х152,4х914,4mm)', 'Бархат Ампир'),
    ]
    for name, val in tests:
        actual,_=expected({'E':name})
        assert actual==val,(name,actual,val)
    print('REGRESSION_TESTS_PASSED',len(tests))
if __name__=='__main__':
    main()
