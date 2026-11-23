"""Correct Art East LVT decor names which are inside quotes.
The leading model code (e.g. AT 752) is not part of the decor.
Only BI values for unambiguous quoted Art East names are corrected.
"""
from pathlib import Path
from zipfile import ZipFile
import re, os
from fill_colors import read_sheet, patch_row, normalized

def desired(name):
    if not re.search(r'\bArt\s*East\b', name, re.I):
        return ''
    quotes = re.findall(r'["«]([^"»]+)["»]', name)
    for value in quotes:
        value = normalized(value)
        m = re.fullmatch(r'(?:[A-ZА-Я]{1,5}\s*[-]?\s*\d{2,6}[A-ZА-Я0-9-]*)\s+(.+)',value,re.I)
        if m:
            return m.group(1).strip()
    return ''

def main():
    p=next(x for x in Path('.').glob('*.xlsx') if not x.name.startswith('~$') and not x.name.endswith('_colors.xlsx'))
    with ZipFile(p) as z:
        sheet, data=read_sheet(z)
        changes={}
        report=[]
        for i,row in data.items():
            if i==1:continue
            name=row.get('E','')
            want=desired(name)
            current=normalized(row.get('BI',''))
            if want and current!=want:
                changes[i]=want
                report.append((i,name[:135],current,want))
        original=z.read(sheet).decode('utf-8')
        modified=re.sub(r'(<row\b[^>]*\br="(\d+)"[^>]*>)(.*?)(</row>)',lambda m:patch_row(m,changes),original,flags=re.S)
        if changes:
            output=p.with_name(p.stem+'_colors.xlsx')
            with ZipFile(output,'w') as dst:
                for item in z.infolist():
                    dst.writestr(item,modified.encode('utf-8') if item.filename==sheet else z.read(item.filename))
    if changes:
        with ZipFile(output) as check:
            _,result=read_sheet(check)
            for i,want in changes.items():
                assert result[i].get('BI')==want,(i,want,result[i].get('BI'))
        os.replace(output,p)
    print('ART_EAST_CORRECTED',len(changes))
    for row in report[:80]:print('CHANGE',repr(row))
    assert desired('Плитка ПВХ LVT клеевая Art East "AT 752 Бархат Ампир" (2.5х152,4х914,4mm)')=='Бархат Ампир'
if __name__=='__main__':main()
