import re, numpy as np, pandas as pd
def fnum(x):
    s=('%g'%float(x)); return s
def grade_family(m):
    s=str(m).upper().replace(' ','')
    if s in ('NAN',''): return '?'
    if s.startswith('Q355') or s.startswith('Q345') or s.startswith('S355') or s.startswith('S335') or 'A709' in s or 'A572' in s or 'GR50' in s or 'SM490' in s or 'AS3678' in s or 'CAP345' in s or 'A500' in s or 'GR.C' in s or 'STKR490' in s or s=='GR345':
        return 'CĐ CAO (~345-355MPa)'
    if s.startswith('SS400') or s.startswith('Q235') or 'A36' in s or 'AS3679' in s or 'CT3' in s or 'STKR400' in s or 'GR300' in s or 'CB' in s:
        return 'THƯỜNG (~235-300MPa)'
    if 'S690' in s or 'STRENX' in s or 'Q460' in s: return 'SIÊU CAO (≥460MPa)'
    return 'KHÁC'
def stock_key(ten):
    t=str(ten)
    m=re.match(r'Thép tấm (?:lỡ cỡ[^\d]*)?(\d+(?:\.\d+)?)\s*mm\s*(.*)',t)
    if m: return 'PL',f'PL{fnum(m.group(1))}',m.group(2).replace('k.gửi','').strip()
    m=re.match(r'Thép tấm\s+(\d+(?:\.\d+)?)x(\d+)x(\d+),\s*([^,]+)',t)   # bravo
    if m: return 'PL',f'PL{fnum(m.group(1))}',m.group(4).strip()
    m=re.match(r'Thép ống[^\d]*(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)\s*(.*)',t)
    if m: return 'ỐNG',f'CHS{fnum(m.group(1))}x{fnum(m.group(2))}',m.group(3).strip()
    m=re.match(r'Thép (?:hình )?L/V\s*(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)[^\s,]*[\s,]*([^,]*)',t)
    if m: return 'HÌNH',f'L{fnum(m.group(1))}x{fnum(m.group(2))}x{fnum(m.group(3))}',m.group(4).strip()
    m=re.match(r'Thép I/H\s*(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)\s*(.*)',t)
    if m: return 'HÌNH',f'H{fnum(m.group(1))}x{fnum(m.group(2))}x{fnum(m.group(3))}x{fnum(m.group(4))}',m.group(5).strip()
    m=re.match(r'Thép (?:hình )?U(?:/C)?\s*(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)(?:x(\d+(?:\.\d+)?))?\s*(.*)',t)
    if m: return 'HÌNH',f'U{fnum(m.group(1))}x{fnum(m.group(2))}x{fnum(m.group(3))}',m.group(5).strip()
    m=re.match(r'Thép hộp(?: hàn)?\s*(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)\s*(.*)',t)
    if m: return 'HÌNH',f'SHS{fnum(m.group(1))}x{fnum(m.group(2))}x{fnum(m.group(3))}',m.group(4).strip()
    m=re.match(r'Thép tròn đặc(?: trơn)?\s*D(\d+(?:\.\d+)?)\s*(.*)',t)
    if m: return 'HÌNH',f'ROD{fnum(m.group(1))}',m.group(2).strip()
    m=re.match(r'Thép vuông đặc\s*(\d+)x(\d+)\s*(.*)',t)
    if m: return 'HÌNH',f'SQ{m.group(1)}x{m.group(2)}',m.group(3).strip()
    if 'HP' in t: 
        m=re.search(r'HP(\d+)x(\d+(?:\.\d+)?)',t); return 'HÌNH',f'HP{m.group(1)}x{fnum(m.group(2))}' if m else 'HP','CCS-A'
    return 'KHÁC',t,''
def demand_key(qc_part,qc_vt,day):
    p=str(qc_part).upper().replace(' ','').replace('\\','')
    v=str(qc_vt).upper()
    m=re.match(r'^PL\s*([\d.]+)',v)
    if m: return 'PL',f'PL{fnum(m.group(1))}'
    if p.startswith('PL'): return 'PL',f'PL{fnum(day)}' if pd.notna(day) else 'PL?'
    m=re.match(r'^(?:CHS|PIPE|O|Ø)-?D?([\d.]+)[X\*]([\d.]+)',p)
    if m:
        a,b=float(m.group(1)),float(m.group(2))
        if a<b and b>=1000: return 'PL',f'PL{fnum(a)}'   # ống lốc từ tấm: CHS-dày x đường kính
        return 'ỐNG',f'CHS{fnum(a)}x{fnum(b)}'
    m=re.match(r'^L-?([\d.]+)[X\*]([\d.]+)$',p)
    if m: return 'HÌNH',f'L{fnum(m.group(1))}x{fnum(m.group(1))}x{fnum(m.group(2))}'
    m=re.match(r'^L-?([\d.]+)[X\*]([\d.]+)[X\*]([\d.]+)',p)
    if m: return 'HÌNH',f'L{fnum(m.group(1))}x{fnum(m.group(2))}x{fnum(m.group(3))}'
    m=re.match(r'^(?:H|I|HN|HW|HM|W)-?([\d.]+)[X\*]([\d.]+)[X\*]([\d.]+)[X\*]([\d.]+)',p)
    if m: return 'HÌNH',f'H{fnum(m.group(1))}x{fnum(m.group(2))}x{fnum(m.group(3))}x{fnum(m.group(4))}'
    m=re.match(r'^(?:U|C|PFC)-?([\d.]+)[X\*]([\d.]+)[X\*]([\d.]+)',p)
    if m: return 'HÌNH',f'U{fnum(m.group(1))}x{fnum(m.group(2))}x{fnum(m.group(3))}'
    m=re.match(r'^(?:SHS|RHS)-?([\d.]+)[X\*]([\d.]+)[X\*]([\d.]+)',p)
    if m: return 'HÌNH',f'SHS{fnum(m.group(1))}x{fnum(m.group(2))}x{fnum(m.group(3))}'
    m=re.match(r'^(?:SHS|RHS)-?([\d.]+)[X\*]([\d.]+)',p)
    if m: return 'HÌNH',f'SHS{fnum(m.group(1))}x{fnum(m.group(1))}x{fnum(m.group(2))}'
    m=re.match(r'^(?:ROD|D)-?([\d.]+)',p)
    if m: return 'HÌNH',f'ROD{fnum(m.group(1))}'
    return 'KHÁC',p
