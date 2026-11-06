# Gộp file Cutting MỚI (bắt đầu từ 01/11/2026, sheet "SƠ CHẾ_THÉP TẤM" / "SƠ CHẾ_THÉP HÌNH")
# vào dữ liệu đọc từ file Cutting CŨ (sheet "List Chi Tiết" / "2-List Chi Tiết" / "LỆNH SX").
# File mới: mỗi dòng 1 chi tiết; cột "ĐÁNH GIÁ CÔNG ĐOẠN" đánh x công đoạn phải làm;
# khu "NHẬP DỮ LIỆU GIA CÔNG" mỗi công đoạn 1 khối: Công đoạn | Ngày thực hiện | Tổ | Khối lượng | QC ...
# Các dòng được đổi sang đúng bố cục cột của List Chi Tiết cũ rồi nối vào cuối, nên phần phân tích phía sau không phải sửa.
import re, pickle, datetime
import pandas as pd

SHEET_TAM, SHEET_HINH = 'SƠ CHẾ_THÉP TẤM', 'SƠ CHẾ_THÉP HÌNH'
LSX_RE = re.compile(r'^[A-Z0-9][A-Z0-9.\-_+ ]*-\d{2,}', re.I)
# tên máy ở file mới → tên máy đang dùng trong báo cáo
MAY_MAP = [(r'laser.*80|80.*laser', 'LASER 80KW (EMC)'), (r'laser.*60|60.*laser|bodoor', 'LASER 60KW (BODOOR)'),
           (r'(ống|ong).*285|285', 'CẮT ỐNG 285 (EMC)'), (r'(ống|ong).*600|d600', 'CẮT ỐNG D600'),
           (r'gas.*(sơn vũ|son vu)|(sơn vũ|son vu)', 'GAS OXY (SƠN VŨ)'), (r'gas.*emc', 'GAS OXY (EMC)'),
           (r'plasma', 'MÁY PLASMA'), (r'uốn|uon', 'MÁY UỐN ỐNG'), (r'chặt|chat', 'MÁY CHẶT TÔN 6M'), (r'chấn|chan', 'MÁY CHẤN 6m (AMADA)')]

def _txt(v):
    if v is None: return ''
    if isinstance(v, float) and v != v: return ''
    s = str(v).strip()
    return '' if s.lower() in ('nan', 'none', 'nat') else s

def _num(v):
    if v is None: return None
    if isinstance(v, (int, float)): return None if v != v else float(v)
    s = _txt(v).replace(' ', '')
    if not s: return None
    if ',' in s and '.' in s: s = s.replace('.', '').replace(',', '.')   # 1.234,5
    elif ',' in s: s = s.replace(',', '.')
    try: return float(s)
    except ValueError: return None

def _date(v):
    """→ 'YYYY-MM-DD 00:00:00' (định dạng ngày List Chi Tiết cũ sau khi đọc) hoặc None."""
    if v is None: return None
    if isinstance(v, (pd.Timestamp, datetime.datetime)): return pd.Timestamp(v).strftime('%Y-%m-%d 00:00:00')
    if isinstance(v, datetime.date): return v.strftime('%Y-%m-%d 00:00:00')
    if isinstance(v, (int, float)) and v == v and 30000 < v < 80000:    # số ngày kiểu Excel
        return (pd.Timestamp('1899-12-30') + pd.Timedelta(days=float(v))).strftime('%Y-%m-%d 00:00:00')
    s = _txt(v)
    m = re.match(r'^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})', s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3)); y = y + 2000 if y < 100 else y
        try: return datetime.date(y, mo, d).strftime('%Y-%m-%d 00:00:00')
        except ValueError: return None
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})', s)
    return f'{m.group(1)}-{m.group(2)}-{m.group(3)} 00:00:00' if m else None

def _cell(v):
    """Giống chay._cell: mọi ô thành chuỗi, ô trống = 'NaN'."""
    if v is None: return 'NaN'
    if isinstance(v, float):
        if v != v: return 'NaN'
        return str(int(v)) if v.is_integer() and abs(v) < 1e15 else ('%.10g' % v)
    if isinstance(v, int): return str(v)
    s = ' '.join(str(v).split())
    return s.replace('_', '\\_').replace('*', '\\*') if s else 'NaN'

def _stage(label):
    t = _txt(label).upper()
    if not t or 'EXTRA' in t: return None
    if 'VÁT' in t or 'VAT MEP' in t: return 'vat'
    if 'KHOAN' in t or 'ĐỘT' in t: return 'khoan'
    if 'LỐC' in t or 'CHẤN' in t or 'UỐN' in t: return 'chan'
    if 'MÀI' in t: return 'mai'
    if 'CẮT' in t or 'CƯA' in t: return 'cut'
    return None

def _may(v):
    s = _txt(v)
    if not s: return None
    lo = s.lower()
    for pat, name in MAY_MAP:
        if re.search(pat, lo): return name
    return s.upper()

def _layout(raw):
    """Tìm hàng tiêu đề và vị trí các cột theo chữ trên tiêu đề (không cố định số cột)."""
    nm = lambda x: ' '.join(_txt(x).split()).upper()
    hr = next(i for i in range(min(20, len(raw))) if any(nm(x) == 'LỆNH SẢN XUẤT' for x in raw.iloc[i]))
    H = [nm(x) for x in raw.iloc[hr]]; H1 = [nm(x) for x in raw.iloc[hr + 1]]
    def find(name, row=H, default=None):
        i = next((i for i, x in enumerate(row) if x.startswith(name)), None)
        if i is None and default is None: raise KeyError(f'không thấy cột "{name}"')
        return default if i is None else i
    c = dict(lsx=find('LỆNH SẢN XUẤT'), ma=find('MÃ CÔNG TRÌNH'), ten=find('TÊN CÔNG TRÌNH'), hm=find('HẠNG MỤC'),
             part=find('PART NAME'), mac=find('MÁC THÉP'), nest=find('NESTING NO'), may=find('MÁY CẮT'), bg=find('BÀN GIAO', H, -1))
    if c['bg'] < 0: c['bg'] = find('BÀN GIAO', H1)
    c['bh'] = find('NGÀY BAN HÀNH', default=c['lsx'] - 1); c['mail'] = find('NGÀY GỬI MAIL', default=c['lsx'] - 2)
    c['day'] = find('DÀY', default=c['part'] + 1); c['rong'] = find('RỘNG', default=c['part'] + 2)
    c['dai'] = find('DÀI', default=c['part'] + 3); c['qty'] = find('SỐ LƯỢNG', default=c['part'] + 4)
    c['kl_ct'] = find('TRỌNG LƯỢNG RIÊNG', default=c['nest'] - 2); c['kl_tong'] = find('TỔNG KHỐI LƯỢNG', default=c['nest'] - 1)
    nk = find('PHÂN LOẠI NVL'); c['g_nk'], c['g_td'] = nk, nk + 2
    dg = find('ĐÁNH GIÁ CÔNG ĐOẠN')
    c['req'] = [(i, _stage(H1[i])) for i in range(dg, c['bg']) if _stage(H1[i])]           # cột đánh x công đoạn phải làm
    c['blk'] = [(i, _stage(H[i])) for i, x in enumerate(H1) if x == 'CÔNG ĐOẠN']              # khối nhập công đoạn: (cột, công đoạn theo tiêu đề)
    return hr, c

def _rows(raw, kind, ct_map, log):
    hr, c = _layout(raw)
    out, lsx_info, skipped = [], {}, 0
    for r in raw.iloc[hr + 2:].itertuples(index=False):
        r = list(r)
        lsx = _txt(r[c['lsx']])
        if not lsx or not LSX_RE.match(lsx) or lsx.upper().startswith('LỆNH'): continue
        kl_tong = _num(r[c['kl_tong']]); qty = _num(r[c['qty']]); kl_ct = _num(r[c['kl_ct']])
        if kl_tong is None and kl_ct is not None and qty is not None: kl_tong = kl_ct * qty
        if not kl_tong: skipped += 1; continue
        req = {st for i, st in c['req'] if _txt(r[i]).lower() == 'x'}
        done = {}
        for i, st in c['blk']:
            s2 = _stage(r[i]) or st                       # ưu tiên tên công đoạn ghi ở dòng, không có thì theo tiêu đề khối
            d = _date(r[i + 1])
            if s2 and d and (s2 not in done or d < done[s2]): done[s2] = d
        ma, ten = _txt(r[c['ma']]), _txt(r[c['ten']])
        ct = ct_map.get(ma) or (f'{ma} {ten}'.strip() if ma else ten) or 'NaN'
        day, rong, dai = _num(r[c['day']]), _num(r[c['rong']]), _num(r[c['dai']])
        g = lambda x: ('%g' % x) if x is not None else '?'
        qc = (f'PL{g(day)}x{g(rong)}x{g(dai)}' if kind == 'tam' else f'{g(rong)}x{g(day)}x{g(dai)}') if (day or rong or dai) else None
        bh = _date(r[c['bh']])
        row = {'cong_trinh': ct, 'lsx': lsx, 'ten_ct': _txt(r[c['part']]) or None, 'mac': _txt(r[c['mac']]) or None, 'qc_part': qc,
               'day': day, 'rong': rong, 'dai': dai, 'qty': qty, 'kl_ct': kl_ct, 'kl_tong': kl_tong,
               'nesting': _txt(r[c['nest']]) or None, 'may': _may(r[c['may']]),
               'x_cut': 'x' if 'cut' in req else None, 'x_khoan': 'x' if 'khoan' in req else None,
               'x_chan': 'x' if 'chan' in req else None, 'x_vat': 'x' if 'vat' in req else None,
               'd_cut': done.get('cut'), 'd_khoan': done.get('khoan'), 'd_chan': done.get('chan'), 'd_vat': done.get('vat'),
               'd_bg': _date(r[c['bg']]), 'ngay_bh': bh, 'nha_may': 'VSC',
               'g_nk': 'x' if _txt(r[c['g_nk']]).lower() == 'x' else None, 'g_td': 'x' if _txt(r[c['g_td']]).lower() == 'x' else None}
        # có ngày làm mà quên đánh x → vẫn tính là công đoạn phải làm (giống cách file cũ báo E05)
        for st in ('cut', 'khoan', 'chan', 'vat'):
            if row['d_' + st] and not row['x_' + st]: row['x_' + st] = 'x'
        out.append(row)
        li = lsx_info.setdefault(lsx, dict(lsx=lsx, ngay_bh=bh, ngay_mail=_date(r[c['mail']]), cong_trinh=ct, hang_muc=_txt(r[c['hm']]) or None, sl=0.0, kl=0.0))
        li['sl'] += qty or 0; li['kl'] += kl_tong
    log(f'  {kind}: {len(out)} dòng chi tiết, {len(lsx_info)} LSX' + (f', bỏ {skipped} dòng không có khối lượng' if skipped else ''))
    return out, lsx_info

LCT = ['cong_trinh','lsx','c2','ten_ct','mac','qc_part','day','rong','dai','qty','kl_ct','kl_tong','qc_vt','nesting','may','ca','c16',
       'x_cut','x_khoan','x_chan','x_vat','c21','d_cut','d_khoan','d_chan','d_vat','d_bg','c27','lan_bh','ngay_bh','note','nha_may','c32','chk_date','chk_kl','c35','g_nk','g_td','nhom']
LSXC = ['stt','lsx','ngay_bh','ngay_mail','cong_trinh','hang_muc','dv_cat','dv_gc','dv_nhan','nha_may','tong_sl','tong_kl','tinh_trang','phieu_xuat','phieu_thu','c15','st_cut','st_khoan','st_chan','st_vat','st_bg','c21','kl_cut','ht_cut','kl_khoan','ht_khoan','kl_chan','ht_chan','kl_vat','ht_vat','kl_bg','ht_bg','may','hoan_thanh']

def _append(df, rows, cols):
    n = df.shape[1]
    data = [[_cell(r.get(k)) for k in cols][:n] + ['NaN'] * max(0, n - len(cols)) for r in rows]
    add = pd.DataFrame(data, columns=df.columns)
    out = pd.concat([df, add], ignore_index=True); out.attrs = dict(df.attrs)
    return out

def _find(d, name):
    return next((k for k in d if k.strip().lower() == name.lower()), None)

def gop(sheets_pkl, xlsx_moi, log=print):
    """Đọc file Cutting mới và nối vào sheets.pkl của file cũ. Trả về số dòng đã nối."""
    d = pickle.load(open(sheets_pkl, 'rb'))
    k1, k2, kl = _find(d, 'List Chi Tiết'), _find(d, '2-List Chi Tiết'), _find(d, 'LỆNH SX')
    # tên công trình đang dùng ở file cũ, theo mã (vd "10626-051" → "10626-051 DG SÂN VẬN ĐỘNG HÙNG VƯƠNG")
    ct_map = {}
    for k in (k1, k2):
        for v in d[k].iloc[3:, 0]:
            s = _txt(v).replace('\\_', '_')
            m = re.match(r'^([\w.\-]+)\s', s)
            if m and m.group(1) not in ct_map and s != 'NaN': ct_map[m.group(1)] = s
    xl = pd.ExcelFile(xlsx_moi)
    names = {n.strip().upper(): n for n in xl.sheet_names}
    allrows, info, tong = {'tam': [], 'hinh': []}, {}, 0
    for kind, sname in (('tam', SHEET_TAM), ('hinh', SHEET_HINH)):
        sn = names.get(sname.upper())
        if not sn: log(f'  file mới không có sheet "{sname}"'); continue
        raw = pd.read_excel(xl, sheet_name=sn, header=None, dtype=object)
        rows, li = _rows(raw, kind, ct_map, log)
        allrows[kind] = rows; tong += len(rows)
        for l, x in li.items():
            if l in info: info[l]['sl'] += x['sl']; info[l]['kl'] += x['kl']
            else: info[l] = x
    # LSX chưa có trong sheet LỆNH SX của file cũ → thêm (ngày ban hành, công trình, hạng mục lấy từ file mới)
    co = {_txt(v) for v in d[kl].iloc[2:, 1]}
    # Ngày bị Google Sheets hiểu ngược tháng/ngày (gõ 04/11 thành 11/04 = 11 tháng 4): chỉ sửa LSX MỚI (chưa có ở file cũ),
    # khi ngày quá cũ so với hôm nay mà đảo lại thì rơi vào 60 ngày gần đây → đảo ngày/tháng.
    nay = datetime.date.today()
    def dao(sd):
        if not sd: return sd
        x = datetime.date.fromisoformat(sd[:10])
        if x.day <= 12 and x < nay - datetime.timedelta(days=60):
            try: y = datetime.date(x.year, x.day, x.month)
            except ValueError: return sd
            if nay - datetime.timedelta(days=60) <= y <= nay + datetime.timedelta(days=1): return y.strftime('%Y-%m-%d 00:00:00')
        return sd
    sua = []
    for l, x in info.items():
        if l in co: continue
        for k in ('ngay_bh', 'ngay_mail'):
            v = dao(x[k])
            if v != x[k]:
                if k == 'ngay_bh': sua.append(f'{l} {x[k][:10]}→{v[:10]}')
                x[k] = v
    if sua:
        fx = {l: info[l]['ngay_bh'] for l in info if l not in co}
        for kind in allrows:
            for r in allrows[kind]:
                if r['lsx'] in fx: r['ngay_bh'] = fx[r['lsx']]
        log('  sửa ngày ban hành bị đảo ngày/tháng: ' + ', '.join(sua))
    if allrows['tam']: d[k1] = _append(d[k1], allrows['tam'], LCT)
    if allrows['hinh']: d[k2] = _append(d[k2], allrows['hinh'], LCT)
    moi = [dict(x, tong_sl=x['sl'], tong_kl=x['kl']) for l, x in info.items() if l not in co]
    if moi: d[kl] = _append(d[kl], moi, LSXC)
    pickle.dump(d, open(sheets_pkl, 'wb'))
    log(f'Gộp file Cutting mới: {tong} dòng chi tiết, {len(info)} LSX ({len(moi)} LSX chưa có trong LỆNH SX cũ)')
    return tong
