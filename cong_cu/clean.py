# Chuẩn hóa dữ liệu file Cutting trước khi tính báo cáo.
# Mỗi quy tắc được ghi vào FIX để báo cáo liệt kê "đã chuẩn hóa gì, bao nhiêu dòng".
import os, re
import numpy as np, pandas as pd

NGAY_SO_LIEU = pd.Timestamp(os.environ.get('NGAY_BAO_CAO', '2026-09-28'))
_DATE_COLS = [('d_cut', 'cắt'), ('d_khoan', 'khoan'), ('d_chan', 'chấn'), ('d_vat', 'vát mép'), ('d_bg', 'bàn giao')]

def _iso(s):
    return pd.to_datetime(s.where(s.astype(str).str.match(r'^\d{4}-\d{2}-\d{2}')).astype(str).str[:10], errors='coerce')

def lam_sach(V, today=None):
    today = today if today is not None else NGAY_SO_LIEU
    V = V.copy(); FIX = []
    def log(ma, quy_tac, mask, nhom='Đã chuẩn hóa', col=None, out=None):
        mask = pd.Series(mask, index=V.index).fillna(False).astype(bool)
        if mask.any(): FIX.append(dict(ma=ma, quy_tac=quy_tac, nhom=nhom, mask=mask.values.copy(), col=col, out=out or col))

    # 0. ô trống thật sự ('' hoặc khoảng trắng) coi là rỗng
    for c, _ in _DATE_COLS + [('x_cut', ''), ('x_khoan', ''), ('x_chan', ''), ('x_vat', ''), ('g_nk', ''), ('g_td', ''), ('may', '')]:
        V.loc[V[c].astype(str).str.strip().isin(['', 'NaN', 'nan']), c] = np.nan
    for c, _ in _DATE_COLS: V[c + '_dt'] = _iso(V[c])
    V0 = V.copy()

    # 1. ô ngày khoan ghi "K" (ký hiệu, không phải ngày) → coi là chưa có ngày khoan
    k = V.d_khoan.astype(str).str.strip().str.upper().isin(['K', 'KHOAN'])
    V.loc[k, 'd_khoan'] = np.nan
    log('R01', 'Ô NGÀY KHOAN ghi "K" (ký hiệu công việc, không phải ngày) → coi là chưa khoan', k, col='d_khoan')

    # 2. ô ngày có ghi chú kèm ngày ("24/6/2026 - SX3", "17/6 - Tuấn") → lấy phần ngày
    for c, lab in _DATE_COLS:
        raw = V[c]; txt = raw.notna() & V[c + '_dt'].isna()
        m = raw.where(txt).astype(str).str.extract(r'(\d{1,2})\s*[/.\-]\s*(\d{1,2})(?:\s*[/.\-]\s*(\d{2,4}))?')
        y = pd.to_numeric(m[2], errors='coerce'); y = y.where(y.isna() | (y >= 100), y + 2000).fillna(today.year)
        dt = pd.to_datetime(dict(year=y, month=pd.to_numeric(m[1], errors='coerce'), day=pd.to_numeric(m[0], errors='coerce')), errors='coerce')
        # thiếu năm mà ra ngày tương lai → năm trước
        dt = dt.where(~(dt > today) | m[2].notna(), dt - pd.DateOffset(years=1))
        got = txt & dt.notna()
        V.loc[got, c + '_dt'] = dt[got]
        log('R02', f'Ô NGÀY {lab.upper()} có ghi chú kèm ngày (vd "24/6/2026 - SX3") → tách lấy ngày', got, col=c, out=c + '_dt')

    # 3. ngày ở tương lai → thử đảo ngày/tháng (09/12 gõ nhầm thành 12/09)
    for c, lab in _DATE_COLS:
        d = V[c + '_dt']; fut = d > today
        sw = pd.to_datetime(dict(year=d.dt.year, month=d.dt.day, day=d.dt.month), errors='coerce')
        ok = fut & sw.notna() & (sw <= today)
        V.loc[ok, c + '_dt'] = sw[ok]
        log('R03', f'NGÀY {lab.upper()} ở tương lai → đảo ngày/tháng (vd 2026-12-09 → 2026-09-12)', ok, col=c, out=c + '_dt')
        bad = fut & ~ok
        V.loc[bad, c + '_dt'] = pd.NaT
        log('R03b', f'NGÀY {lab.upper()} ở tương lai, không đảo được → vẫn tính đã làm, không đưa vào số liệu tháng', bad, col=c)

    # 4. đã bàn giao nhưng chưa có ngày cắt → đã cắt, ngày cắt = ngày bàn giao
    e1 = V.d_bg.notna() & V.d_cut.isna()
    V.loc[e1, 'd_cut'] = '(suy từ bàn giao)'; V.loc[e1, 'd_cut_dt'] = V.loc[e1, 'd_bg_dt']
    V.loc[e1 & V.x_cut.isna(), 'x_cut'] = 'x'
    log('R04', 'Có BÀN GIAO nhưng trống NGÀY CẮT → tính đã cắt, ngày cắt = ngày bàn giao', e1, col='d_cut', out='d_cut_dt')

    # 5. bàn giao ghi tên người/tổ nhận (không có ngày) → đã bàn giao, ngày = ngày cắt
    bt = V.d_bg.notna() & V.d_bg_dt.isna()
    V.loc[bt, 'd_bg_dt'] = V.loc[bt, 'd_cut_dt']
    log('R05', 'Ô BÀN GIAO ghi tên người/tổ nhận (không có ngày) → tính đã bàn giao, ngày = ngày cắt', bt, col='d_bg', out='d_bg_dt')

    # 6. ngày bàn giao sớm hơn ngày cắt → ngày bàn giao = ngày cắt
    e2 = V.d_bg_dt.notna() & V.d_cut_dt.notna() & (V.d_bg_dt < V.d_cut_dt)
    V.loc[e2, 'd_bg_dt'] = V.loc[e2, 'd_cut_dt']
    log('R06', 'NGÀY BÀN GIAO sớm hơn NGÀY CẮT → lấy ngày bàn giao = ngày cắt', e2, col='d_bg_dt')

    # 7. có ngày gia công nhưng thiếu dấu "x" công việc → thêm dấu x
    for st, lab in [('khoan', 'KHOAN'), ('chan', 'CHẤN-LỐC'), ('vat', 'VÁT MÉP')]:
        m = V['d_' + st].notna() & V['x_' + st].isna()
        V.loc[m, 'x_' + st] = 'x'
        log('R07', f'Có NGÀY {lab} nhưng thiếu dấu "x" công việc → tính là có công đoạn {lab.lower()}', m, col='x_' + st)

    # 8. phân loại thép trống hoặc tick cả hai → Nguyên khổ
    both = V.g_nk.notna() & V.g_td.notna(); V.loc[both, 'g_td'] = np.nan
    log('R08', 'Tick cả NGUYÊN KHỔ và TẬN DỤNG → tính Nguyên khổ', both, col='g_td')
    none = V.g_nk.isna() & V.g_td.isna(); V.loc[none, 'g_nk'] = 'x'
    log('R08', 'Trống phân loại thép → tính Nguyên khổ (an toàn khi cân đối kho)', none, col='g_nk')

    # 9. đã cắt nhưng trống máy
    nm = V.d_cut.notna() & V.may.isna(); V.loc[nm, 'may'] = '(chưa ghi máy)'
    log('R09', 'Đã cắt nhưng trống MÁY CẮT → xếp vào nhóm "(chưa ghi máy)"', nm, col='may')

    # 10. quy ước khối lượng (không sửa, chỉ ghi nhận)
    r = (V.kl_ct * V.qty - V.kl_tong).abs() / V.kl_tong.replace(0, np.nan)
    log('R10', 'KL TỔNG khác SL × KL chi tiết → dùng KL TỔNG của file (khớp LỆNH SX)', r > 0.01, 'Quy ước', col='kl_tong')
    dup = V.duplicated(['lsx', 'ten_ct', 'nesting', 'qty', 'kl_tong', 'qc_vt', 'may'], keep=False)
    log('R11', 'Dòng giống hệt nhau (cùng LSX, chi tiết, nesting, SL, KL, máy) → giữ nguyên như file, cần phòng Cutting xác nhận', dup, 'Cần xác nhận')
    def fmt(v):
        if isinstance(v, np.datetime64): v = pd.Timestamp(v)
        if v is None or (isinstance(v, float) and v != v) or v is pd.NaT: return ''
        if isinstance(v, pd.Timestamp): return v.strftime('%d/%m/%Y')
        v = str(v).replace('\\_', '_')
        return re.sub(r'^(\d{4})-(\d{2})-(\d{2}) 00:00:00$', r'\3/\2/\1', v)
    for f in FIX:
        m = f['mask']
        f['old'] = [fmt(x) for x in V0[f['col']].values[m]] if f['col'] else [''] * int(m.sum())
        f['new'] = [fmt(x) for x in V[f['out']].values[m]] if f['out'] else [''] * int(m.sum())
    return V, FIX
