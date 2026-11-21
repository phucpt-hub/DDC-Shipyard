# Bản lưu file Sơ chế CŨ (link 16FQ…, dữ liệu đến 31/10/2026).
# File cũ có thể bị xóa bớt LSX/dòng, ngừng cập nhật hoặc mất link → mỗi lần chạy:
#   1. lấy dữ liệu file cũ vừa tải + những LSX chỉ còn trong bản lưu (đã bị xóa khỏi file) → ghi lại bản lưu
#   2. nối các LSX chỉ còn trong bản lưu vào dữ liệu báo cáo (bỏ LSX đã có ở file Cutting mới → không cộng trùng)
#   3. không tải được file cũ → dùng nguyên bản lưu thay cho file cũ
# Bản lưu: du_lieu/1_CUTTING_LUU/so_che_cu.pkl.gz (workflow commit lại khi có thay đổi)
import os, gzip, pickle, hashlib, datetime
import pandas as pd

# (tên sheet, số dòng tiêu đề phía trên dữ liệu, cột số LSX)
SHEETS = [('List Chi Tiết', 3, 1), ('2-List Chi Tiết', 3, 1), ('LỆNH SX', 2, 1)]

def _find(d, name):
    return next((k for k in d if k.strip().lower() == name.lower()), None)

def _key(v):
    s = ' '.join(str(v).replace('\\_', '_').split())
    return '' if s.lower() in ('', 'nan', 'none', 'nat') else s.upper()

def _lsx(df, h, col):
    return df.iloc[h:, col].map(_key)

def doc(path):
    """Đọc bản lưu → dict {'ngay': 'YYYY-MM-DD', 'sheets': {...}} hoặc None."""
    if not os.path.exists(path): return None
    try:
        with gzip.open(path, 'rb') as f: return pickle.load(f)
    except Exception as e:
        print('>> Không đọc được bản lưu Sơ chế cũ:', e); return None

def dung_ban_luu(path, sheets_pkl, log=print):
    """Link file cũ hỏng/không còn: dùng nguyên bản lưu làm dữ liệu file cũ. Trả về ngày của bản lưu hoặc None."""
    L = doc(path)
    if not L: return None
    pickle.dump(L['sheets'], open(sheets_pkl, 'wb'))
    log('Không có file Sơ chế cũ mới tải → dùng BẢN LƯU Sơ chế cũ ngày', L['ngay'])
    return L['ngay']

def chup(sheets_pkl):
    """Giữ bản sao dữ liệu file cũ (trước khi nối file Cutting mới)."""
    return pickle.load(open(sheets_pkl, 'rb'))

def gop(sheets_pkl, cu, path, ngay, log=print, ghi=True):
    """cu = dữ liệu file cũ vừa đọc (chup()). Cập nhật bản lưu và nối LSX chỉ còn trong bản lưu vào sheets_pkl.
    Trả về danh sách LSX đã giữ lại từ bản lưu."""
    L = doc(path)
    snap = {k: v for k, v in cu.items()}
    if L:
        for name, h, col in SHEETS:
            k, o = _find(cu, name), L['sheets'].get(_find(L['sheets'], name) or '')
            if k is None or o is None: continue
            if o.shape[1] != cu[k].shape[1]:
                log(f'  bản lưu: sheet "{name}" đổi số cột ({o.shape[1]} → {cu[k].shape[1]}) – không giữ dòng cũ của sheet này'); continue
            co = set(_lsx(cu[k], h, col))
            ok = _lsx(o, h, col)
            add = o.iloc[h:][(ok != '') & ~ok.isin(co)]
            if len(add):
                add = add.copy(); add.columns = cu[k].columns
                s = pd.concat([cu[k], add], ignore_index=True); s.attrs = dict(cu[k].attrs); snap[k] = s
    # nối vào dữ liệu báo cáo: chỉ LSX chưa có (kể cả LSX đã chuyển sang file Cutting mới)
    d = pickle.load(open(sheets_pkl, 'rb')); giu = set()
    for name, h, col in SHEETS:
        k, kd = _find(snap, name), _find(d, name)
        if k is None or kd is None or snap[k] is cu.get(k): continue
        co = set(_lsx(d[kd], h, col))
        extra = snap[k].iloc[len(cu[k]):]
        ek = extra.iloc[:, col].map(_key)
        add = extra[~ek.isin(co)]
        if len(add):
            s = pd.concat([d[kd], add], ignore_index=True); s.attrs = dict(d[kd].attrs); d[kd] = s
            giu |= set(ek[add.index])
    pickle.dump(d, open(sheets_pkl, 'wb'))
    if giu: log(f'Bản lưu Sơ chế cũ: giữ lại {len(giu)} LSX đã bị xóa khỏi file cũ: ' + ', '.join(sorted(giu)[:30]) + (' …' if len(giu) > 30 else ''))
    # ghi bản lưu khi nội dung thay đổi
    if ghi:
        moi = pickle.dumps({k: snap[k] for k in snap}, protocol=4)
        cu_h = hashlib.sha1(pickle.dumps(L['sheets'], protocol=4)).hexdigest() if L else ''
        if hashlib.sha1(moi).hexdigest() != cu_h:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with gzip.open(path, 'wb', compresslevel=9) as f: pickle.dump({'ngay': str(ngay), 'sheets': snap}, f, protocol=4)
            log(f'Đã cập nhật bản lưu Sơ chế cũ ({os.path.getsize(path) / 1e6:.1f} MB, ngày {ngay})')
    return sorted(giu)
