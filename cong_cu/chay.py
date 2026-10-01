# Chạy toàn bộ: đọc Excel trong du_lieu/ -> phân tích -> index.html + lich_su/<ngày>/
import os, sys, re, glob, json, pickle, subprocess, datetime, shutil, runpy
import pandas as pd

CC = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(CC)
DL = os.path.join(ROOT, 'du_lieu')
TMP = os.path.join(CC, '_tmp'); os.makedirs(TMP, exist_ok=True)
EXT = ('.xlsx', '.xlsm', '.xls', '.xlsb')

def log(*a): print('>>', *a, flush=True)

def commit_time(p):
    try:
        t = subprocess.run(['git', '-C', ROOT, 'log', '-1', '--format=%ct', '--', p], capture_output=True, text=True).stdout.strip()
        return int(t) if t else os.path.getmtime(p)
    except Exception:
        return os.path.getmtime(p)

def files_in(folder):
    fs = [f for f in glob.glob(os.path.join(DL, folder, '*')) if f.lower().endswith(EXT) and not os.path.basename(f).startswith('~$')]
    # cùng tên nhưng khác đuôi (.xlsx và .xlsb) → ưu tiên .xlsx
    best = {}
    for f in fs:
        stem, ext = os.path.splitext(os.path.basename(f)); rank = EXT.index(ext.lower())
        if stem not in best or rank < best[stem][0]: best[stem] = (rank, f)
    skipped = sorted(set(fs) - {v[1] for v in best.values()})
    for f in skipped: log('Bỏ qua (trùng tên, đã có bản .xlsx):', os.path.basename(f))
    return sorted([v[1] for v in best.values()], key=commit_time)

def _cell(v):
    if v is None: return 'NaN'
    if isinstance(v, float):
        if v != v: return 'NaN'
        return str(int(v)) if v.is_integer() and abs(v) < 1e15 else ('%.10g' % v)
    if isinstance(v, bool): return str(v)
    if isinstance(v, int): return str(v)
    if isinstance(v, (pd.Timestamp, datetime.datetime)): return pd.Timestamp(v).strftime('%Y-%m-%d %H:%M:%S')
    if isinstance(v, datetime.date): return v.strftime('%Y-%m-%d 00:00:00')
    if isinstance(v, datetime.time): return v.strftime('%H:%M:%S')
    s = ' '.join(str(v).split())            # gộp xuống dòng như bản markdown
    return s.replace('_', '\\_').replace('*', '\\*') if s else 'NaN'

def to_sheets(xlsx, out_pkl):
    """Excel -> bảng chuỗi theo từng sheet, cùng định dạng với bản .md đã dùng để phân tích."""
    raw = pd.read_excel(xlsx, sheet_name=None, header=0, dtype=object, engine='pyxlsb' if xlsx.lower().endswith('.xlsb') else None)
    out = {}
    for k, df in raw.items():
        hdr = [_cell(c) if not str(c).startswith('Unnamed') else str(c) for c in df.columns]
        n = len(hdr)
        data = [[_cell(v) for v in row] for row in df.itertuples(index=False, name=None)]
        d2 = pd.DataFrame(data, columns=[f'c{i}' for i in range(n)]); d2.attrs['hdr'] = hdr; out[k.strip() if k.strip() != k and k.strip() in ('TỔNG HỢP NHẬP (TỒN + XUẤT)',) else k] = d2
    pickle.dump(out, open(out_pkl, 'wb'))
    log('Đọc', os.path.basename(xlsx), '→', len(out), 'sheet:', ', '.join(list(out)[:8]))
    return out

def main():
    cut = files_in('1_CUTTING'); kho = files_in('2_TON_KHO'); nes = files_in('3_NESTING'); cfg = files_in('4_CAU_HINH')
    if not cut: sys.exit('THIẾU file Cutting trong du_lieu/1_CUTTING')
    if not kho: sys.exit('THIẾU file tồn kho trong du_lieu/2_TON_KHO')
    f_cut, f_kho = cut[-1], kho[-1]
    log('File Cutting:', os.path.basename(f_cut)); log('File tồn kho:', os.path.basename(f_kho))
    # Ngày báo cáo: lấy từ tên file tồn kho (vd "TỒN KHO 28.09"), không có thì lấy hôm nay (giờ VN)
    now = datetime.datetime.utcnow() + datetime.timedelta(hours=7)
    m = re.search(r'(\d{1,2})[._-](\d{1,2})(?:[._-](\d{2,4}))?', os.path.basename(f_kho))
    if m:
        y = int(m.group(3)) if m.group(3) else now.year
        y = y + 2000 if y < 100 else y
        try: ngay = datetime.date(y, int(m.group(2)), int(m.group(1)))
        except ValueError: ngay = now.date()
    else:
        ngay = now.date()
    log('Ngày báo cáo:', ngay)
    to_sheets(f_cut, os.path.join(CC, 'sheets.pkl'))
    to_sheets(f_kho, os.path.join(CC, 'kho', 'sheets.pkl'))
    nfiles = []
    NES_OK = []
    for i, f in enumerate(nes):
        p = os.path.join(TMP, f'nes{i}.pkl')
        try:
            sh = to_sheets(f, p)
        except Exception as e:
            log('KHÔNG đọc được file nesting', os.path.basename(f), '-', e); continue
        if not any(k.strip().lower() == 'phieu vat tu' for k in sh):
            log('Bỏ qua', os.path.basename(f), '- không có sheet "Phieu vat tu" (không phải file nesting)'); continue
        nfiles.append([p, os.path.splitext(os.path.basename(f))[0]])
    os.environ.update(NGAY_BAO_CAO=str(ngay), NESTING_FILES=json.dumps(nfiles, ensure_ascii=False),
                      FILE_CUTTING=os.path.splitext(os.path.basename(f_cut))[0], FILE_KHO=os.path.splitext(os.path.basename(f_kho))[0],
                      OUT_JSON=os.path.join(TMP, 'data.json'), OUT_LOG=os.path.join(TMP, 'clean_log.pkl'))
    if cfg: os.environ['FILE_CAU_HINH'] = cfg[-1]
    os.chdir(CC)
    runpy.run_path('data.py', run_name='__main__')
    data = open(os.path.join(TMP, 'data.json'), encoding='utf-8').read().replace('</', '<\\/')
    html = open('template.html', encoding='utf-8').read().replace('/*DATA*/null', data)
    open(os.path.join(ROOT, 'index.html'), 'w', encoding='utf-8').write(html)
    hd = os.path.join(ROOT, 'lich_su', str(ngay)); os.makedirs(hd, exist_ok=True)
    open(os.path.join(hd, 'index.html'), 'w', encoding='utf-8').write(html.replace('href="lich_su/"', 'href="../"'))
    # trang danh sách báo cáo cũ
    days = sorted([d for d in os.listdir(os.path.join(ROOT, 'lich_su')) if re.match(r'\d{4}-\d{2}-\d{2}$', d)], reverse=True)
    items = ''.join(f'<li><a href="{d}/">Báo cáo ngày {d[8:10]}/{d[5:7]}/{d[:4]}</a></li>' for d in days)
    open(os.path.join(ROOT, 'lich_su', 'index.html'), 'w', encoding='utf-8').write(
        '<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>Báo cáo cũ</title>'
        '<style>body{font:15px/1.6 Arial,sans-serif;max-width:640px;margin:0 auto;padding:24px 16px;color:#18232d;background:#f3f5f6}a{color:#1d5f86}</style>'
        f'<h1>Lịch sử báo cáo chuỗi thép</h1><p><a href="../">← Báo cáo mới nhất</a></p><ul>{items}</ul>')
    # công cụ làm sạch: nhật ký Excel
    sys.path.insert(0, CC); import nhat_ky
    n = nhat_ky.ghi(pickle.load(open(os.path.join(TMP, 'clean_log.pkl'), 'rb')), os.path.join(ROOT, 'nhat_ky_lam_sach.xlsx'), ngay.strftime('%d/%m/%Y'))
    shutil.copy(os.path.join(ROOT, 'nhat_ky_lam_sach.xlsx'), os.path.join(hd, 'nhat_ky_lam_sach.xlsx'))
    log('Nhật ký làm sạch:', n, 'ô')
    log('XONG: index.html + lich_su/' + str(ngay))

if __name__ == '__main__':
    main()
