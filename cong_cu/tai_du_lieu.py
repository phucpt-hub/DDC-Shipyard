# Tải bản mới nhất của file nguồn từ link (Google Sheets / Google Drive / link tải trực tiếp)
# vào du_lieu/<thư mục>, đặt tên kèm ngày hôm nay để chay.py chọn làm file mới nhất.
# Link lấy từ biến môi trường (GitHub Secrets), không ghi trong mã nguồn:
#   NGUON_CUTTING -> du_lieu/1_CUTTING
#   NGUON_KHO     -> du_lieu/2_TON_KHO
#   NGUON_NESTING -> du_lieu/3_NESTING   (nhiều link, cách nhau bằng dấu xuống dòng hoặc dấu ;)
#   NGUON_VIEC    -> du_lieu/4_CAU_HINH  (Google Sheet danh sách việc cần xử lý / cần quyết định)
# Link Apps Script (script.google.com/.../exec) cần đã gắn file apps_script_xuat_du_lieu.gs;
# khóa đặt trong biến <TÊN>_KHOA (vd NGUON_KHO_KHOA) hoặc ghi sẵn trong link (?khoa=...).
import os, re, sys, time, base64, datetime, urllib.request, urllib.parse

CC = os.path.dirname(os.path.abspath(__file__))
DL = os.path.join(os.path.dirname(CC), 'du_lieu')
NGUON = [('NGUON_CUTTING', '1_CUTTING', 'CUTTING_LINK'),
         ('NGUON_KHO', '2_TON_KHO', 'TON_KHO_LINK'),
         ('NGUON_NESTING', '3_NESTING', 'NESTING_LINK'),
         ('NGUON_VIEC', '4_CAU_HINH', 'VIEC_LINK')]

def log(*a): print('>>', *a, flush=True)

def export_url(link, khoa=''):
    link = link.strip()
    m = re.search(r'docs\.google\.com/spreadsheets/d/([\w-]+)', link)
    if m: return f'https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=xlsx', 'Google Sheets'
    m = re.search(r'drive\.google\.com/(?:file/d/|open\?id=|uc\?id=)([\w-]+)', link)
    if m: return f'https://drive.google.com/uc?export=download&id={m.group(1)}', 'Google Drive'
    if 'script.google.com' in link and '/exec' in link:
        q = dict(urllib.parse.parse_qsl(urllib.parse.urlsplit(link).query))
        q['xuat'] = 'xlsx'
        if khoa and 'khoa' not in q: q['khoa'] = khoa
        return link.split('?')[0] + '?' + urllib.parse.urlencode(q), 'Apps Script'
    if re.fullmatch(r'[\w-]{30,}', link):          # chỉ dán mã bảng tính
        return f'https://docs.google.com/spreadsheets/d/{link}/export?format=xlsx', 'Google Sheets'
    return link, 'link trực tiếp'

def cac_url(url):
    """Google Sheets có vài đường xuất Excel khác nhau – đường này kẹt thì thử đường khác."""
    m = re.search(r'docs\.google\.com/spreadsheets/d/([\w-]+)/export\?format=xlsx', url)
    if not m: return [url]
    i = m.group(1)
    return [url,
            f'https://docs.google.com/spreadsheets/export?id={i}&exportFormat=xlsx',
            f'https://docs.google.com/feeds/download/spreadsheets/Export?key={i}&exportFormat=xlsx']

TEN = {}   # tên file gốc máy chủ trả về (Google Sheets: tên bảng tính)
def ten_file(cd):
    m = re.search(r"filename\*=UTF-8''([^;]+)", cd, re.I)
    if m: return urllib.parse.unquote(m.group(1)).strip('"')
    m = re.search(r'filename="?([^";]+)"?', cd, re.I)
    if not m: return ''
    t = m.group(1)
    try: t = t.encode('latin-1').decode('utf-8')
    except Exception: pass
    return t

def ma_rut_gon(url):
    m = re.search(r'/d/([\w-]{10,})', url) or re.search(r'[?&](?:id|key)=([\w-]{10,})', url)
    return (m.group(1)[:4] + '…' + m.group(1)[-4:]) if m else ''

def tai_mot_lan(url, dich, cho_toi_da):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (bao-cao-chuoi-thep)'})
    t0 = time.time(); n = 0; tmp = dich + '.part'
    with urllib.request.urlopen(req, timeout=cho_toi_da) as r, open(tmp, 'wb') as fh:
        log(f'  máy chủ bắt đầu trả file sau {time.time() - t0:.0f} giây')
        TEN[dich] = ten_file(r.headers.get('Content-Disposition', ''))
        while True:
            b = r.read(1 << 20)
            if not b: break
            fh.write(b); n += len(b)
    data = open(tmp, 'rb').read(); os.remove(tmp)
    return data

def tai(url, dich, lan=6, cho_toi_da=420, tong_toi_da=36 * 60):
    """Tải file; mỗi lần chờ tối đa cho_toi_da giây không thấy dữ liệu thì thử lại (đổi đường tải nếu có);
    tổng thời gian không quá tong_toi_da giây để lần chạy không bị GitHub cắt ngang."""
    urls = cac_url(url); t_bd = time.time(); data = None
    for i in range(lan):
        u = urls[i % len(urls)]
        try:
            data = tai_mot_lan(u, dich, cho_toi_da)
            break
        except Exception as e:
            con = tong_toi_da - (time.time() - t_bd)
            if i == lan - 1 or con < cho_toi_da + 60: raise
            cho = 15 * (i + 1)
            log(f'  lỗi tải lần {i + 1} ({e}), thử lại sau {cho} giây' + (' bằng đường tải khác' if len(urls) > 1 else ''))
            time.sleep(cho)
    if data[:2] != b'PK' and data[:4] == b'UEsD':   # Apps Script trả file Excel dạng base64
        data = base64.b64decode(data)
    if data[:7] == b'LOI: sa':
        raise RuntimeError('Apps Script báo sai khóa (kiểm tra secret ..._KHOA và Script property KHOA_XUAT)')
    if data[:2] != b'PK':   # file xlsx là file zip, bắt đầu bằng "PK"
        head = data[:300].decode('utf-8', 'ignore').lower()
        why = 'link đòi đăng nhập – cần chia sẻ "Bất kỳ ai có đường liên kết đều xem được"' if ('sign in' in head or 'accounts.google' in head or '<html' in head) else 'nội dung tải về không phải file Excel'
        raise RuntimeError(why)
    open(dich, 'wb').write(data)
    return len(data)

def file_tay_hom_nay(thu_muc, ten_link):
    """File Cutting tải tay lên du_lieu/<thu_muc> có ngày hôm nay trong tên (vd '... SC 02.11.2026.xlsx')."""
    nay = (datetime.datetime.utcnow() + datetime.timedelta(hours=7)).date()
    for f in sorted(os.listdir(os.path.join(DL, thu_muc))):
        if f.startswith(ten_link) or f.startswith('~$') or not f.lower().endswith(('.xlsx', '.xlsm', '.xlsb', '.xls')): continue
        for m in re.finditer(r'(?<!\d)(\d{1,2})[._-](\d{1,2})(?:[._-](\d{2,4}))?(?!\d)', f):
            y = int(m.group(3)) if m.group(3) else nay.year
            y = y + 2000 if y < 100 else y
            try:
                if datetime.date(y, int(m.group(2)), int(m.group(1))) >= nay: return f
            except ValueError: pass
    return None

NGUON_DA_TAI = {}

def main():
    hom_nay = (datetime.datetime.utcnow() + datetime.timedelta(hours=7)).strftime('%d.%m.%Y')
    co_link = False
    loi_cutting = False
    for bien, thu_muc, ten in NGUON:
        links = [x for x in re.split(r'[\n;]+', os.environ.get(bien, '')) if x.strip()]
        for i, link in enumerate(links):
            co_link = True
            url, loai = export_url(link, os.environ.get(bien + '_KHOA', ''))
            os.makedirs(os.path.join(DL, thu_muc), exist_ok=True)
            dich = os.path.join(DL, thu_muc, f'{ten}{"_" + str(i + 1) if len(links) > 1 else ""} {hom_nay}.xlsx')
            t0 = time.time()
            try:
                n = tai(url, dich)
                ten_goc = re.sub(r'\.xlsx?$', '', TEN.get(dich, ''), flags=re.I)
                log(f'Tải {thu_muc} từ {loai}: {n / 1e6:.1f} MB trong {time.time() - t0:.0f} giây' + (f' – file: {ten_goc}' if ten_goc else '') + (f' (mã {ma_rut_gon(url)})' if ma_rut_gon(url) else ''))
                NGUON_DA_TAI[thu_muc] = dict(ten=ten_goc, ma=ma_rut_gon(url), loai=loai, mb=round(n / 1e6, 1),
                                             luc=(datetime.datetime.utcnow() + datetime.timedelta(hours=7)).strftime('%d/%m/%Y %H:%M'))
            except Exception as e:
                # không dừng cả báo cáo: dùng file đã có sẵn trong thư mục (nếu có)
                log(f'KHÔNG tải được {thu_muc} ({loai}): {e}. Dùng file đang có trong du_lieu/{thu_muc}.')
                if os.path.exists(dich): os.remove(dich)
                if bien == 'NGUON_CUTTING': loi_cutting = True
    try:
        os.makedirs(os.path.join(CC, '_tmp'), exist_ok=True)
        import json
        json.dump(NGUON_DA_TAI, open(os.path.join(CC, '_tmp', 'nguon.json'), 'w', encoding='utf-8'), ensure_ascii=False)
    except Exception as e: log('Không ghi được thông tin nguồn:', e)
    if not co_link: log('Chưa khai báo link nguồn (NGUON_CUTTING / NGUON_KHO / NGUON_NESTING) – dùng file trong du_lieu/.')
    if loi_cutting:
        tay = file_tay_hom_nay('1_CUTTING', 'CUTTING_LINK')
        if tay:
            log(f'Link Cutting lỗi nhưng đã có file tải tay hôm nay: {tay} – dùng file này để cập nhật báo cáo.')
            return
        # Không dựng báo cáo từ file Cutting cũ trong repo (sẽ làm báo cáo lùi ngày, mất chi tiết LSX).
        # Dừng lần chạy này: website giữ nguyên báo cáo gần nhất, lần chạy sau sẽ thử tải lại.
        log('Cách cập nhật ngay: tải file Sơ chế về (Tệp → Tải xuống → .xlsx), đặt tên có ngày hôm nay, vd "DDC SHIPYARD - KHGC - SC dd.mm.yyyy.xlsx", rồi tải lên du_lieu/1_CUTTING.')
        log('DỪNG: không tải được file Cutting từ link – giữ nguyên báo cáo đang có, không cập nhật lần này.')
        sys.exit(1)

if __name__ == '__main__':
    main()
