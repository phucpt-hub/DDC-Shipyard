# Tải bản mới nhất của file nguồn từ link (Google Sheets / Google Drive / link tải trực tiếp)
# vào du_lieu/<thư mục>, đặt tên kèm ngày hôm nay để chay.py chọn làm file mới nhất.
# Link lấy từ biến môi trường (GitHub Secrets), không ghi trong mã nguồn:
#   NGUON_CUTTING -> du_lieu/1_CUTTING
#   NGUON_KHO     -> du_lieu/2_TON_KHO
#   NGUON_NESTING -> du_lieu/3_NESTING   (nhiều link, cách nhau bằng dấu xuống dòng hoặc dấu ;)
# Link Apps Script (script.google.com/.../exec) cần đã gắn file apps_script_xuat_du_lieu.gs;
# khóa đặt trong biến <TÊN>_KHOA (vd NGUON_KHO_KHOA) hoặc ghi sẵn trong link (?khoa=...).
import os, re, sys, time, base64, datetime, urllib.request, urllib.parse

CC = os.path.dirname(os.path.abspath(__file__))
DL = os.path.join(os.path.dirname(CC), 'du_lieu')
NGUON = [('NGUON_CUTTING', '1_CUTTING', 'CUTTING_LINK'),
         ('NGUON_KHO', '2_TON_KHO', 'TON_KHO_LINK'),
         ('NGUON_NESTING', '3_NESTING', 'NESTING_LINK')]

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

def tai(url, dich, lan=3):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (bao-cao-chuoi-thep)'})
    for i in range(lan):
        try:
            with urllib.request.urlopen(req, timeout=180) as r: data = r.read()
            break
        except Exception as e:
            if i == lan - 1: raise
            log(f'  lỗi tải ({e}), thử lại sau 10 giây'); time.sleep(10)
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

def main():
    hom_nay = (datetime.datetime.utcnow() + datetime.timedelta(hours=7)).strftime('%d.%m.%Y')
    co_link = False
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
                log(f'Tải {thu_muc} từ {loai}: {n / 1e6:.1f} MB trong {time.time() - t0:.0f} giây')
            except Exception as e:
                # không dừng cả báo cáo: dùng file đã có sẵn trong thư mục (nếu có)
                log(f'KHÔNG tải được {thu_muc} ({loai}): {e}. Dùng file đang có trong du_lieu/{thu_muc}.')
                if os.path.exists(dich): os.remove(dich)
    if not co_link: log('Chưa khai báo link nguồn (NGUON_CUTTING / NGUON_KHO / NGUON_NESTING) – dùng file trong du_lieu/.')

if __name__ == '__main__':
    main()
