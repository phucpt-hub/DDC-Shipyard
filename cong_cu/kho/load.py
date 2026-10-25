import pickle,pandas as pd,numpy as np,re
d=pickle.load(open('kho/sheets.pkl','rb'))
def _sheet(inc, exc=()):
    # tìm sheet theo từ khóa (tên sheet đổi giữa các kỳ: 'TỔNG THÉP XUẤT' / 'BẢNG NHẬT KÝ THÉP XUẤT'...)
    for name in d:
        u = name.strip().upper()
        if all(w in u for w in inc) and not any(w in u for w in exc): return d[name]
    raise KeyError('Không tìm thấy sheet có chữ ' + ' + '.join(inc) + ' trong file tồn kho. Các sheet: ' + ', '.join(d))
def num(s): return pd.to_numeric(pd.Series(s).replace({'NaN':np.nan,'':np.nan}),errors='coerce').values
def _norm(s): return re.sub(r'\s+',' ',str(s).replace('\\','')).strip().upper()
def _by_header(df, spec, sheet):
    # gán tên cột theo TIÊU ĐỀ (không theo vị trí) – file kho hay thêm cột (Plant, Ghi Chú...)
    hdr=[_norm(h) for h in df.attrs.get('hdr', df.columns)]
    pick={}
    for name, keys in spec:
        i=next((j for j,h in enumerate(hdr) if h in keys and j not in pick.values()),None)
        if i is None: i=next((j for j,h in enumerate(hdr) if any(k in h for k in keys) and j not in pick.values()),None)
        if i is None:
            if len(hdr)==len(spec): pick={n:j for j,(n,_) in enumerate(spec)}; break   # dự phòng: đúng số cột như mẫu cũ
            raise KeyError(f'Sheet "{sheet}" thiếu cột {keys[0]}. Các cột: '+', '.join(hdr))
        pick[name]=i
    out=pd.DataFrame({n:df.iloc[:,j].values for n,j in pick.items()},index=df.index)
    return out.replace({'NaN':np.nan})
def _find(inc, exc=()):
    for name in d:
        u=name.strip().upper()
        if all(w in u for w in inc) and not any(w in u for w in exc): return name
    raise KeyError('Không tìm thấy sheet có chữ ' + ' + '.join(inc) + ' trong file tồn kho. Các sheet: ' + ', '.join(d))
DV=('ĐVT','DVT'); DV2=('ĐVT2','DVT2')
def nhap():
    s=_find(['NHẬP']); df=_by_header(d[s],[('ma',('MÃ VẬT TƯ','MÃ')),('ten',('TÊN VẬT TƯ','TÊN')),('qc',('QUY CÁCH',)),('kl',('KHỐI LƯỢNG','K.LƯỢNG')),('dvt',DV),('sl',('S.LƯỢNG','SỐ LƯỢNG')),('dvt2',DV2),('du_an',('DỰ ÁN',))],s)
    df['xrow']=df.index+2; df['kl']=num(df.kl); df['sl']=num(df.sl); return df
def xuat():
    s=_find(['XUẤT'],['NHẬP']); df=_by_header(d[s],[('ngay',('NGÀY',)),('ma',('MÃ VẬT TƯ','MÃ')),('ten',('TÊN VẬT TƯ','TÊN')),('qc',('QUY CÁCH',)),('bp',('BỘ PHẬN NHẬN','BỘ PHẬN')),('kl',('K.LƯỢNG XUẤT (KG)','K.LƯỢNG','KHỐI LƯỢNG')),('dvt',DV),('sl',('S.LƯỢNG','SỐ LƯỢNG')),('dvt2',DV2),('du_an',('DỰ ÁN',))],s)
    df['xrow']=df.index+2; df['kl']=num(df.kl); df['sl']=num(df.sl); df['ngay_dt']=pd.to_datetime(df.ngay,format='%m/%d/%Y',errors='coerce'); return df
def ton():
    s=_find(['TỒN'],['NHẬP']); df=_by_header(d[s],[('nhom',('PHÂN NHÓM','NHÓM')),('ma',('MÃ VẬT TƯ','MÃ')),('ten',('TÊN VẬT TƯ','TÊN')),('qc',('QUY CÁCH (BATCH)','QUY CÁCH','BATCH')),('kl',('KHỐI LƯỢNG TỒN (KG)','KHỐI LƯỢNG','K.LƯỢNG')),('dvt',DV),('sl',('S.LƯỢNG','SỐ LƯỢNG')),('dvt2',DV2),('du_an',('DỰ ÁN',))],s)
    df['xrow']=df.index+2; df['kl']=num(df.kl); df['sl']=num(df.sl); return df

# ================= FILE TỒN KHO DẠNG SAP (Google Sheet "THEO DÕI SĐC - THÉP - THU HỒI") =================
# Sheet "Tổng hợp nhập - xuất - tồn" = tồn hiện tại theo dự án/lô · sheet "lệnh xuất" = các lần xuất (221 xuất cho dự án, 261 xuất cho LSX)
# Đổi về đúng 3 bảng NHẬP / XUẤT / TỒN như file tồn kho cũ để phần tính cân đối giữ nguyên. NHẬP = TỒN + XUẤT (như file cũ).
SAP_DU_AN={'':'Kho Chung','DG TN SÂN VẬN ĐỘNG HÙNG VƯƠNG':'DG TN SVĐ HÙNG VƯƠNG','DG TN APEC S2 KẾT CẤU':'DG TN APEC S2 (Tổng Hợp)',
 'DG TN FACADE SÂN BAY PHÚ QUỐC':'DG TN SÂN BAY PHÚ QUỐC','DG TN GA T2 CẢNG HÀNG KHÔNG QT PHÚ QUỐC':'DG TN SÂN BAY PHÚ QUỐC',
 'DG XK MOMBASA PORT DEVELOPMENT PROJECT':'DG XK MOMBASA PORT','DG TN CAO TỐC VIN':'CAO TỐC VIN'}
def sap_du_an(s):
    s=re.sub(r'\s+',' ',str(s).replace('\\_','_').replace('\\*','*')).strip()
    if s in ('NaN','nan'): s=''
    if s.startswith(('VT TN ','VT XK ')): s='DG '+s[3:]          # "VT ..." (xuất cho LSX tại xưởng) = cùng dự án với "DG ..."
    return SAP_DU_AN.get(s,s)
def _sap_sheet(*keys):
    for name in d:
        u=_norm(name)
        if all(k in u for k in keys): return name
    return None
SAP_TON=_sap_sheet('TỔNG HỢP','NHẬP','XUẤT','TỒN'); SAP_XUAT=_sap_sheet('LỆNH XUẤT')
IS_SAP=bool(SAP_TON and SAP_XUAT)
def _sap_col(df,sheet,*names,off=0):
    hdr=[_norm(h) for h in df.attrs.get('hdr',df.columns)]
    for n in names:
        if _norm(n) in hdr: return df.iloc[:,hdr.index(_norm(n))+off]
    for n in names:
        j=next((j for j,h in enumerate(hdr) if h.startswith(_norm(n))),None)
        if j is not None: return df.iloc[:,j+off]
    raise KeyError(f'Sheet "{sheet}" thiếu cột {names[0]}. Các cột: '+', '.join(hdr))
def _sap_num(s):
    s=pd.Series(s).astype(str).str.replace(',','',regex=False).str.strip()
    return pd.to_numeric(s.replace({'NaN':np.nan,'':np.nan,'-':np.nan}),errors='coerce').values
def _sap_date(s):
    s=pd.Series(s).astype(str).str.strip()
    iso=pd.to_datetime(s.where(s.str.match(r'^\d{4}-\d{2}-\d{2}')).str[:10],format='%Y-%m-%d',errors='coerce')   # ô ngày thật trong Google Sheet
    mdy=pd.to_datetime(s.where(s.str.match(r'^\d{1,2}/\d{1,2}/\d{4}$')),format='%m/%d/%Y',errors='coerce')        # chữ dạng tháng/ngày/năm (SAP)
    return iso.fillna(mdy)
def _sap_ton():
    S=d[SAP_TON]; c=lambda *n: _sap_col(S,SAP_TON,*n)
    df=pd.DataFrame({'nhom':'','ma':c('Mã vật tư').values,'ten':c('Diễn giải').values,'qc':c('Lô (Batch)','Lô').values,'kl':_sap_num(c('Tổng tồn')),
        'dvt':c('Đơn vị tính chính').values,'sl':_sap_num(c('Tồn quy đổi')),'dvt2':c('Đơn vị tính quy đổi').values,'du_an':c('Tên dự án').map(sap_du_an).values},index=S.index)
    df=df.replace({'NaN':np.nan}); df['du_an']=df.du_an.fillna('Kho Chung'); df['xrow']=df.index+2
    return df[df.kl.fillna(0)!=0].copy()
def _sap_xuat():
    S=d[SAP_XUAT]; c=lambda *n,off=0: _sap_col(S,SAP_XUAT,*n,off=off)
    ngay=_sap_date(c('Base date').values).fillna(_sap_date(c('Requirements date').values))
    df=pd.DataFrame({'ngay':ngay.dt.strftime('%m/%d/%Y').values,'ma':c('Material').values,'ten':c('Material',off=1).values,'qc':c('Material',off=2).values,
        'bp':c('Movement Type Text').values,'kl':_sap_num(c('Removed quantity')),'dvt':c('Base Unit of Measure').values,'sl':_sap_num(c('Removed quantity (HSQĐ)')),
        'dvt2':c('Đơn vị tính quy đổi').values,'du_an':c('Project name').map(sap_du_an).values,'xoa':c('Item Deleted').values},index=S.index)
    df=df.replace({'NaN':np.nan}); df['du_an']=df.du_an.fillna('Kho Chung'); df['xrow']=df.index+2
    df=df[(df.kl.fillna(0)>0)&df.xoa.isna()].drop(columns='xoa').copy()
    df['ngay_dt']=pd.to_datetime(df.ngay,format='%m/%d/%Y',errors='coerce'); return df
if IS_SAP:
    print('File tồn kho dạng SAP: sheet', repr(SAP_TON), '+', repr(SAP_XUAT))
    _T0=_sap_ton(); _X0=_sap_xuat()
    def ton(): return _T0.copy()
    def xuat(): return _X0.copy()
    def nhap():
        k=['ma','ten','qc','kl','dvt','sl','dvt2','du_an']
        df=pd.concat([_T0[k],_X0[k]],ignore_index=True); df['xrow']=df.index+2; return df

# ================= GỘP TÊN DỰ ÁN KHO bị đổi giữa các kỳ (áp dụng cho mọi loại file tồn kho) =================
KHO_ALIAS={'XK THREE (3) SETS RMG GANTRY CRANE':'DG XK THREE (3) SETS RMG GANTRY CRANE','DG XK HSC':'DG XK STRUCTURE HSC PROJECT',
 'DG TN CAO TỐC VIN':'CAO TỐC VIN'}
def kho_alias(s):
    if not isinstance(s,str): return s
    s=re.sub(r'\s+',' ',s.replace('\\_','_')).strip()
    m=re.match(r'^(?:VT|NS) (TN|XK) (.+)$',s)                 # "VT TN ..." / "NS TN ..." = cùng dự án với "DG TN ..."
    if m: s=f'DG {m.group(1)} {m.group(2)}'
    return KHO_ALIAS.get(s,s)
_nhap0,_xuat0,_ton0=nhap,xuat,ton
def _al(df): df['du_an']=df.du_an.map(kho_alias); return df
def nhap(): return _al(_nhap0())
def xuat(): return _al(_xuat0())
def ton(): return _al(_ton0())
