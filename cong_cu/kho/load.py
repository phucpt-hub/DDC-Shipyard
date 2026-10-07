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
    s=_find(['TỒN'],['NHẬP']); df=_by_header(d[s],[('nhom',('PHÂN NHÓM','NHÓM')),('ma',('MÃ VẬT TƯ','MÃ')),('ten',('TÊN VẬT TƯ','TÊN')),('qc',('QUY CÁCH (BATCH)','QUY CÁCH')),('kl',('KHỐI LƯỢNG TỒN (KG)','KHỐI LƯỢNG','K.LƯỢNG')),('dvt',DV),('sl',('S.LƯỢNG','SỐ LƯỢNG')),('dvt2',DV2),('du_an',('DỰ ÁN',))],s)
    df['xrow']=df.index+2; df['kl']=num(df.kl); df['sl']=num(df.sl); return df
