import pickle,pandas as pd,numpy as np,re
d=pickle.load(open('kho/sheets.pkl','rb'))
def _sheet(inc, exc=()):
    # tìm sheet theo từ khóa (tên sheet đổi giữa các kỳ: 'TỔNG THÉP XUẤT' / 'BẢNG NHẬT KÝ THÉP XUẤT'...)
    for name in d:
        u = name.strip().upper()
        if all(w in u for w in inc) and not any(w in u for w in exc): return d[name]
    raise KeyError('Không tìm thấy sheet có chữ ' + ' + '.join(inc) + ' trong file tồn kho. Các sheet: ' + ', '.join(d))
def num(s): return pd.to_numeric(pd.Series(s).replace({'NaN':np.nan,'':np.nan}),errors='coerce').values
def nhap():
    df=_sheet(['NHẬP'])
    df=df.copy(); df.columns=['ma','ten','qc','kl','dvt','sl','dvt2','du_an']; df=df.replace({'NaN':np.nan}); df['xrow']=df.index+2
    df['kl']=num(df.kl); df['sl']=num(df.sl); return df
def xuat():
    df=_sheet(['XUẤT'],['NHẬP']).copy(); df.columns=['ngay','ma','ten','qc','bp','kl','dvt','sl','dvt2','du_an']; df=df.replace({'NaN':np.nan}); df['xrow']=df.index+2
    df['kl']=num(df.kl); df['sl']=num(df.sl); df['ngay_dt']=pd.to_datetime(df.ngay,format='%m/%d/%Y',errors='coerce'); return df
def ton():
    df=_sheet(['TỒN'],['NHẬP']).copy(); df.columns=['nhom','ma','ten','qc','kl','dvt','sl','dvt2','du_an']; df=df.replace({'NaN':np.nan}); df['xrow']=df.index+2
    df['kl']=num(df.kl); df['sl']=num(df.sl); return df
