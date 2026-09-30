import pickle,pandas as pd,numpy as np,re
d=pickle.load(open('kho/sheets.pkl','rb'))
def num(s): return pd.to_numeric(pd.Series(s).replace({'NaN':np.nan,'':np.nan}),errors='coerce').values
def nhap():
    df=d[' TỔNG HỢP NHẬP (TỒN + XUẤT)'] if ' TỔNG HỢP NHẬP (TỒN + XUẤT)' in d else d['TỔNG HỢP NHẬP (TỒN + XUẤT)']
    df=df.copy(); df.columns=['ma','ten','qc','kl','dvt','sl','dvt2','du_an']; df=df.replace({'NaN':np.nan}); df['xrow']=df.index+2
    df['kl']=num(df.kl); df['sl']=num(df.sl); return df
def xuat():
    df=d['TỔNG THÉP XUẤT'].copy(); df.columns=['ngay','ma','ten','qc','bp','kl','dvt','sl','dvt2','du_an']; df=df.replace({'NaN':np.nan}); df['xrow']=df.index+2
    df['kl']=num(df.kl); df['sl']=num(df.sl); df['ngay_dt']=pd.to_datetime(df.ngay,format='%m/%d/%Y',errors='coerce'); return df
def ton():
    df=d['TỔNG TỒN HIỆN TẠI'].copy(); df.columns=['nhom','ma','ten','qc','kl','dvt','sl','dvt2','du_an']; df=df.replace({'NaN':np.nan}); df['xrow']=df.index+2
    df['kl']=num(df.kl); df['sl']=num(df.sl); return df
