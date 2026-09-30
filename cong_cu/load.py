import pickle,pandas as pd,numpy as np
d=pickle.load(open('sheets.pkl','rb'))
def num(s): return pd.to_numeric(s.replace({'NaN':np.nan,'':np.nan}),errors='coerce')
LCT_COLS=['cong_trinh','lsx','c2','ten_ct','mac','qc_part','day','rong','dai','qty','kl_ct','kl_tong','qc_vt','nesting','may','ca','c16',
 'x_cut','x_khoan','x_chan','x_vat','c21','d_cut','d_khoan','d_chan','d_vat','d_bg','c27','lan_bh','ngay_bh','note','nha_may','c32','chk_date','chk_kl','c35','g_nk','g_td','nhom']
def lct(name='List Chi Tiết'):
    df=d[name].iloc[3:].copy(); df.columns=LCT_COLS
    df=df.replace({'NaN':np.nan})
    for c in ['day','rong','dai','qty','kl_ct','kl_tong']: df[c]=num(df[c])
    return df
LSX_COLS=['stt','lsx','ngay_bh','ngay_mail','cong_trinh','hang_muc','dv_cat','dv_gc','dv_nhan','nha_may','tong_sl','tong_kl','tinh_trang','phieu_xuat','phieu_thu','c15','st_cut','st_khoan','st_chan','st_vat','st_bg','c21','kl_cut','ht_cut','kl_khoan','ht_khoan','kl_chan','ht_chan','kl_vat','ht_vat','kl_bg','ht_bg','may','hoan_thanh']
def lsxs():
    df=d['LỆNH SX'].iloc[2:].copy(); df.columns=LSX_COLS; df=df.replace({'NaN':np.nan})
    for c in ['tong_sl','tong_kl','kl_cut','ht_cut','kl_khoan','ht_khoan','kl_chan','ht_chan','kl_vat','ht_vat','kl_bg','ht_bg']: df[c]=num(df[c])
    return df
