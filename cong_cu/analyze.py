exec(open('load.py').read())
import re
import os
TODAY=pd.Timestamp(os.environ.get('NGAY_BAO_CAO','2026-09-28'))
L1=lct(); L1['src']='List Chi Tiết'; L2=lct('2-List Chi Tiết'); L2['src']='2-List Chi Tiết'
L=pd.concat([L1,L2]); L['excel_row']=L.index+2
V=L[L.lsx.notna()].copy()
exec(open('clean.py').read())
V,FIX=lam_sach(V,TODAY)
S=lsxs(); S=S[S.lsx.notna()].copy()
# ---------- per LSX recompute
stages=[('cut','x_cut','d_cut'),('khoan','x_khoan','d_khoan'),('chan','x_chan','d_chan'),('vat','x_vat','d_vat')]
for st,x,dc in stages:
    V['req_'+st]=np.where(V[x].notna(),V.kl_tong,0)
    V['done_'+st]=np.where(V[x].notna()&V[dc].notna(),V.kl_tong,0)
V['done_bg']=np.where(V.d_bg.notna(),V.kl_tong,0)
A=V.groupby('lsx').agg(cong_trinh=('cong_trinh','first'),n_rows=('qty','size'),sl=('qty','sum'),kl=('kl_tong','sum'),
    **{f'{p}_{st}':(f'{p}_{st}','sum') for st in ['cut','khoan','chan','vat'] for p in ['req','done']},done_bg=('done_bg','sum'),
    last_cut=('d_cut_dt','max'),first_cut=('d_cut_dt','min'),last_bg=('d_bg_dt','max'),src=('src','first')).reset_index()
LSXT=S[['lsx','ngay_bh','hang_muc','dv_nhan','nha_may','tinh_trang','st_cut','st_khoan','st_chan','st_vat','st_bg','may']].merge(A,on='lsx',how='left')
LSXT['ngay_bh']=pd.to_datetime(LSXT.ngay_bh,errors='coerce')
# ---------- project
P=V.groupby('cong_trinh').agg(n_lsx=('lsx','nunique'),n_rows=('qty','size'),sl=('qty','sum'),kl=('kl_tong','sum'),
   **{f'{p}_{st}':(f'{p}_{st}','sum') for st in ['cut','khoan','chan','vat'] for p in ['req','done']},done_bg=('done_bg','sum')).reset_index().sort_values('kl',ascending=False)
# ---------- monthly
V['m_cut']=V.d_cut_dt.dt.to_period('M').astype(str); V['m_bg']=V.d_bg_dt.dt.to_period('M').astype(str)
MC=V[V.d_cut_dt.notna()].pivot_table(index='m_cut',columns='may',values='kl_tong',aggfunc='sum',fill_value=0)/1000
MB=V[V.d_bg_dt.notna()].groupby('m_bg').kl_tong.sum()/1000
MS={}
for st in ['khoan','chan','vat']:
    W=V[V['x_'+st].notna()&V['d_'+st+'_dt'].notna()]; MS[st]=W.groupby(W['d_'+st+'_dt'].dt.to_period('M').astype(str)).kl_tong.sum()/1000
# machine
MM=V.groupby(V.may.fillna('(trống)')).agg(n_rows=('qty','size'),sl=('qty','sum'),kl=('kl_tong','sum'),kl_cut=('done_cut','sum'),
    days=('d_cut_dt',lambda s: s.dt.date.nunique()),n_lsx=('lsx','nunique')).reset_index().sort_values('kl',ascending=False)
