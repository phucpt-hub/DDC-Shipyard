# Unified pipeline -> dash/data.json  (Cutting + Nesting + Kho)
import json, re, pickle, numpy as np, pandas as pd
BASE=''
import os
REPORT_DATE=os.environ.get('NGAY_BAO_CAO','2026-09-28')
# ---------------- CUTTING ----------------
exec(open(BASE+'kho/model.py').read())          # N,X,T  C,G,DEM, MAP,C2I
KX=X; KC=C
exec(open(BASE+'analyze.py').read())          # V (đã chuẩn hóa), LSXT, P, MC, MB, MM, FIX
CUT_V=V.copy(); CUT_LSXT=LSXT.copy(); X=KX; C=KC
from openpyxl.utils import get_column_letter as _gcl
_COLNAME={'d_cut':'NGÀY CUTTING','d_cut_dt':'NGÀY CUTTING','d_khoan':'NGÀY KHOAN','d_chan':'NGÀY CHẤN','d_vat':'NGÀY VÁT MÉP','d_bg':'NGÀY BÀN GIAO','d_bg_dt':'NGÀY BÀN GIAO',
  'x_khoan':'CÔNG VIỆC KHOAN','x_chan':'CÔNG VIỆC CHẤN','x_vat':'CÔNG VIỆC VÁT MÉP','g_nk':'NGUYÊN KHỔ','g_td':'TẬN DỤNG','may':'MÁY CẮT','kl_tong':'KL TỔNG'}
def _colref(c):
    if not c: return ''
    b=c.replace('_dt','')
    return (_gcl(LCT_COLS.index(b)+1)+' · ' if b in LCT_COLS else '')+_COLNAME.get(c,c)
CLEAN_LOG=[]
for f in FIX:
    idx=np.flatnonzero(f['mask'])
    for j,i in enumerate(idx):
        r=CUT_V.iloc[i]
        CLEAN_LOG.append(dict(file=os.environ.get('FILE_CUTTING','Cutting'),sheet=r.src,row=int(r.excel_row),obj=str(r.lsx),du_an=str(r.cong_trinh).replace('\\_','_'),
            cot=_colref(f['col']),cu=f['old'][j],moi=f['new'][j],ma=f['ma'],quy_tac=f['quy_tac'],nhom=f['nhom'],kl=float(r.kl_tong) if pd.notna(r.kl_tong) else 0))
# kho: ngày xuất bị đảo ngày/tháng
_sw=X[X.date_swapped&X.date_fix.notna()]
for r in _sw.itertuples():
    CLEAN_LOG.append(dict(file=os.environ.get('FILE_KHO','Tồn kho'),sheet='TỔNG THÉP XUẤT',row=int(r.xrow),obj=str(r.ma),du_an=str(r.du_an),cot='A · NGÀY',cu=str(r.ngay)[:10],moi=r.date_fix.strftime('%d/%m/%Y'),
        ma='K01',quy_tac='NGÀY XUẤT bị Excel đọc đảo ngày/tháng (ngày ≤ 12) → đổi lại đúng ngày/tháng',nhom='Đã chuẩn hóa',kl=float(r.kl or 0)))
FX=pd.DataFrame([dict(ma=f['ma'],quy_tac=f['quy_tac'],nhom=f['nhom'],row=i) for f in FIX for i in np.flatnonzero(f['mask'])])
exec(open(BASE+'nes_generic.py').read())
f=lambda x: None if x is None or (isinstance(x,float) and not np.isfinite(x)) else round(float(x),3)
t_=lambda kg: f(kg/1000)
def short(ct):
    s=re.sub(r'^[\w.]+-?\d+\s+','',ct) if re.match(r'^\d',ct) or ct.startswith('TCTN') else ct
    s=s.replace('\\_','').replace('DGRP ','').replace('DG XK ','').replace('DG TN ','').replace('DG ','').strip(' _')
    return s.title() if s.isupper() else s
NAMES={'10626-051':'SVĐ Hùng Vương','10725-006':'Mombasa Port','10626-010':'Facade Sân bay Phú Quốc','10626-048':'APEC S2','10625-036':'APEC S3','10625-030':'Ga T2 Phú Quốc','10625-031':'Cầu đi bộ sông Sài Gòn','10726-141':'Shiplift Úc','10626-130':'Quảng trường TTTP & TTHC','10726-075':'RMG Crane','10626-022':'Núi Chứa Chan','TCTN.DNS.26.004':'Cao tốc VIN','10726-040':'BPI HO Redev','10726-049':'HSC Mockup Casing','10626-125':'Cầu Cỏ May'}
def short(ct):
    c=ct.split(' ')[0]
    return NAMES.get(c,ct.replace('\\_','').strip())
DATE_MIN=(pd.Timestamp(REPORT_DATE)-pd.DateOffset(months=11)).replace(day=1)
months=[str(p) for p in pd.period_range(DATE_MIN,REPORT_DATE,freq='M')]
# ---------- stock balance at project-kho level (thickness key, NK only) ----------
Tpl=T.copy()
own=Tpl.groupby(['du_an','key']).kl.sum()
POOL=Tpl[Tpl.du_an.isin(['Kho Chung','TỒN BRAVO'])].groupby('key').kl.sum().to_dict()
D=DEM[DEM.nguon!='Tận dụng'].groupby(['proj_kho','key','loai']).agg(need=('kl_phoi','sum'),n=('nesting','size')).reset_index()
D['own']=[own.get((a,b),0) for a,b in zip(D.proj_kho,D.key)]
D['from_own']=np.minimum(D.need,D.own); D['short1']=D.need-D.from_own
D=D.sort_values('short1',ascending=False); pool=dict(POOL); fc=[]
for r in D.itertuples():
    a=min(r.short1,pool.get(r.key,0)); pool[r.key]=pool.get(r.key,0)-a; fc.append(a)
D['from_common']=fc; D['buy']=D.short1-D.from_common
DT=DEM[DEM.nguon=='Tận dụng'].groupby('proj_kho').kl_phoi.sum()
# ---------- nesting files ----------
NES_FILES=json.loads(os.environ.get('NESTING_FILES','[]'))  # [[pkl, file_name], ...]
NES={}
_nrm=lambda x: pd.Series(x).astype(str).str.upper().str.replace(r'[^A-Z0-9]','',regex=True)
_CN=C.assign(_k=_nrm(C.nesting).values)
NES_SKIP=[]
for pkl,fname in NES_FILES:
    try:
        _L,_P,_t=load_nes(pkl)
        # 1) gán dự án theo số nesting trùng với file Cutting
        ks=set(_nrm(_P.nesting.dropna()))
        hit=_CN[_CN._k.isin(ks)].cong_trinh.value_counts()
        ct=hit.index[0] if len(hit) else None; how='trùng số nesting'
        # 2) không trùng: theo mã dự án trong LSX (VT-HV, VT-FD, VT-S2...)
        if ct is None:
            tok=pd.concat([_P.lsx.dropna().astype(str),_P.nesting.dropna().astype(str)]).str.extract(r'(VT-[A-Z0-9]+)')[0].dropna()
            if len(tok):
                tk=tok.mode().iloc[0]; cc=C[C.lsx.astype(str).str.contains(tk,regex=False)].cong_trinh.mode()
                if len(cc): ct=cc.iloc[0]; how='mã '+tk
        # 3) theo mã trong tên file (vd LSX-VT-BPI-TH-001 → 'BPI')
        if ct is None:
            m=re.search(r'VT-([A-Z0-9]+)-T[TH]',fname.upper())
            if m:
                cc=C[C.lsx.astype(str).str.upper().str.contains('-'+m.group(1)+'-',regex=False)].cong_trinh.mode()
                if len(cc): ct=cc.iloc[0]; how='tên file '+m.group(1)
        if ct: NES.setdefault(ct,[]).append((pkl,fname)); print('Nesting',fname,'->',ct,'(',how,')')
        else: print('Nesting',fname,': chưa gán được dự án'); NES_SKIP.append(fname)
    except Exception as e: print('Lỗi đọc nesting',fname,e); NES_SKIP.append(fname)
norm=lambda s: pd.Series(s).astype(str).str.upper().str.replace(r'[^A-Z0-9]','',regex=True).values
def gclass(m):
    s=str(m).upper().replace(' ','')
    if 'AS3678' in s: return 'AS3678-350'
    if 'S690' in s or 'STRENX' in s or 'S460' in s: return 'CĐ SIÊU CAO'
    return {'CĐ CAO (~345-355MPa)':'CĐ 345-355','THƯỜNG (~235-300MPa)':'THƯỜNG'}.get(grade_family(m),'?')
Ts=T[T.loai=='PL'].copy(); Ts['gc']=Ts.mac.map(gclass); Ts['avail']=Ts.sl.where(Ts.sl>0,np.nan)
def allocate(pend,pk):
    stock=Ts.copy(); res=[]
    for r in pend.itertuples():
        w,l=r.w/1000,r.l/1000; fit=(stock.t==r.t)&(stock.bw>=w-0.01)&(stock.bl>=l-0.05)&(stock.avail>=1); same=stock.gc==r.gc; lab='Phải mua'
        for lb,m in [('Tồn dự án',fit&same&(stock.du_an==pk)),('Kho chung/BRAVO',fit&same&stock.du_an.isin(['Kho Chung','TỒN BRAVO'])),('Điều chuyển dự án khác',fit&same&~stock.du_an.isin([pk,'Kho Chung','TỒN BRAVO']))]:
            idx=stock[m].sort_values(['bw','bl']).index
            if len(idx): stock.loc[idx[0],'avail']-=1; lab=lb; break
        res.append(lab)
    return res
def nesting_block(ct,files,pk):
    Pn=pd.concat([load_nes(p)[1].assign(srcfile=f) for p,f in files]); fname=', '.join(f for _,f in files)
    Pn=Pn[Pn.t.notna()].copy()
    if not len(Pn): print('Nesting',fname,': không có tấm (thép hình) – bỏ qua phần cân đối tấm'); return None
    _d=Pn[Pn.duplicated('nesting',keep='last')]
    for r in _d.itertuples(): CLEAN_LOG.append(dict(file=r.srcfile,sheet='Phieu vat tu',row=int(r.xrow),obj=str(r.nesting),du_an=ct,cot='NESTING NO',cu=str(r.nesting),moi='(bỏ, dùng dòng sau)',ma='N01',quy_tac='Số nesting lặp lại trong Phiếu vật tư → giữ dòng ban hành sau cùng',nhom='Đã chuẩn hóa',kl=float(r.weight or 0)))
    Pn=Pn.drop_duplicates('nesting',keep='last')
    _m=Pn.mac.astype(str); _mn=_m.str.upper().str.replace(' ','')
    for r,a,b in zip(Pn.itertuples(),_m,_mn):
        if a!=b and a not in ('nan','NaN'): CLEAN_LOG.append(dict(file=r.srcfile,sheet='Phieu vat tu',row=int(r.xrow),obj=str(r.nesting),du_an=ct,cot='MATERIAL',cu=a,moi=b,ma='N02',quy_tac='Mác thép viết khác kiểu (hoa/thường, dấu cách) → chuẩn hóa để ghép kho',nhom='Đã chuẩn hóa',kl=float(r.weight or 0)))
    _bal=Pn.weight-Pn.kl_tp.fillna(0)-Pn.thu_rm.fillna(0)-Pn.thu_pl.fillna(0)-Pn.tieu_hao.fillna(0)
    for r,bv in zip(Pn.itertuples(),_bal):
        if pd.notna(bv) and abs(bv)>1: CLEAN_LOG.append(dict(file=r.srcfile,sheet='Phieu vat tu',row=int(r.xrow),obj=str(r.nesting),du_an=ct,cot='WEIGHT',cu=f'{r.weight:.2f}',moi=f'lệch {bv:+.2f} kg',ma='N03',quy_tac='Weight ≠ thành phẩm + remain + phế + tiêu hao (lệch > 1 kg) → giữ nguyên, cần phòng Nesting kiểm tra',nhom='Cần xác nhận',kl=float(r.weight or 0)))
    Cc=C[C.cong_trinh==ct].copy(); Cc['nk']=norm(Cc.nesting); Pn['nk']=norm(Pn.nesting)
    cs=Cc.groupby('nk').agg(n=('cut','size'),nc=('cut','sum'))
    def st(nk):
        if nk not in cs.index: return 'Chưa nạp Cutting'
        r=cs.loc[nk]; return 'Đã cắt hết' if r.nc==r.n else ('Chưa cắt' if r.nc==0 else 'Cắt dở')
    Pn['cut_st']=Pn.nk.map(st); Pn['mac']=Pn.mac.astype(str).str.upper().str.replace(' ','')
    Pn['vao']=np.where(Pn.nguon=='NEWPLATE','Tấm mới',np.where(Pn.nguon=='TAN DUNG','Tận dụng','Remain'))
    s=lambda c: float(Pn[c].fillna(0).sum())
    new=float(Pn.loc[Pn.vao=='Tấm mới','weight'].sum()); rin=float(Pn.loc[Pn.vao!='Tấm mới','weight'].sum())
    tp,trm,tpl,th=s('kl_tp'),s('thu_rm'),s('thu_pl'),s('tieu_hao')
    byt=Pn.groupby('t').agg(n=('nesting','nunique'),w=('weight','sum'),tp=('kl_tp','sum'),pl=('thu_pl','sum'),rm=('thu_rm','sum')).reset_index()
    stc=Pn.groupby('cut_st').weight.agg(['size','sum'])
    pend=Pn[(Pn.cut_st!='Đã cắt hết')&(Pn.vao=='Tấm mới')].sort_values(['lsx','nesting']).copy(); pend['gc']=pend.mac.map(gclass)
    pend['alloc']=allocate(pend,pk) if len(pend) else []
    al=pend.groupby('alloc').weight.agg(['size','sum'])
    buy=pend[pend.alloc=='Phải mua'].groupby(['t','w','l','mac']).agg(n=('qty','sum'),kg=('weight','sum')).reset_index().sort_values('kg',ascending=False)
    # LSX reconciliation
    lsx_n=set(Pn.lsx.dropna().astype(str)); lsx_c=set(Cc.lsx.dropna().astype(str))
    miss=Pn[Pn.cut_st=='Chưa nạp Cutting'].groupby('lsx').weight.sum().sort_values(ascending=False)
    return dict(file=fname,n_nesting=int(Pn.nesting.nunique()),n_lsx=len(lsx_n),
      flow=dict(new=t_(new),rin=t_(rin),tp=t_(tp),rm=t_(trm),pl=t_(tpl),th=t_(th)),
      pct=dict(remain=f(trm/(new+rin)) if new+rin else None,phe=f(tpl/(new+rin)) if new+rin else None,hh=f(tpl/tp) if tp else None,tp=f(tp/(new+rin)) if new+rin else None),
      by_t=[dict(t=f(r.t),n=int(r.n),w=t_(r.w),tp=t_(r.tp),pl=t_(r.pl),rm=t_(r.rm)) for r in byt.itertuples()],
      cut_st=[dict(k=k,n=int(v['size']),t=t_(v['sum'])) for k,v in stc.iterrows()],
      alloc=[dict(k=k,n=int(v['size']),t=t_(v['sum'])) for k,v in al.iterrows()],
      buy=[dict(qc=f"PL{r.t:g}×{r.w:g}×{r.l:g}",mac=r.mac,n=int(r.n),t=t_(r.kg)) for r in buy.head(15).itertuples()],
      lsx_not_in_cut=[dict(lsx=k,t=t_(v)) for k,v in miss.items()][:15],
      lsx_cut_not_in_nes=sorted(lsx_c-lsx_n)[:15])
# ---------------- PER PROJECT ----------------
V=CUT_V.reset_index(drop=True); V['m_cut']=V.d_cut_dt.dt.to_period('M').astype(str); V['m_bg']=V.d_bg_dt.dt.to_period('M').astype(str)
kho_tot=lambda df,pk: float(df.loc[df.du_an==pk,'kl'].sum())
PR=[]
for ct,Vp in V.groupby('cong_trinh'):
    kl=Vp.kl_tong.sum(); req=Vp.req_cut.sum(); cut=Vp.done_cut.sum(); bg=Vp.done_bg.sum()
    Lx=CUT_LSXT[CUT_LSXT.cong_trinh==ct].drop_duplicates('lsx').copy()
    Lx['pc']=Lx.done_cut/Lx.req_cut.replace(0,np.nan); Lx['pb']=Lx.done_bg/Lx.kl.replace(0,np.nan)
    Lx['st']=np.where(Lx.pc.fillna(0)>=0.999,'Đã cắt xong',np.where(Lx.pc.fillna(0)>0,'Đang cắt','Chưa cắt'))
    Lx=Lx.sort_values('ngay_bh',ascending=False)
    mc=Vp[Vp.d_cut_dt.notna()].groupby('m_cut').kl_tong.sum(); mb=Vp[Vp.d_bg_dt.notna()].groupby('m_bg').kl_tong.sum()
    mach=Vp[Vp.done_cut>0].groupby(Vp.may.fillna('(trống)')).done_cut.sum().sort_values(ascending=False)
    iss=FX[FX.row.isin(np.flatnonzero((V.cong_trinh==ct).values))].assign(kl=lambda x: V.kl_tong.values[x.row]).groupby(['ma','quy_tac','nhom'],sort=False).agg(n=('row','size'),kl=('kl','sum')).reset_index()
    Gp=G[G.cong_trinh==ct]; gst=Gp.groupby('st').agg(n=('nesting','size'),kl=('kl_phoi','sum'))
    pk=C2I.get(ct)
    kho=None
    if pk:
        Dp=D[D.proj_kho==pk]; Tp=T[T.du_an==pk]
        xm=X[(X.du_an==pk)&X.date_fix.notna()].groupby(X.date_fix.dt.to_period('M').astype(str)).kl.sum()
        tk=Tp.groupby('key').kl.sum().sort_values(ascending=False)
        shared=[c for c,k in C2I.items() if k==pk and c!=ct and c in set(V.cong_trinh)]
        kho=dict(name=pk,shared=[short(s) for s in shared],nhap=t_(kho_tot(N,pk)),xuat=t_(kho_tot(X,pk)),ton=t_(kho_tot(T,pk)),
          ton_key=[dict(k=k,t=t_(v)) for k,v in tk.head(12).items()],
          xuat_m=[t_(xm.get(m,0)) for m in months],
          need=t_(Dp.need.sum()),own=t_(Dp.from_own.sum()),common=t_(Dp.from_common.sum()),buy=t_(Dp.buy.sum()),tandung=t_(DT.get(pk,0)),
          bal=[dict(k=r.key,loai=r.loai,n=int(r.n),need=t_(r.need),own=t_(r.own),fo=t_(r.from_own),fc=t_(r.from_common),buy=t_(r.buy)) for r in Dp.sort_values('need',ascending=False).itertuples()])
    nes=nesting_block(ct,NES[ct],pk) if ct in NES else None
    PR.append(dict(id=ct,name=short(ct),code=ct.split(' ')[0] if re.match(r'^[\d]',ct) else '',kl=t_(kl),req=t_(req),cut=t_(cut),bg=t_(bg),
      pc=f(cut/req) if req else None,pb=f(bg/kl) if kl else None,
      stages={s:dict(req=t_(Vp['req_'+s].sum()),done=t_(Vp['done_'+s].sum())) for s in ['cut','khoan','chan','vat']},
      n_lsx=int(Lx.lsx.nunique()),lsx_st={k:int(v) for k,v in Lx.st.value_counts().items()},
      last_cut=str(Vp.d_cut_dt.max().date()) if Vp.d_cut_dt.notna().any() else None,
      m_cut=[t_(mc.get(m,0)) for m in months],m_bg=[t_(mb.get(m,0)) for m in months],
      mach=[dict(k=k,t=t_(v)) for k,v in mach.items()],
      lsx=[dict(lsx=r.lsx,hm=(None if pd.isna(r.hang_muc) else str(r.hang_muc).replace('\\_','_'))[:60] if not pd.isna(r.hang_muc) else None,bh=None if pd.isna(r.ngay_bh) else str(r.ngay_bh.date()),kl=t_(r.kl or 0),pc=f(r.pc),pb=f(r.pb),st=r.st,
               last=None if pd.isna(r.last_cut) else str(r.last_cut.date()),may=None if pd.isna(r.may) else str(r.may)) for r in Lx.itertuples()],
      fix=[dict(c=r.ma,d=r.quy_tac,g=r.nhom,n=int(r.n),t=t_(r.kl)) for r in iss.itertuples()],
      nes_st={k:dict(n=int(v.n),t=t_(v.kl)) for k,v in gst.iterrows()},
      kho=kho,nes=nes))
PR.sort(key=lambda p:-p['kl'])
# alerts
for p in PR:
    a=[]
    k=p['kho']
    if k and k['buy'] and k['buy']>0.5: a.append(('bad',f"Thiếu {k['buy']:,.1f} t thép tấm/hình cho nesting chưa cắt – phải mua hoặc điều chuyển"))
    if k and k['common'] and k['common']>0.5: a.append(('warn',f"Cần lấy {k['common']:,.1f} t từ Kho chung/BRAVO"))
    if p['nes']:
        m=sum(x['t'] or 0 for x in p['nes']['lsx_not_in_cut'])
        if m>0.5: a.append(('bad',f"{len(p['nes']['lsx_not_in_cut'])} LSX có trong file nesting nhưng chưa nạp vào Cutting ({m:,.1f} t)"))
        b=[x for x in p['nes']['alloc'] if x['k']=='Phải mua']
        if b: a.append(('bad',f"Nesting: {b[0]['n']} tấm mới chưa cắt không có phôi trong kho ({b[0]['t']:,.1f} t)"))
    if p['pc'] is not None and p['pc']<0.999 and p['last_cut'] and p['last_cut']<'2026-08-28': a.append(('warn',f"Còn {p['req']-p['cut']:,.1f} t chưa cắt, lần cắt cuối {p['last_cut']}"))
    p['alerts']=[dict(s=s,m=m) for s,m in a]
# totals
tot=dict(kl=sum(p['kl'] for p in PR),req=sum(p['req'] for p in PR),cut=sum(p['cut'] for p in PR),bg=sum(p['bg'] for p in PR),n_lsx=sum(p['n_lsx'] for p in PR),n_proj=len(PR))
m_all=V[V.d_cut_dt.notna()].groupby('m_cut').kl_tong.sum(); b_all=V[V.d_bg_dt.notna()].groupby('m_bg').kl_tong.sum()
mach_all=V[V.done_cut>0].groupby(V.may.fillna('(trống)')).done_cut.sum().sort_values(ascending=False)
bal=dict(need=t_(D.need.sum()),own=t_(D.from_own.sum()),common=t_(D.from_common.sum()),buy=t_(D.buy.sum()))
bykey=D.groupby('key')[['need','from_own','from_common','buy']].sum().sort_values('buy',ascending=False)
nomap=sorted(set(V.cong_trinh)-set(C2I))
kho_only=T[~T.du_an.isin(set(C2I.values())|{'Kho Chung','TỒN BRAVO'})].groupby('du_an').kl.sum().sort_values(ascending=False)
errs=FX.assign(kl=lambda x: V.kl_tong.values[x.row]).groupby(['ma','quy_tac','nhom'],sort=False).agg(n=('row','size'),kl=('kl','sum')).reset_index()

# ---------------- CHUỖI NGÀY (cho báo cáo ngày/tuần/tháng/quý/năm) ----------------
PIDX={p['id']:i for i,p in enumerate(PR)}
MACH=sorted(V.may.fillna('(chưa ghi máy)').astype(str).unique().tolist())
MIDX={m:i for i,m in enumerate(MACH)}
_ev=[]
def _add(df,dcol,stage,wcol='kl_tong'):
    d=df[df[dcol].notna()&(df[dcol]<=pd.Timestamp(REPORT_DATE))]
    if not len(d): return
    g=d.assign(_d=d[dcol].dt.normalize(),_p=d.cong_trinh.map(PIDX),_m=d.may.fillna('(chưa ghi máy)').astype(str).map(MIDX)).groupby(['_d','_p','_m'],dropna=False)[wcol].sum()
    for (dd,pp,mm),v in g.items():
        if v and pd.notna(pp): _ev.append([dd,int(pp),-1 if pd.isna(mm) else int(mm),stage,round(v/1000,4)])
_add(V[V.req_cut>0],'d_cut_dt',0)
for i,st in enumerate(['khoan','chan','vat']): _add(V[V['x_'+st].notna()],'d_'+st+'_dt',i+1)
_add(V,'d_bg_dt',4)
_L=CUT_LSXT.drop_duplicates('lsx'); _L=_L[_L.ngay_bh.notna()]
for r in _L.itertuples():
    if pd.notna(r.kl) and r.cong_trinh in PIDX and r.ngay_bh<=pd.Timestamp(REPORT_DATE): _ev.append([r.ngay_bh.normalize(),PIDX[r.cong_trinh],-1,5,round(r.kl/1000,4)])
_k2p={}
for p in PR:
    if p['kho'] and p['kho']['name'] not in _k2p: _k2p[p['kho']['name']]=PIDX[p['id']]
_x=X[X.date_fix.notna()&X.du_an.isin(_k2p)]
for (dd,da),v in _x.groupby([_x.date_fix.dt.normalize(),'du_an']).kl.sum().items():
    if dd<=pd.Timestamp(REPORT_DATE): _ev.append([dd,_k2p[da],-1,6,round(v/1000,4)])
_d0=min(e[0] for e in _ev) if _ev else pd.Timestamp(REPORT_DATE)
EV=[[int((e[0]-_d0).days)]+e[1:] for e in _ev]
# mục tiêu tháng: sheet MUC_TIEU trong file cấu hình, không có thì = TB 3 tháng đủ gần nhất
_mt={}
if os.environ.get('FILE_CAU_HINH'):
    try:
        _t=pd.read_excel(os.environ['FILE_CAU_HINH'],sheet_name='MUC_TIEU',header=None)
        for r in _t.itertuples(index=False):
            m=re.match(r'^(\d{4})-(\d{1,2})',str(r[0]))
            if m and pd.notna(r[1]): _mt[f"{m.group(1)}-{int(m.group(2)):02d}"]=dict(cut=float(r[1]),bg=float(r[2]) if len(r)>2 and pd.notna(r[2]) else None)
    except Exception as e: print('Không có sheet MUC_TIEU:',e)
_full=[m for m in months if m<REPORT_DATE[:7]][-3:]
_base=round(float(np.mean([m_all.get(m,0) for m in _full]))/1000/100)*100 if _full else 0

OUT=dict(date=REPORT_DATE,months=months,projects=PR,tot=tot,bal=bal,
  kho=dict(nhap=t_(N.kl.sum()),xuat=t_(X.kl.sum()),ton=t_(T.kl.sum()),common=t_(T[T.du_an=='Kho Chung'].kl.sum()),bravo=t_(T[T.du_an=='TỒN BRAVO'].kl.sum())),
  m_cut=[t_(m_all.get(m,0)) for m in months],m_bg=[t_(b_all.get(m,0)) for m in months],mach=[dict(k=k,t=t_(v)) for k,v in mach_all.items()],
  bykey=[dict(k=k,need=t_(r.need),own=t_(r.from_own),fc=t_(r.from_common),buy=t_(r.buy)) for k,r in bykey.head(14).iterrows()],
  kho_only=[dict(k=k,t=t_(v)) for k,v in kho_only.items()],nomap=nomap,
  fix=[dict(c=r.ma,d=r.quy_tac,g=r.nhom,n=int(r.n),t=t_(r.kl)) for r in errs.itertuples()],
  files=dict(cutting=os.environ.get('FILE_CUTTING','DDC_SHIPYARD - KHGC - SC'),kho=os.environ.get('FILE_KHO','TỒN KHO'),nesting=[f for v in NES.values() for _,f in v],nes_skip=NES_SKIP))
_CL=pd.DataFrame(CLEAN_LOG)
if len(_CL):
    _g=_CL.groupby(['ma','quy_tac','nhom','file'],sort=False).agg(n=('row','size'),kl=('kl','sum')).reset_index()
    OUT['fix']=[dict(c=r.ma,d=r.quy_tac,g=r.nhom,f=r.file,n=int(r.n),t=t_(r.kl)) for r in _g.itertuples()]
OUT['fix_file']='nhat_ky_lam_sach.xlsx'
OUT['ev']=dict(d0=_d0.strftime('%Y-%m-%d'),mach=MACH,rows=EV)
OUT['target']=dict(base=_base,by_month=_mt)
pickle.dump(CLEAN_LOG,open(os.environ.get('OUT_LOG','clean_log.pkl'),'wb'))
json.dump(OUT,open(os.environ.get('OUT_JSON','data.json'),'w'),ensure_ascii=False,default=str)
print(json.dumps(tot,ensure_ascii=False),bal,len(json.dumps(OUT))//1024,'KB')
for p in PR: print(p['name'][:40],p['kl'],p['pc'],p['kho']['buy'] if p['kho'] else '-',len(p['alerts']))
