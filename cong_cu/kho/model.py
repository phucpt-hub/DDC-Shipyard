exec(open('kho/load.py').read()); exec(open('kho/keys.py').read())
N=nhap();X=xuat();T=ton()
for df in (N,X,T):
    r=df.ten.map(stock_key); df['loai']=r.str[0]; df['key']=r.str[1]; df['mac']=r.str[2].str.replace('k.gửi','',regex=False).str.strip(' ,').replace('','?'); df['fam']=df.mac.map(grade_family)
    df['t']=pd.to_numeric(df.key.str.extract(r'^PL([\d.]+)$')[0])
    m=df.qc.astype(str).str.extract(r'^([\d.]+)X([\d.]+)([A-Z]*)$'); df['bw']=pd.to_numeric(m[0]); df['bl']=pd.to_numeric(m[1]); df['org']=m[2]
    mm=df.ten.str.extract(r'Thép tấm\s+[\d.]+x(\d+)x(\d+)'); df['bw']=df.bw.fillna(pd.to_numeric(mm[0])/1000); df['bl']=df.bl.fillna(pd.to_numeric(mm[1])/1000)
    df['kgd']=df.ten.str.contains('k.gửi',regex=False); df['loco']=df.ten.str.contains('lỡ cỡ',regex=False)
iso=pd.to_datetime(X.ngay.where(X.ngay.astype(str).str.match(r'^\d{4}-')),errors='coerce')
X['date_fix']=X.ngay_dt.fillna(iso.map(lambda t: pd.Timestamp(t.year,t.day,t.month) if pd.notna(t) else pd.NaT)); X['date_swapped']=iso.notna()
MAP={'DG TN SVĐ HÙNG VƯƠNG':'10626-051 DG SÂN VẬN ĐỘNG HÙNG VƯƠNG','DG TN APEC S2 (Tổng Hợp)':'10626-048 DG APEC S2','DG TN APEC S3':'10625-036 DGRP APEC S3',
'DG TN SÂN BAY PHÚ QUỐC':'10625-030 DGRP GA T2 CẢNG HÀNG KHÔNG QT PHÚ QUỐC; 10626-010 DGRP FACADE SÂN BAY PHÚ QUỐC','DG TN CẦU ĐI BỘ SÔNG SÀI GÒN - CẦU CHÍNH':'10625-031 DGRP CẦU ĐI BỘ SÔNG SÀI GÒN',
'DG TN CẦU ĐI BỘ QUA SÔNG SÀI GÒN':'10625-031 DGRP CẦU ĐI BỘ SÔNG SÀI GÒN','DG XK MOMBASA PORT':'10725-006 DGRP MOMBASA PORT DEVELOPMENT PROJECT','DG XK THREE (3) SETS RMG GANTRY CRANE':'10726-075 RMG CRANE',
'DG XK SHIPLIFT ÚC':'10726-141 DG XK SHIPLIFT ÚC','DG TN DU LỊCH SINH THÁI NÚI CHỨA CHAN ĐN':'10626-022 DU LỊCH SINH THÁI NGHỈ DƯỠNG GIẢI TRÍ NÚI CHỨA CHAN','DG TN QUẢNG TRƯỜNG TTTP & TTHC THÀNH PHỐ':'10626-130 DG TN QUẢNG TRƯỜNG TTTP & TTHC THÀNH PHỐ',
'DG XK STRUCTURE HSC PROJECT':'10726-049 \\_ HSC Project \\_ Mockup Casing','DG XK BPI HO REDEV SUPPLY MATERIAL':'10726-040 DG XK BPI HO REDEV SUPPLY MATERIAL','CAO TỐC VIN':'TCTN.DNS.26.004 ĐƯỜNG CAO TỐC VIN','DG TN CẦU CỎ MAY':'10626-125 DG TN CẦU CỎ MAY'}
C2I={}
for k,v in MAP.items():
    for c in v.split('; '): C2I.setdefault(c,k)
C2I['10625-031 DGRP CẦU ĐI BỘ SÔNG SÀI GÒN']='DG TN CẦU ĐI BỘ SÔNG SÀI GÒN - CẦU CHÍNH'
import os
_cfg=os.environ.get('FILE_CAU_HINH')
if _cfg and os.path.exists(_cfg):
    try:
        _m=pd.read_excel(_cfg,sheet_name='MAP_DU_AN',header=None,skiprows=4,usecols=[0,1]).dropna(subset=[0])
        for a,b in _m.itertuples(index=False):
            a=str(a).strip()
            if a.startswith('Danh sách'): break
            a2=a.replace('_','\\_')
            if isinstance(b,str) and b.strip():
                for k in (a,a2): C2I[k]=b.strip()
        print('Đã đọc MAP_DU_AN từ file cấu hình:',len(_m),'dòng')
    except Exception as e: print('Không đọc được file cấu hình:',e)
exec(open('kho/demand.py').read())
C['proj_kho']=C.cong_trinh.map(C2I).fillna('(chưa map)')
kk=C.apply(lambda r: demand_key(r.qc_part,r.qc_vt,r.day),axis=1); C['loai']=kk.str[0]; C['key']=kk.str[1]
C['fam']=C.mac.map(grade_family); C['Lbar']=pd.to_numeric(C.qc_vt,errors='coerce'); C['used']=C.dai*C.qty
G=C.groupby(['cong_trinh','nesting'],dropna=False).agg(proj_kho=('proj_kho','first'),lsx=('lsx','first'),n=('qty','size'),ncut=('cut','sum'),loai=('loai','first'),key=('key','first'),mac=('mac','first'),fam=('fam','first'),
   nguon=('nguon','first'),qc_vt=('qc_vt','first'),is_pl=('is_pl','first'),pt=('pt','first'),pw=('pw','first'),pl=('pl','first'),kl=('kl_tong','sum'),Lbar=('Lbar','first'),used=('used','sum'),
   may=('may','first'),ngay_bh=('ngay_bh','first'),xrow=('xrow','min'),src=('src','first')).reset_index()
G['st']=np.where(G.ncut==0,'Chưa cắt',np.where(G.ncut==G.n,'Đã cắt hết','Cắt dở'))
def phoi(r):
    if r.is_pl and pd.notna(r.pw): return r.pt*r.pw*r.pl*7.85e-6,'Tấm theo QUY CÁCH VẬT TƯ (dày×rộng×dài×7,85)'
    if pd.notna(r.Lbar) and r.used>0 and r.loai!='PL': return r.kl*r.Lbar/r.used,'Cây = KL chi tiết × chiều dài cây / tổng chiều dài chi tiết'
    return r.kl/0.85,'Ước tính KL chi tiết / 0,85 (không đọc được quy cách phôi)'
pp=G.apply(phoi,axis=1); G['kl_phoi']=pp.str[0]; G['cach_tinh']=pp.str[1]
DEM=G[G.st=='Chưa cắt'].copy()
