import pickle,pandas as pd,numpy as np,re
def U(s): return s.replace('\\_','_').replace('\\*','*') if isinstance(s,str) else s
def num(s): return pd.to_numeric(pd.Series(s).replace({'NaN':np.nan,'':np.nan}),errors='coerce').values
def load_nes(path):
    d=pickle.load(open(path,'rb'))
    # List chi tiet: header row index 3
    L=d['List chi tiet']; h=[str(x) for x in L.iloc[3].tolist()]
    Lm={'Tên chi tiết':'ten_ct','Mác thép':'mac','Quy.Cách.Part':'qc_part','Dày':'day','Rộng':'rong','Dài':'dai','Qty':'qty','KL chi tiết':'kl_ct','KL Tổng':'kl_tong','Quy Cách Vật Tư':'qc_vt','Số Nesting':'nesting','Khoan':'x_khoan','Chấn':'x_chan','Vát mép':'x_vat','Lần-BH':'lan_bh','Ngày B.Hành':'ngay_bh','REV01':'rev01','REV02':'rev02','Hạng Mục':'hang_muc','Nguồn Gốc Vật Tư':'nguon'}
    Ld=L.iloc[4:].copy(); Ld.columns=[Lm.get(x,f'c{i}') for i,x in enumerate(h)]; Ld=Ld.replace({'NaN':np.nan}); Ld['xrow']=Ld.index+2
    Ld=Ld[Ld.nesting.notna()]
    for c in ['day','rong','dai','qty','kl_ct','kl_tong','lan_bh']: Ld[c]=num(Ld[c])
    for c in ['nesting','ten_ct','nguon','qc_vt']: Ld[c]=Ld[c].map(U)
    P=d['Phieu vat tu']; ph=[str(x).split('\n')[-1].strip() if isinstance(x,str) else str(x) for x in P.iloc[2].tolist()]
    Pm={'STT':'stt','Dự Án':'du_an','NESTING NO':'nesting','Hạng Mục':'hang_muc','PCS':'type','Mác Thép':'mac','Quy Cách':'size','(mm)':None,'S.Lượng':'qty','(kg)':None,'Vật Tư':'nguon','Thành Phẩm':'kl_tp','Nguyên Khổ (kg)':'xuat_nk','Remain (kg)':None,'Phế Liệu (kg)':'thu_pl','(Không Thu Hồi) (kg)':'tieu_hao','Hành':'lan_bh','Lệnh Sản Xuất':'lsx'}
    raw=[str(x) for x in P.iloc[2].tolist()]
    cols=[]
    for i,x in enumerate(raw):
        s=x.replace('\n',' ')
        if 'NESTING' in s: cols.append('nesting')
        elif s.startswith('Thickness'): cols.append('t')
        elif s.startswith('Width'): cols.append('w')
        elif s.startswith('Length'): cols.append('l')
        elif s.startswith('Qty'): cols.append('qty')
        elif s.startswith('Weight'): cols.append('weight')
        elif 'Nguồn Gốc' in s: cols.append('nguon')
        elif 'Thành Phẩm' in s: cols.append('kl_tp')
        elif 'Nguyên Khổ' in s: cols.append('xuat_nk')
        elif 'Lệnh Xuất' in s and 'Remain' in s: cols.append('xuat_rm')
        elif 'Thu' in s and 'Remain' in s: cols.append('thu_rm')
        elif 'Tiêu Hao' in s: cols.append('tieu_hao')
        elif 'Phế Liệu (kg)' in s: cols.append('thu_pl')
        elif 'Lần' in s: cols.append('lan_bh')
        elif 'Lệnh Sản Xuất' in s: cols.append('lsx')
        elif s.startswith('Material'): cols.append('mac')
        elif s.startswith('Size'): cols.append('size')
        elif s.startswith('Phase'): cols.append('hang_muc')
        elif s.startswith('Project'): cols.append('du_an')
        else: cols.append(f'c{i}')
    Pd=P.iloc[3:].copy(); Pd.columns=cols; Pd=Pd.replace({'NaN':np.nan}); Pd['xrow']=Pd.index+2; Pd=Pd[Pd.nesting.notna()]
    for c in ['t','w','l','qty','weight','kl_tp','xuat_nk','xuat_rm','thu_rm','thu_pl','tieu_hao','lan_bh']: Pd[c]=num(Pd[c])
    for c in ['nesting','nguon']: Pd[c]=Pd[c].map(U)
    Pd.attrs['cur']=[x for x in P.iloc[0].tolist()+P.iloc[1].tolist() if str(x)!='NaN']
    H=d['Hao hut vat tu']; top={str(H.iloc[i,8]).strip():H.iloc[i,10] for i in range(0,9)}
    return Ld,Pd,top
