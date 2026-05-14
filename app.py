import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import math
import os
import io
import re
from openpyxl import load_workbook
from copy import copy
from collections import Counter as _C

st.set_page_config(page_title="CALT-LOGIS CLP System", layout="wide")
MAIN_COLOR="#001f3f"; SUB_COLOR="#f0f2f6"; ACCENT_COLOR="#007bff"; ALERT_COLOR="#e74c3c"

st.markdown(f"""<style>
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
* {{ font-family:'Pretendard',sans-serif; }}
.main {{ background-color:#f8f9fa; }}
.header-container {{ background-color:{MAIN_COLOR};padding:25px;border-radius:10px;color:white;margin-bottom:25px;text-align:center;box-shadow:0 4px 6px rgba(0,0,0,0.1); }}
[data-testid="stFileUploadDropzone"] {{ border:2px dashed {MAIN_COLOR};background-color:#eaf1fb;padding:40px;border-radius:12px; margin-top: 10px; }}
[data-testid="stMetric"] {{ background-color:white;padding:15px;border-radius:10px;border-left:5px solid {MAIN_COLOR};box-shadow:0 2px 4px rgba(0,0,0,0.05); }}
.stButton>button {{ width:100%;font-weight:600;color:white;background-color:{MAIN_COLOR};border-radius:8px;border:none;padding:0.5rem 1rem; }}
.container-box {{ background-color:white;padding:20px;border-radius:10px;border:1px solid #e1e4e8;margin-bottom:20px; }}
.guide-table {{ font-size:11px;width:100%;border-collapse:collapse; }}
.guide-table th,.guide-table td {{ border:1px solid #ddd;padding:5px;text-align:center; }}
.guide-table th {{ background-color:#eee; }}
.essential {{ color:{ALERT_COLOR};font-weight:bold; }}
div.row-widget.stRadio > div {{ flex-direction:row; background-color:white; padding: 10px 20px; border-radius: 8px; border: 1px solid #e1e4e8; }}
</style>""", unsafe_allow_html=True)

st.markdown(f'<div class="header-container"><div style="font-size:30px;font-weight:800;">CALT-LOGIS CLP SYSTEM</div><div style="font-size:15px;opacity:0.8;margin-top:5px;">Busan New Port Center | Logistics Management</div></div>', unsafe_allow_html=True)

# ---------------- 유틸리티 함수 ----------------
def clean_num(val):
    try:
        if pd.isna(val) or str(val).strip() in ['','.', 'X','x','NaN','nan']: return 0.0
        return float(str(val).replace(',','').strip())
    except: return 0.0

def reset_data():
    for k in ['bins','manual_mode']:
        if k in st.session_state: del st.session_state[k]

def extract_num(s):
    nums = re.findall(r'\d+', str(s))
    return int(nums[0]) if nums else 999999

# ---------------- 핵심 적재 로직 ----------------
def pack_items_into_bin(pieces, b, max_wt, max_len, max_h=2670):
    for piece in pieces:
        placed = False
        max_stk = piece.get('MAX_STK', 1)
        
        if max_stk > 1 and b['total_W'] + piece['WEIGHT'] <= max_wt:
            dim = (piece['L'], piece['W'], piece['H'])
            base_cnt = sum(1 for r in b['rows'] for i in r['items'] if (i['L'], i['W'], i['H']) == dim)
            if base_cnt > 0:
                stk_cnt = sum(1 for s in b.get('stacked_items', []) if (s['L'], s['W'], s['H']) == dim)
                if stk_cnt < (max_stk - 1) * base_cnt:
                    layers_will_be = 2 + (stk_cnt // base_cnt)
                    if piece['H'] * layers_will_be <= max_h:
                        if 'stacked_items' not in b: b['stacked_items'] = []
                        b['stacked_items'].append(piece)
                        b['total_W'] += piece['WEIGHT']
                        b['groups'].add(piece['GROUP'])
                        placed = True
                        continue
        if not placed:
            row_found = False
            for r in b['rows']:
                tL = max(r['max_L'], piece['L'])
                if r['used_W']+piece['W']<=2340 and b['used_L']+(tL-r['max_L'])<=max_len and b['total_W']+piece['WEIGHT']<=max_wt:
                    r['items'].append(piece); r['used_W']+=piece['W']; b['used_L']+=(tL-r['max_L'])
                    r['max_L']=tL; b['total_W']+=piece['WEIGHT']
                    b['max_W']=max(b['max_W'],piece['W']); b['max_H']=max(b['max_H'],piece['H'])
                    b['groups'].add(piece['GROUP']); row_found=True; placed=True; break
            if not row_found:
                if b['used_L']+piece['L']<=max_len and b['total_W']+piece['WEIGHT']<=max_wt:
                    b['rows'].append({'items':[piece],'used_W':piece['W'],'max_L':piece['L']})
                    b['used_L']+=piece['L']; b['total_W']+=piece['WEIGHT']
                    b['max_W']=max(b['max_W'],piece['W']); b['max_H']=max(b['max_H'],piece['H'])
                    b['groups'].add(piece['GROUP']); placed=True

def apply_labels(bins, max_20_len, max_20_wt, max_dry_h, max_hc_h, max_fr20_len, max_fr20_wt, max_fr40_len, max_fr40_wt):
    for b in bins:
        if b.get('forced_label'): continue
        is_20ft = b['used_L']<=max_20_len and b['total_W']<=max_20_wt
        is_20fr = b['used_L']<=max_fr20_len and b['total_W']<=max_fr20_wt
        is_ow = b['max_W']>2340; is_oh = b['max_H']>max_hc_h
        tags=[]
        if is_oh: tags.append("OH")
        if is_ow: tags.append("OW")
        if tags:
            b['c_label'] = f"{'20ft' if is_20fr else '40ft'} Flat Rack [{' + '.join(tags)}] #{b['id']}"
        else:
            base = "40ft HC" if b['max_H']>max_dry_h else ("20ft Dry" if is_20ft else "40ft Dry")
            b['c_label'] = f"{base} #{b['id']}"
    return bins

def calculate_expert_packing(df, max_40_wt, max_40_len, max_20_wt, max_20_len, max_dry_h, max_hc_h, load_mode):
    raw = []
    for _, row in df.iterrows():
        l,w,h,wt = int(row['L']+.5),int(row['W']+.5),int(row['H']+.5),int(row['WEIGHT']+.5)
        # PKG NO 정렬용 숫자 미리 추출
        p_seq = extract_num(row['PKG NO'])
        raw.append({**row.to_dict(),'L':l,'W':w,'H':h,'WEIGHT':wt, 'p_seq': p_seq})

    all_pieces=[]
    for p in raw:
        rule = p.get('LOAD_KW', '') or load_mode
        if max(p['L'], p['W']) > 2340 or p['H'] > max_hc_h: rule = '4WAY'
        
        s, lg = min(p['L'], p['W']), max(p['L'], p['W'])
        if rule == 'FORK_L': el, ew = (s, lg) if lg <= 2340 else (lg, s)
        elif rule == 'FORK_W': el, ew = (lg, s) if s <= 2340 else (s, lg)
        else: el, ew = (s, lg) if lg <= 2340 else (lg, s)
        all_pieces.append({**p,'L':el,'W':ew})

    # 사이즈 큰 것 우선, 그 다음 PKG NO 숫자 순 정렬
    sk=lambda x:(-x['W'],-x['H'],-x['L'],x['p_seq'])
    fr_p  = sorted([p for p in all_pieces if p['W']>2340 or p['H']>max_hc_h], key=sk)
    hc_p  = sorted([p for p in all_pieces if p['W']<=2340 and max_dry_h<p['H']<=max_hc_h], key=sk)
    dry_p = sorted([p for p in all_pieces if p['W']<=2340 and p['H']<=max_dry_h], key=sk)

    def _pack(pieces, c_no, mwt, mlen):
        if not pieces: return [], c_no
        bins=[]; c=c_no
        for piece in pieces:
            placed=False
            for b in bins:
                pack_items_into_bin([piece],b,mwt,mlen,max_hc_h)
                if piece in b.get('stacked_items',[]) or any(piece in r['items'] for r in b['rows']):
                    placed=True; break
            if not placed:
                nb={'id':c,'rows':[],'used_L':0,'total_W':0,'max_W':0,'max_H':0,'stacked_items':[],'groups':set()}
                pack_items_into_bin([piece],nb,mwt,mlen,max_hc_h); bins.append(nb); c+=1
        return bins, c

    def _fill(bins, cands, mwt, mlen):
        rem=list(cands)
        for b in sorted(bins,key=lambda x:x['used_L']):
            for piece in list(rem):
                before=sum(len(r['items']) for r in b['rows'])+len(b.get('stacked_items',[]))
                pack_items_into_bin([piece],b,mwt,mlen,max_hc_h)
                if (sum(len(r['items']) for r in b['rows'])+len(b.get('stacked_items',[]))) > before: rem.remove(piece)
        return rem

    fr_bins,c = _pack(fr_p,1,30000,11600) # FR 기본값
    rem_hc = _fill(fr_bins,hc_p,30000,11600)
    rem_dry = _fill(fr_bins,dry_p,30000,11600)
    hc_bins,c = _pack(rem_hc,c,max_40_wt,max_40_len)
    rem_dry = _fill(hc_bins,rem_dry,max_40_wt,max_40_len)
    dry_bins,c = _pack(rem_dry,c,max_40_wt,max_40_len)

    res = fr_bins + hc_bins + dry_bins
    for i,b in enumerate(res): b['id']=i+1
    return apply_labels(res,max_20_len,max_20_wt,max_dry_h,max_hc_h,5600,25000,11600,30000)

# ---------------- 메인 UI ----------------
with st.sidebar:
    if os.path.exists("칼트로지스로고.png"): st.image("칼트로지스로고.png", use_column_width=True)
    with st.expander("📄 엑셀 업로드 가이드", expanded=True):
        st.markdown("""<table class="guide-table">
<tr><th>항목</th><th>열</th></tr>
<tr class="essential"><td>PKG NO</td><td>B</td></tr>
<tr class="essential"><td>L / W / H</td><td>J / L / N</td></tr>
</table>""", unsafe_allow_html=True)
    
    st.header("⚙️ 배정 옵션")
    max_20_wt  = st.number_input("최대 중량 (20ft)", 15000,40000,28250)
    max_20_len = st.number_input("최대 길이 (20ft)", 5500,6500,5900)
    max_40_wt  = st.number_input("최대 중량 (40ft)", 20000,40000,29500)
    max_40_len = st.number_input("최대 길이 (40ft)", 11000,13000,11900)
    max_dry_h  = st.number_input("DRY 높이", 2000,3000,2370)
    max_hc_h   = st.number_input("HC 높이",  2000,3500,2670)
    if st.button("🔄 AI 재계산 실행"): reset_data()

st.markdown("### 📤 패킹리스트 업로드")
load_mode_ui = st.radio("👉 **지게차 진입 방향 설정**", ["FORK_L", "FORK_W", "4WAY"], horizontal=True, on_change=reset_data)
file = st.file_uploader("파일을 업로드하세요.", type=['csv','xlsx'], on_change=reset_data)

if file is not None:
    try:
        COL_PKG=1;COL_ITEM=3;COL_DESC=4;COL_QTY=5;COL_WEIGHT=8;COL_L=9;COL_W=11;COL_H=13;COL_REMARK=14;COL_LOAD=15
        raw_full=None; selected_sheet=None

        if file.name.lower().endswith('.xlsx'):
            file.seek(0); sheets=pd.read_excel(file,sheet_name=None,header=None)
            best=0
            for sn,sd in sheets.items():
                if sn=='사용설명서': continue
                cnt=sum(1 for _,r in sd.iterrows() if str(r.iloc[COL_PKG]).strip().lower() not in ['','nan'] and clean_num(r.iloc[COL_L])>0)
                if cnt>best: best=cnt; raw_full=sd; selected_sheet=sn
        else:
            file.seek(0); raw_full=pd.read_csv(file,header=None)

        p_data=[]
        for oi in range(len(raw_full)):
            row=raw_full.iloc[oi]
            if len(row) <= COL_H: continue
            pkg=str(row.iloc[COL_PKG]).replace('00:00:00','').replace('.0','').strip()
            lv, wv, hv = clean_num(row.iloc[COL_L]), clean_num(row.iloc[COL_W]), clean_num(row.iloc[COL_H])
            if pkg.lower() in ['','nan','.'] or lv==0: continue
            qty=int(clean_num(row.iloc[COL_QTY])) if clean_num(row.iloc[COL_QTY])>0 else 1
            wtv=clean_num(row.iloc[COL_WEIGHT]); rem=str(row.iloc[COL_REMARK]).upper(); load_val=str(row.iloc[COL_LOAD]).upper()
            ms = int(re.search(r'(\d+)단', rem).group(1)) if re.search(r'(\d+)단', rem) else 1
            load_kw = "FORK_W" if "FORK_W" in load_val else ("FORK_L" if "FORK_L" in load_val else ("4WAY" if "4WAY" in load_val else ""))
            repeat = qty if 'BOX' in rem else 1
            for seq in range(repeat):
                sfx = f"-{seq+1:03d}" if repeat > 1 else ""
                p_data.append({'PKG NO':f"{pkg}{sfx}",'ITEM':str(row.iloc[COL_ITEM]),'GROUP':str(row.iloc[COL_DESC]),
                               'L':lv,'W':wv,'H':hv,'WEIGHT':wtv/repeat,'MAX_STK':ms,'LOAD_KW':load_kw,'row_idx':oi})

        df=pd.DataFrame(p_data)
        if not df.empty:
            if 'bins' not in st.session_state or not st.session_state.get('manual_mode',False):
                st.session_state.bins=calculate_expert_packing(df,max_40_wt,max_40_len,max_20_wt,max_20_len,max_dry_h,max_hc_h, load_mode_ui)
                st.session_state.manual_mode=True

            bins=st.session_state.bins
            st.subheader("📊 적재 요약")
            c1,c2,c3,c4=st.columns(4)
            packed=sum(len(r['items']) for b in bins for r in b['rows'])+sum(len(b.get('stacked_items',[])) for b in bins)
            c1.metric("전체 화물",f"{len(df)} PKG"); c2.metric("배정 완료",f"{packed} PKG")
            c3.metric("컨테이너 수",f"{len(bins)} UNIT"); c4.metric("평균 중량",f"{sum(b['total_W'] for b in bins)/max(1, len(bins)):,.0f} kg")

            for b in bins:
                st.markdown('<div class="container-box">', unsafe_allow_html=True)
                hc1,hc2,hc3=st.columns([4,2,1])
                hc1.markdown(f"### 📦 {b['c_label']}")
                
                # 제원 한도 계산
                lbl=b['c_label']
                if '20ft Dry' in lbl: cml,cmw,cmh=max_20_len,max_20_wt,max_dry_h
                elif '40ft HC' in lbl: cml,cmw,cmh=max_40_len,max_40_wt,max_hc_h
                elif '40ft Dry' in lbl: cml,cmw,cmh=max_40_len,max_40_wt,max_dry_h
                else: cml,cmw,cmh=max_40_len,max_40_wt,max_hc_h # 기본 FR급

                # 상단 진행바 (이미지 구성 복구)
                uw=max([r['used_W'] for r in b['rows']]+[0])
                uh=b['max_H'] + (max([s['H'] for s in b.get('stacked_items',[])]+[0]) if b.get('stacked_items') else 0)
                cc1,cc2,cc3,cc4=st.columns(4)
                cc1.markdown(f"**📏 길이:** {b['used_L']:,}/{cml:,}mm"); cc1.progress(min(1.0,b['used_L']/cml))
                cc2.markdown(f"**⚖️ 중량:** {b['total_W']:,}/{cmw:,}kg"); cc2.progress(min(1.0,b['total_W']/cmw))
                cc3.markdown(f"**↔️ 폭:** {uw:,}/2340mm"); cc3.progress(min(1.0,uw/2340))
                cc4.markdown(f"**↕️ 높이:** {uh:,}/{cmh:,}mm"); cc4.progress(min(1.0,uh/cmh))

                # 시각화 (이미지 cb6abc 스타일 복구)
                stk_cnt=_C((s['L'],s['W'],s['H']) for s in b.get('stacked_items',[]))
                base_cnt=_C((i['L'],i['W'],i['H']) for r in b['rows'] for i in r['items'])
                
                fig=go.Figure()
                fig.add_shape(type="rect",x0=0,y0=0,x1=cml,y1=2340,line=dict(color=MAIN_COLOR,width=2))
                cx=0
                for r in b['rows']:
                    cy=(2340-r['used_W'])/2
                    for item in r['items']:
                        dim=(item['L'],item['W'],item['H'])
                        layers=1+math.ceil(stk_cnt.get(dim,0)/max(1,base_cnt.get(dim,1)))
                        # 색상 로직 복구
                        if item['W']>2340 or item['H']>max_hc_h: ic=ALERT_COLOR
                        elif item['H']>max_dry_h: ic="#e67e22"
                        else: ic=ACCENT_COLOR
                        
                        border=dict(color="#FFD700",width=3) if layers>1 else dict(color="white",width=1)
                        fig.add_shape(type="rect",x0=cx,y0=cy,x1=cx+item['L'],y1=cy+item['W'],fillcolor=ic,opacity=0.85,line=border)
                        lbl_txt=f"×{layers}단<br>{item['PKG NO']}" if layers>1 else f"{item['PKG NO']}"
                        fig.add_annotation(x=cx+item['L']/2,y=cy+item['W']/2,text=lbl_txt,showarrow=False,font=dict(color="white",size=9))
                        cy+=item['W']
                    cx+=r['max_L']
                
                fig.add_shape(type="line",x0=b['used_L'],y0=-100,x1=b['used_L'],y1=2440,line=dict(color=ALERT_COLOR,width=2,dash="dash"))
                fig.update_layout(xaxis=dict(visible=False,range=[-100,cml+200]),yaxis=dict(visible=False,range=[-100,2440]),
                                  height=280,margin=dict(l=10,r=10,t=10,b=10),paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig,use_container_width=True,key=f"plot_{b['id']}")

                # 데이터 에디터 (이동 기능 유지)
                t_data=[]
                for r in b['rows']:
                    for item in r['items']: t_data.append({**item,'위치':'바닥','이동':f"{b['id']}번"})
                for s in b.get('stacked_items',[]): t_data.append({**s,'위치':'단적','이동':f"{b['id']}번"})
                df_edit=pd.DataFrame(t_data)[['위치','PKG NO','ITEM','L','W','H','WEIGHT','이동']]
                st.data_editor(df_edit,hide_index=True,use_container_width=True,key=f"ed_{b['id']}",
                               column_config={"이동":st.column_config.SelectboxColumn("🚚 이동",options=[f"{bx['id']}번" for bx in bins]+["새 컨테이너"])})
                st.markdown('</div>',unsafe_allow_html=True)

            # 다운로드
            if selected_sheet:
                file.seek(0); wb=load_workbook(file); ws=wb[selected_sheet]
                tc=ws.max_column+1; ws.cell(row=1,column=tc).value="배정 결과"
                mapping={}
                for bx in bins:
                    for it in ([i for r in bx['rows'] for i in r['items']]+bx.get('stacked_items',[])):
                        rid=it['row_idx']+1; mapping[rid]=mapping.get(rid,[])+[bx['c_label']]
                for rid,lbls in mapping.items():
                    ws.cell(row=rid,column=tc).value=", ".join(set(lbls))
                out=io.BytesIO(); wb.save(out)
                st.download_button("📥 최종 결과 다운로드",out.getvalue(),"CLP_RESULT.xlsx",use_container_width=True)

    except Exception as e: st.error(f"오류: {e}")
