import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import math
import os
import io
import re
from openpyxl import load_workbook
from copy import copy

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
/* 라디오 버튼 스타일 강조 */
div.row-widget.stRadio > div {{ flex-direction:row; background-color:white; padding: 10px 20px; border-radius: 8px; border: 1px solid #e1e4e8; }}
</style>""", unsafe_allow_html=True)

st.markdown(f'<div class="header-container"><div style="font-size:30px;font-weight:800;">CALT-LOGIS CLP SYSTEM</div><div style="font-size:15px;opacity:0.8;margin-top:5px;">Busan New Port Center | Logistics Management</div></div>', unsafe_allow_html=True)

def clean_num(val):
    try:
        if pd.isna(val) or str(val).strip() in ['','.',  'X','x','NaN','nan']: return 0.0
        return float(str(val).replace(',','').strip())
    except: return 0.0

def reset_data():
    for k in ['bins','manual_mode']:
        if k in st.session_state: del st.session_state[k]

with st.sidebar:
    if os.path.exists("칼트로지스로고.png"): st.image("칼트로지스로고.png", use_column_width=True)
    with st.expander("📄 엑셀 업로드 표준 규격", expanded=True):
        st.markdown(f"""<table class="guide-table">
<tr><th>항목</th><th>엑셀 열</th><th>구분</th></tr>
<tr class="essential"><td>No.of PKG</td><td>B 열</td><td>[필수]</td></tr>
<tr class="essential"><td>LENGTH (L)</td><td>J 열</td><td>[필수]</td></tr>
<tr class="essential"><td>WIDTH (W)</td><td>L 열</td><td>[필수]</td></tr>
<tr class="essential"><td>HEIGHT (H)</td><td>N 열</td><td>[필수]</td></tr>
<tr><td>GROSS WEIGHT</td><td>I 열</td><td>선택</td></tr>
<tr><td>ITEM</td><td>D 열</td><td>참고</td></tr>
<tr><td>REMARK</td><td>O 열</td><td>선택</td></tr>
<tr><td style="color:#007bff;font-weight:bold;">LOAD방향</td><td style="color:#007bff;font-weight:bold;">P 열</td><td>선택</td></tr>
</table>
<div style='font-size:10px;color:#555;margin-top:8px;line-height:1.7;'>
<b>REMARK 키워드 (O열)</b><br>
· <b>BOX</b> : Q'ty = 실제 박스 수<br>
· <b>N단</b> (예: 2단, 5단) : 숫자만큼 다단적재<br>
</div>
<div style='font-size:10px;color:#555;margin-top:8px;line-height:1.7;'>
<b>LOAD방향 제어 (P열)</b><br>
· <b>FORK_L</b><br>
· <b>FORK_W</b><br>
· <b>4WAY</b>
</div>""", unsafe_allow_html=True)
        st.markdown("---")
        tpl = {"Invoice No":[""],"No.of PKG":["PKG-001"],"LOCATION":[""],"ITEM":["SAMPLE ITEM"],
               "Description of Goods":["DETAIL DESC"],"Q'ty":[1],"UNIT":["EA"],
               "Net Weight (kg)":[500],"Gross Weight (kg)":[550],
               "Dimension L (mm)":[1200],"X1":["X"],"Dimension W (mm)":[1000],"X2":["X"],
               "Dimension H (mm)":[2300],"REMARK":[""],"LOAD":["FORK_L"]}
        tow=io.BytesIO(); pd.DataFrame(tpl).to_excel(tow,index=False,engine='openpyxl')
        st.download_button("📥 신규 화주용 양식 다운로드",tow.getvalue(),"CALT_CLP_TEMPLATE.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)

    st.header("⚙️ 배정 옵션 설정")
    with st.expander("⚖️ 컨테이너 제원", expanded=False):
        st.markdown("**🟦 20ft DRY**")
        max_20_wt  = st.number_input("최대 중량 (kg)", 15000,40000,28250,key="i_20wt")
        max_20_len = st.number_input("최대 길이 (mm)", 5500,6500,5900,key="i_20len")
        st.markdown("---")
        st.markdown("**🟫 40ft DRY / HC**")
        max_40_wt  = st.number_input("최대 중량 (kg)", 20000,40000,29500,key="i_40wt")
        max_40_len = st.number_input("최대 길이 (mm)", 11000,13000,11900,key="i_40len")
        max_dry_h  = st.number_input("DRY 내부 높이 (mm)", 2000,3000,2370,key="i_dryh")
        max_hc_h   = st.number_input("HC 내부 높이 (mm)",  2000,3500,2670,key="i_hch")
        st.markdown("---")
        st.markdown("**🔴 20ft FR (Flat Rack)**")
        max_fr20_wt  = st.number_input("최대 중량 (kg)", 10000,40000,25000,key="i_fr20wt")
        max_fr20_len = st.number_input("최대 길이 (mm)", 4000,7000,5600,key="i_fr20len")
        max_fr_h     = st.number_input("내부 높이 (mm)", 1000,3000,2260,key="i_frh")
        max_fr_w     = st.number_input("기둥간 폭 (mm)", 1000,3000,2080,key="i_frw")
        st.markdown("---")
        st.markdown("**🔴 40ft FR (Flat Rack)**")
        max_fr40_wt  = st.number_input("최대 중량 (kg)", 10000,45000,30000,key="i_fr40wt")
        max_fr40_len = st.number_input("최대 길이 (mm)", 8000,14000,11600,key="i_fr40len")

    if st.button("🔄 AI 재계산 실행"):
        st.session_state['manual_mode'] = False


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
        raw.append({**row.to_dict(),'L':l,'W':w,'H':h,'WEIGHT':wt})

    sg={}
    for p in raw:
        rule = p.get('LOAD_KW', '')
        if rule not in ['FORK_L', 'FORK_W', '4WAY']:
            rule = load_mode 
        
        is_fr_size = max(p['L'], p['W']) > 2340 or p['H'] > max_hc_h
        if is_fr_size:
            rule = '4WAY' 
            
        p['eff_rule'] = rule
        key=(min(p['L'],p['W']),max(p['L'],p['W']),p['H']>max_dry_h, rule)
        sg.setdefault(key,[]).append(p)
        
    all_pieces=[]
    for (s,lg,_,rule),items in sg.items():
        n=len(items); ca=lg<=2340; cb=s<=2340
        
        # 💡 로직 스왑 적용 완료
        if rule == 'FORK_L':
            # 지게차가 긴 쪽(L)으로 진입 -> 긴 쪽이 컨테이너의 폭(W) 방향으로 향함 (가로 눕힘)
            el, ew = (s, lg) if ca else (lg, s)
        elif rule == 'FORK_W':
            # 지게차가 짧은 쪽(W)으로 진입 -> 긴 쪽이 컨테이너의 길이(L) 방향으로 향함 (세로 세움)
            el, ew = (lg, s) if cb else (s, lg)
        else: # 4WAY
            if ca and cb:
                sa=max(1,int(2340//lg)); sb=max(1,int(2340//s))
                La=math.ceil(n/sa)*s; Lb=math.ceil(n/sb)*lg
                el,ew=(lg,s) if (Lb<La and sb>=2) else (s,lg)
            elif cb: el,ew=lg,s
            else: el,ew=s,lg
            
        for p in items: all_pieces.append({**p,'L':el,'W':ew})

    sk=lambda x:(-x['W'],-x['H'],-x['L'],x['GROUP'])
    fr_p  = sorted([p for p in all_pieces if p['W']>2340 or p['H']>max_hc_h], key=sk)
    hc_p  = sorted([p for p in all_pieces if p['W']<=2340 and max_dry_h<p['H']<=max_hc_h], key=sk)
    dry_p = sorted([p for p in all_pieces if p['W']<=2340 and p['H']<=max_dry_h], key=sk)

    def _pack(pieces, c_no, mwt=None, mlen=None):
        wt=mwt or max_40_wt; ln=mlen or max_40_len
        if not pieces: return [],c_no
        bins=[]; c=c_no
        for piece in pieces:
            placed=False
            for b in bins:
                pack_items_into_bin([piece],b,wt,ln,max_hc_h)
                if piece in b.get('stacked_items',[]) or any(piece in r['items'] for r in b['rows']):
                    placed=True; break
            if not placed:
                nb={'id':c,'rows':[],'used_L':0,'total_W':0,'max_W':0,'max_H':0,'stacked_items':[],'groups':set()}
                pack_items_into_bin([piece],nb,wt,ln,max_hc_h); bins.append(nb); c+=1
        return [b for b in bins if b['used_L']>0 or b.get('stacked_items')], c

    def _fill(bins, cands, mwt=None, mlen=None):
        wt=mwt or max_40_wt; ln=mlen or max_40_len
        rem=list(cands)
        for b in sorted(bins,key=lambda x:x['used_L']):
            for piece in list(rem):
                before=sum(len(r['items']) for r in b['rows'])+len(b.get('stacked_items',[]))
                pack_items_into_bin([piece],b,wt,ln,max_hc_h)
                after=sum(len(r['items']) for r in b['rows'])+len(b.get('stacked_items',[]))
                if after>before: rem.remove(piece)
        return rem

    fr_bins,c = _pack(fr_p,1,max_fr40_wt,max_fr40_len)
    rem_hc  = _fill(fr_bins,hc_p,max_fr40_wt,max_fr40_len)
    rem_dry = _fill(fr_bins,dry_p,max_fr40_wt,max_fr40_len)
    hc_bins,c = _pack(rem_hc,c)
    rem_dry = _fill(hc_bins,rem_dry)
    dry_bins,c = _pack(rem_dry,c)

    bins=fr_bins+hc_bins+dry_bins
    for i,b in enumerate(bins): b['id']=i+1
    return apply_labels(bins,max_20_len,max_20_wt,max_dry_h,max_hc_h,max_fr20_len,max_fr20_wt,max_fr40_len,max_fr40_wt)


# 💡 UI 텍스트 간소화 완료
st.markdown("### 📤 기본 설정 및 패킹리스트 업로드")

load_mode_ui = st.radio(
    "👉 **지게차(포크) 기본 진입 방향 설정** (※ P열 개별 지정 우선)", 
    ["FORK_L", "FORK_W", "4WAY"],
    horizontal=True,
    on_change=reset_data
)
default_load_mode = load_mode_ui

file = st.file_uploader("방향을 선택한 후 이곳에 파일을 끌어다 놓으세요.", type=['csv','xlsx'], on_change=reset_data)

if file is not None:
    try:
        COL_PKG=1;COL_ITEM=3;COL_DESC=4;COL_QTY=5;COL_WEIGHT=8;COL_L=9;COL_W=11;COL_H=13;COL_REMARK=14;COL_LOAD=15
        raw_full=None; selected_sheet=None

        if file.name.lower().endswith('.xlsx'):
            file.seek(0)
            sheets=pd.read_excel(file,sheet_name=None,header=None)
            best=0
            for sn,sd in sheets.items():
                if sn=='사용설명서': continue
                if sd.shape[1] <= COL_H: continue  
                cnt=sum(1 for _,r in sd.iterrows() if
                    str(r.iloc[COL_PKG] if pd.notna(r.iloc[COL_PKG]) else '').strip().lower() not in ['','nan','none','.']
                    and clean_num(r.iloc[COL_L])>0 and clean_num(r.iloc[COL_W])>0 and clean_num(r.iloc[COL_H])>0)
                if cnt>best: best=cnt; raw_full=sd; selected_sheet=sn
            if raw_full is None: st.error("필수 데이터를 찾지 못했습니다."); st.stop()
            st.info(f"읽은 시트: {selected_sheet}")
        else:
            file.seek(0); raw_full=pd.read_csv(file,header=None)

        p_data=[]
        for oi in range(len(raw_full)):
            row=raw_full.iloc[oi]
            if len(row) <= COL_H: continue 
            try:
                pkg=str(row.iloc[COL_PKG]).replace('00:00:00','').replace('.0','').strip() if pd.notna(row.iloc[COL_PKG]) else ""
                lv=clean_num(row.iloc[COL_L]); wv=clean_num(row.iloc[COL_W]); hv=clean_num(row.iloc[COL_H])
                if pkg.lower() in ['','nan','none','.'] or lv==0 or wv==0 or hv==0: continue
                wtv=clean_num(row.iloc[COL_WEIGHT]) if len(row) > COL_WEIGHT else 0
                qty=int(clean_num(row.iloc[COL_QTY])) if (len(row) > COL_QTY and clean_num(row.iloc[COL_QTY])>0) else 1
                
                rem = str(row.iloc[COL_REMARK]).strip().upper() if (len(row) > COL_REMARK and pd.notna(row.iloc[COL_REMARK])) else ""
                is_box = 'BOX' in rem
                match = re.search(r'(\d+)단', rem)
                ms = int(match.group(1)) if match else 1
                
                load_val = str(row.iloc[COL_LOAD]).strip().upper() if (len(row) > COL_LOAD and pd.notna(row.iloc[COL_LOAD])) else ""
                if "FORK_W" in load_val: load_kw = "FORK_W"
                elif "FORK_L" in load_val: load_kw = "FORK_L"
                elif "4WAY" in load_val: load_kw = "4WAY"
                else: load_kw = ""
                
                repeat=qty if is_box else 1
                for seq in range(repeat):
                    sfx=f"-{seq+1:03d}" if repeat>1 else ""
                    p_data.append({'PKG NO':f"{pkg}{sfx}",'GROUP':str(row.iloc[COL_DESC]) if (len(row) > COL_DESC and pd.notna(row.iloc[COL_DESC])) else "-",
                        'ITEM':str(row.iloc[COL_ITEM]) if (len(row) > COL_ITEM and pd.notna(row.iloc[COL_ITEM])) else "-",
                        'DESC':str(row.iloc[COL_DESC]) if (len(row) > COL_DESC and pd.notna(row.iloc[COL_DESC])) else "-",
                        'L':lv,'W':wv,'H':hv,'WEIGHT':wtv,'MAX_STK':ms,'STACK_OK':ms>1,'LOAD_KW':load_kw,'row_idx':int(oi)})
            except: continue

        df=pd.DataFrame(p_data)
        if df.empty: st.warning("⚠️ 필수 데이터를 찾을 수 없습니다.")
        else:
            if 'bins' not in st.session_state or not st.session_state.get('manual_mode',False):
                st.session_state.bins=calculate_expert_packing(df,max_40_wt,max_40_len,max_20_wt,max_20_len,max_dry_h,max_hc_h, default_load_mode)
                st.session_state.manual_mode=True

            bins=st.session_state.bins
            topts=[f"{b['id']}번" for b in bins]+["✨ 새 컨테이너"]
            TYPE_SPECS={'20ft Dry':{'ml':max_20_len,'mw':max_20_wt,'mh':max_dry_h,'lb':'20ft Dry'},
                        '40ft Dry':{'ml':max_40_len,'mw':max_40_wt,'mh':max_dry_h,'lb':'40ft Dry'},
                        '40ft HC' :{'ml':max_40_len,'mw':max_40_wt,'mh':max_hc_h, 'lb':'40ft HC'},
                        '20ft FR' :{'ml':max_fr20_len,'mw':max_fr20_wt,'mh':max_fr_h,'lb':'20ft Flat Rack'},
                        '40ft FR' :{'ml':max_fr40_len,'mw':max_fr40_wt,'mh':max_fr_h,'lb':'40ft Flat Rack'}}
            def get_type(lbl):
                for k in ['20ft Dry','40ft HC','40ft Dry','20ft Flat Rack','40ft Flat Rack']:
                    if k in lbl: return {'20ft Flat Rack':'20ft FR','40ft Flat Rack':'40ft FR'}.get(k,k)
                return '40ft Dry'

            st.subheader("📊 실시간 적재 요약")
            c1,c2,c3,c4=st.columns(4)
            packed=sum(len(r['items']) for b in bins for r in b['rows'])+sum(len(b.get('stacked_items',[])) for b in bins)
            c1.metric("전체 화물",f"{len(df)} PKG"); c2.metric("배정 완료",f"{packed} PKG")
            c3.metric("컨테이너 수",f"{len(bins)} UNIT"); c4.metric("평균 중량",f"{sum(b['total_W'] for b in bins)/max(1, len(bins)):,.0f} kg")

            for b in bins:
                if b['used_L']==0 and not b.get('stacked_items'): continue
                st.markdown('<div class="container-box">', unsafe_allow_html=True)
                hc1,hc2,hc3=st.columns([4,2,1])
                hc1.markdown(f"### 📦 {b['c_label']}")
                cur_type=get_type(b['c_label'])
                new_type=hc2.selectbox("타입",list(TYPE_SPECS.keys()),
                    index=list(TYPE_SPECS.keys()).index(cur_type) if cur_type in TYPE_SPECS else 1,
                    key=f"ts_{b['id']}",label_visibility="collapsed")
                if hc3.button("🔄",key=f"rc_{b['id']}"):
                    sp=TYPE_SPECS[new_type]; all_it=[i for r in b['rows'] for i in r['items']]+b.get('stacked_items',[])
                    nb={'id':b['id'],'rows':[],'used_L':0,'total_W':0,'max_W':0,'max_H':0,'stacked_items':[],'groups':set()}
                    ov=[]
                    for it in all_it:
                        bf=sum(len(r['items']) for r in nb['rows'])+len(nb.get('stacked_items',[]))
                        pack_items_into_bin([it],nb,sp['mw'],sp['ml'],sp['mh'])
                        af=sum(len(r['items']) for r in nb['rows'])+len(nb.get('stacked_items',[]))
                        if af==bf: ov.append(it)
                    h_ex=[i for i in ([x for r in nb['rows'] for x in r['items']]+nb.get('stacked_items',[])) if i['H']>sp['mh']]
                    nb['c_label']=f"{sp['lb']} #{b['id']}"; nb['forced_label']=True
                    upd=[]
                    for bx in st.session_state.bins:
                        upd.append(nb if bx['id']==b['id'] else bx)
                    if ov:
                        mid=max(bx['id'] for bx in upd)
                        ob={'id':mid+1,'rows':[],'used_L':0,'total_W':0,'max_W':0,'max_H':0,'stacked_items':[],'groups':set()}
                        for it in ov: pack_items_into_bin([it],ob,max_40_wt,max_40_len,max_hc_h)
                        upd.append(ob)
                        st.toast(f"⚠️ {len(ov)}개 화물 제원 초과 → 새 컨테이너 #{mid+1}로 분리됐습니다.")
                    if h_ex: st.toast(f"🚨 {len(h_ex)}개 화물 H 초과! CLP 작성 시 확인하세요.")
                    st.session_state.bins=apply_labels(sorted(upd,key=lambda x:x['id']),
                        max_20_len,max_20_wt,max_dry_h,max_hc_h,max_fr20_len,max_fr20_wt,max_fr40_len,max_fr40_wt)
                    st.rerun()

                base_items=[item for r in b['rows'] for item in r['items']]
                base_min_L=min((i['L'] for i in base_items),default=0)
                base_min_W=min((i['W'] for i in base_items),default=0)
                t_data=[]
                for r in b['rows']:
                    for item in r['items']: t_data.append({**item,'위치':'바닥','이동':f"{b['id']}번",'⚠️':''})
                for s in b.get('stacked_items',[]):
                    same=[i for i in base_items if i['L']==s['L'] and i['W']==s['W'] and i['H']==s['H']]
                    danger='' if same else ('🚨' if (s['L']>base_min_L or s['W']>base_min_W) else '')
                    t_data.append({**s,'위치':'단적','이동':f"{b['id']}번",'⚠️':danger})
                df_edit=pd.DataFrame(t_data)[['⚠️','위치','PKG NO','ITEM','L','W','H','WEIGHT','이동']]
                edited_df=st.data_editor(df_edit,hide_index=True,use_container_width=True,key=f"ed_{b['id']}",
                    column_config={"이동":st.column_config.SelectboxColumn("🚚 이동",options=topts)},
                    disabled=['⚠️','위치','PKG NO','ITEM','L','W','H','WEIGHT'])

                if st.button(f"🚀 {b['id']}번 변경사항 적용",key=f"btn_{b['id']}"):
                    moves=[(r['PKG NO'],r['이동']) for _,r in edited_df.iterrows() if r['이동']!=f"{b['id']}번"]
                    if moves:
                        new_alloc=[]; mid=max(bx['id'] for bx in st.session_state.bins)
                        for bx in st.session_state.bins:
                            for item in ([i for r in bx['rows'] for i in r['items']]+bx.get('stacked_items',[])):
                                tgt=bx['id']
                                for mp,mt in moves:
                                    if str(item['PKG NO'])==str(mp): tgt=mid+1 if "새" in mt else int(mt.replace("번",""))
                                new_alloc.append((item,tgt))
                        nd={}
                        for item,tid in new_alloc:
                            if tid not in nd: nd[tid]={'id':tid,'rows':[],'used_L':0,'total_W':0,'max_W':0,'max_H':0,'stacked_items':[],'groups':set()}
                            pack_items_into_bin([item],nd[tid],max_40_wt,max_40_len,max_hc_h)
                        st.session_state.bins=apply_labels(sorted(nd.values(),key=lambda x:x['id']),
                            max_20_len,max_20_wt,max_dry_h,max_hc_h,max_fr20_len,max_fr20_wt,max_fr40_len,max_fr40_wt)
                        st.rerun()

                with st.expander("👁️ 적재 단면도 및 제원 확인",expanded=False):
                    lbl=b['c_label']
                    if   '20ft Dry'       in lbl: cml,cmw,cmh=max_20_len,max_20_wt,max_dry_h
                    elif '40ft HC'        in lbl: cml,cmw,cmh=max_40_len,max_40_wt,max_hc_h
                    elif '40ft Dry'       in lbl: cml,cmw,cmh=max_40_len,max_40_wt,max_dry_h
                    elif '20ft Flat Rack' in lbl: cml,cmw,cmh=max_fr20_len,max_fr20_wt,max_fr_h
                    elif '40ft Flat Rack' in lbl: cml,cmw,cmh=max_fr40_len,max_fr40_wt,max_fr_h
                    else:                         cml,cmw,cmh=max_40_len,max_40_wt,max_dry_h
                    uw=max([r['used_W'] for r in b['rows']]+[0])
                    uh=b['max_H']+max([s['H'] for s in b.get('stacked_items',[])]+[0]) if b.get('stacked_items') else b['max_H']
                    cc1,cc2,cc3,cc4=st.columns(4)
                    cc1.markdown(f"**📏 길이:** {b['used_L']:,}/{cml:,}mm"); cc1.progress(min(1.0,b['used_L']/cml))
                    cc2.markdown(f"**⚖️ 중량:** {b['total_W']:,}/{cmw:,}kg"); cc2.progress(min(1.0,b['total_W']/cmw))
                    if uw>2340: cc3.markdown(f"**↔️ 폭:** {uw:,}/2340mm <span style='color:{ALERT_COLOR};font-weight:bold;'>[OW +{uw-2340:,}]</span>",unsafe_allow_html=True)
                    else: cc3.markdown(f"**↔️ 폭:** {uw:,}/2340mm"); cc3.progress(min(1.0,uw/2340))
                    if uh>cmh: cc4.markdown(f"**↕️ 높이:** {uh:,}/{cmh:,}mm <span style='color:{ALERT_COLOR};font-weight:bold;'>[OH +{uh-cmh:,}]</span>",unsafe_allow_html=True)
                    else: cc4.markdown(f"**↕️ 높이:** {uh:,}/{cmh:,}mm"); cc4.progress(min(1.0,uh/cmh))

                    st.markdown(f"<span style='background:{ALERT_COLOR};padding:2px 8px;border-radius:4px;color:white;font-size:12px;'>■ FR (OW/OH)</span>&nbsp;&nbsp;"
                        f"<span style='background:#e67e22;padding:2px 8px;border-radius:4px;color:white;font-size:12px;'>■ HC (H>{max_dry_h}mm)</span>&nbsp;&nbsp;"
                        f"<span style='background:{ACCENT_COLOR};padding:2px 8px;border-radius:4px;color:white;font-size:12px;'>■ DRY (H≤{max_dry_h}mm)</span>",
                        unsafe_allow_html=True)

                    from collections import Counter as _C
                    stk_cnt=_C((s['L'],s['W'],s['H']) for s in b.get('stacked_items',[]))
                    base_cnt=_C((i['L'],i['W'],i['H']) for r in b['rows'] for i in r['items'])
                    fig=go.Figure()
                    fig.add_shape(type="rect",x0=b['used_L'],y0=0,x1=cml,y1=2340,fillcolor="#e1e4e8",opacity=0.4,line_width=0)
                    fig.add_shape(type="rect",x0=0,y0=0,x1=cml,y1=2340,line=dict(color=MAIN_COLOR,width=2))
                    cx=0
                    for r in b['rows']:
                        cy=(2340-r['used_W'])/2
                        for item in r['items']:
                            dim=(item['L'],item['W'],item['H'])
                            layers=1+math.ceil(stk_cnt.get(dim,0)/max(1,base_cnt.get(dim,1)))
                            if item['W']>2340 or item['H']>max_hc_h: ic=ALERT_COLOR
                            elif item['H']>max_dry_h: ic="#e67e22"
                            else: ic=ACCENT_COLOR
                            border=dict(color="#FFD700",width=3) if layers>1 else dict(color="white",width=1)
                            fig.add_shape(type="rect",x0=cx,y0=cy,x1=cx+item['L'],y1=cy+item['W'],fillcolor=ic,opacity=0.85,line=border)
                            
                            if layers>1: lbl_txt=f"×{layers}단<br>{item['PKG NO']}<br>H{item['H']}"
                            else:        lbl_txt=f"{item['PKG NO']}<br>H{item['H']}"
                            fig.add_annotation(x=cx+item['L']/2,y=cy+item['W']/2,text=lbl_txt,showarrow=False,font=dict(color="white",size=9))
                            cy+=item['W']
                        cx+=r['max_L']
                    fig.add_shape(type="line",x0=b['used_L'],y0=-200,x1=b['used_L'],y1=2800,line=dict(color=ALERT_COLOR,width=2,dash="dash"))
                    if b['used_L']>100: fig.add_annotation(x=b['used_L']/2,y=2650,text=f"적재: {b['used_L']:,}mm",showarrow=False,font=dict(color=MAIN_COLOR,size=13,weight="bold"))
                    if cml-b['used_L']>100: fig.add_annotation(x=b['used_L']+(cml-b['used_L'])/2,y=2650,text=f"잔여: {cml-b['used_L']:,}mm",showarrow=False,font=dict(color=ALERT_COLOR,size=13,weight="bold"))
                    fig.update_layout(xaxis=dict(visible=False,range=[-200,max_40_len+400]),yaxis=dict(visible=False,range=[-300,3100]),
                        height=280,margin=dict(l=10,r=10,t=30,b=10),paper_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig,use_container_width=True,key=f"plot_{b['id']}")
                st.markdown('</div>',unsafe_allow_html=True)

            st.markdown("---")
            from collections import defaultdict, Counter as _Cx
            rcc=defaultdict(_Cx); rdl=defaultdict(list)
            for bx in bins:
                for item in ([i for r in bx['rows'] for i in r['items']]+bx.get('stacked_items',[])):
                    rcc[item['row_idx']][bx['c_label']]+=1
                    rdl[item['row_idx']].append((item['PKG NO'],bx['c_label'],item['L'],item['W'],item['H'],item['WEIGHT']))
            mapping={}
            for ri,ctr in rcc.items():
                total=sum(ctr.values())
                mapping[ri]=" / ".join(f"{lb} ({cn}개)" for lb,cn in ctr.items()) if total>1 else list(ctr.keys())[0]

            if file.name.endswith('.xlsx'):
                file.seek(0); wb=load_workbook(file); ws=wb[selected_sheet]
                tc=ws.max_column+1; tl=ws.cell(row=1,column=tc).column_letter
                if ws.max_row >= 4:
                    ws.cell(row=4,column=tc).value="배정 컨테이너"
                    if tc>1:
                        sc=ws.cell(row=4,column=tc-1); dc=ws.cell(row=4,column=tc)
                        dc.font=copy(sc.font); dc.fill=copy(sc.fill); dc.border=copy(sc.border); dc.alignment=copy(sc.alignment)
                
                for ri,label in mapping.items():
                    er=int(ri)+1; ws.cell(row=er,column=tc).value=label
                    if tc>1:
                        sc=ws.cell(row=er,column=tc-1); dc=ws.cell(row=er,column=tc)
                        dc.font=copy(sc.font); dc.fill=copy(sc.fill); dc.border=copy(sc.border); dc.alignment=copy(sc.alignment)
                ws.column_dimensions[tl].width=30
                box_det=[(p,lb,l,w,h,wt) for ri,dets in rdl.items() if len(dets)>1 for p,lb,l,w,h,wt in dets]
                if box_det:
                    if "배정상세" in wb.sheetnames: del wb["배정상세"]
                    wd=wb.create_sheet("배정상세")
                    wd.append(["PKG NO","L (mm)","W (mm)","H (mm)","WEIGHT (kg)","배정 컨테이너"])
                    wd.column_dimensions["A"].width=20; wd.column_dimensions["F"].width=30
                    for row in box_det: wd.append(list(row))
                out=io.BytesIO(); wb.save(out)
                st.download_button("📥 최종 결과 다운로드 (원본 양식 유지)",out.getvalue(),"CLP_RESULT_FINAL.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)

    except Exception as e: st.error(f"데이터 처리 중 오류 발생: {e}")
