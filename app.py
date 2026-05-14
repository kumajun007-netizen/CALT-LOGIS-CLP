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

/* 커스텀 KPI 카드 디자인 */
.kpi-card {{ background-color:white;padding:15px;border-radius:10px;border-left:5px solid {MAIN_COLOR};box-shadow:0 2px 4px rgba(0,0,0,0.05); height:100%; }}
.kpi-title {{ font-size:14px; color:#555; margin-bottom:5px; font-weight:600; }}
.kpi-value {{ font-size:28px; font-weight:800; color:#111; }}
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


# ★★★ 안정적인 2D + 혼적 테트리스 (버그 원천 차단형) ★★★
def pack_items_into_bin(pieces, b, max_wt, max_len, max_h=2670):
    placed_any = False
    for piece in pieces:
        placed = False
        
        # 0. 컨테이너 중량 초과 시 스킵
        if b['total_W'] + piece['WEIGHT'] > max_wt:
            continue
            
        # 1. 다단 적재(혼합 적재) 시도
        for r in b['rows']:
            for base_item in r['items']:
                if '_stacked' not in base_item:
                    base_item['_stacked'] = []
                
                # 맨 위 화물을 기준으로 면적 검사
                top_item = base_item['_stacked'][-1] if base_item['_stacked'] else base_item
                
                # 얹을 화물의 크기가 같거나 작을 때만 쌓음
                if piece['L'] <= top_item['L'] and piece['W'] <= top_item['W']:
                    current_h = base_item['H'] + sum(s['H'] for s in base_item['_stacked'])
                    current_layers = 1 + len(base_item['_stacked'])
                    limit_stk = min(base_item.get('MAX_STK', 1), piece.get('MAX_STK', 1))
                    
                    if current_h + piece['H'] <= max_h and current_layers < limit_stk:
                        new_piece = piece.copy()
                        new_piece['_stacked'] = []
                        base_item['_stacked'].append(new_piece)
                        if 'stacked_items' not in b: b['stacked_items'] = []
                        b['stacked_items'].append(new_piece)
                        b['total_W'] += new_piece['WEIGHT']
                        b['groups'].add(new_piece['GROUP'])
                        placed = True
                        placed_any = True
                        break
            if placed: break
                
        if placed: continue

        # 2. 바닥에 빈 공간 찾아서 깔기
        row_found = False
        for r in b['rows']:
            tL = max(r['max_L'], piece['L'])
            if r['used_W'] + piece['W'] <= 2340 and b['used_L'] + (tL - r['max_L']) <= max_len:
                new_piece = piece.copy()
                new_piece['_stacked'] = []
                r['items'].append(new_piece)
                r['used_W'] += new_piece['W']
                b['used_L'] += (tL - r['max_L'])
                r['max_L'] = tL
                b['total_W'] += new_piece['WEIGHT']
                b['max_W'] = max(b['max_W'], new_piece['W'])
                b['max_H'] = max(b['max_H'], new_piece['H'])
                b['groups'].add(new_piece['GROUP'])
                row_found = True
                placed = True
                placed_any = True
                break
        
        if not row_found:
            if b['used_L'] + piece['L'] <= max_len:
                new_piece = piece.copy()
                new_piece['_stacked'] = []
                b['rows'].append({'items':[new_piece], 'used_W':new_piece['W'], 'max_L':new_piece['L']})
                b['used_L'] += new_piece['L']
                b['total_W'] += new_piece['WEIGHT']
                b['max_W'] = max(b['max_W'], new_piece['W'])
                b['max_H'] = max(b['max_H'], new_piece['H'])
                b['groups'].add(new_piece['GROUP'])
                placed = True
                placed_any = True

    return placed_any


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
        nums = re.findall(r'\d+', str(row['PKG NO']))
        p_seq = int("".join(nums)) if nums else 999999
        raw.append({**row.to_dict(),'L':l,'W':w,'H':h,'WEIGHT':wt, 'p_seq': p_seq})

    sg={}
    for p in raw:
        rule = p.get('LOAD_KW', '')
        if rule not in ['FORK_L', 'FORK_W', '4WAY']:
            rule = load_mode 
        is_fr_size = max(p['L'], p['W']) > 2340 or p['H'] > max_hc_h
        if is_fr_size: rule = '4WAY' 
        p['eff_rule'] = rule
        key=(min(p['L'],p['W']),max(p['L'],p['W']),p['H']>max_dry_h, rule)
        sg.setdefault(key,[]).append(p)
        
    all_pieces=[]
    for (s,lg,_,rule),items in sg.items():
        n=len(items); ca=lg<=2340; cb=s<=2340
        
        # FORK_L: 포크 구멍이 있는 긴 쪽(lg)을 W방향으로 눕힘
        if rule == 'FORK_L': 
            el, ew = (s, lg) if lg <= 2340 else (lg, s)
        elif rule == 'FORK_W': 
            el, ew = (lg, s) if s <= 2340 else (s, lg)
        else: 
            if ca and cb:
                sa=max(1,int(2340//lg)); sb=max(1,int(2340//s))
                La=math.ceil(n/sa)*s; Lb=math.ceil(n/sb)*lg
                el,ew=(lg,s) if (Lb<La and sb>=2) else (s,lg)
            elif cb: el,ew=lg,s
            else: el,ew=s,lg
            
        for p in items: all_pieces.append({**p,'L':el,'W':ew})

    sk=lambda x:(-x['W'],-x['H'],-x['L'],x['p_seq'],x['GROUP'])
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
                if pack_items_into_bin([piece], b, wt, ln, max_hc_h):
                    placed=True
                    break
            if not placed:
                nb={'id':c,'rows':[],'used_L':0,'total_W':0,'max_W':0,'max_H':0,'stacked_items':[],'groups':set()}
                pack_items_into_bin([piece], nb, wt, ln, max_hc_h)
                bins.append(nb)
                c+=1
        return [b for b in bins if b['used_L']>0 or b.get('stacked_items')], c

    def _fill(bins, cands, mwt=None, mlen=None):
        wt=mwt or max_40_wt; ln=mlen or max_40_len
        rem=list(cands)
        for b in sorted(bins,key=lambda x:x['used_L']):
            for piece in list(rem):
                if pack_items_into_bin([piece], b, wt, ln, max_hc_h):
                    rem.remove(piece)
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


st.markdown("### 📤 기본 설정 및 패킹리스트 업로드")

load_mode_ui = st.radio(
    "👉 **지게차(포크) 기본 진입 방향 설정** (※ P열 개별 지정 우선)", 
    ["FORK_L", "FORK_W", "4WAY"], horizontal=True, on_change=reset_data
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

            count_items = lambda bx: sum(len(r['items']) for r in bx['rows']) + len(bx.get('stacked_items',[]))

            # --- 통일성 있게 개편된 KPI 섹션 ---
            st.subheader("📊 실시간 적재 요약")
            c1, c2, c3, c4 = st.columns(4)
            
            packed = sum(count_items(b) for b in bins)
            total_w = sum(b['total_W'] for b in bins)
            raw_total_w = sum(df['WEIGHT']) 
            
            from collections import Counter as _C
            c_types = [re.sub(r' #\d+$', '', b['c_label']) for b in bins]
            type_counts = _C(c_types)
            
            fr_counts = {k: v for k, v in type_counts.items() if 'Flat Rack' in k or 'FR' in k}
            dry_counts = {k: v for k, v in type_counts.items() if 'Dry' in k or 'HC' in k}
            
            fr_total = sum(fr_counts.values())
            dry_total = sum(dry_counts.values())

            fr_html = "".join([f"<div style='font-size:13px; color:#444; margin-top:4px;'>· {k} <b style='color:{ACCENT_COLOR}; font-size:15px;'>{v}</b> 대</div>" for k, v in fr_counts.items()])
            if not fr_html: fr_html = "<div style='font-size:13px; color:#999; margin-top:4px;'>배정된 FR 컨테이너 없음</div>"

            dry_html = "".join([f"<div style='font-size:13px; color:#444; margin-top:4px;'>· {k} <b style='color:{ACCENT_COLOR}; font-size:15px;'>{v}</b> 대</div>" for k, v in dry_counts.items()])
            if not dry_html: dry_html = "<div style='font-size:13px; color:#999; margin-top:4px;'>배정된 DRY/HC 컨테이너 없음</div>"
            
            c1.markdown(f'<div class="kpi-card"><div class="kpi-title">배정 화물 / 전체 화물</div><div class="kpi-value"><span style="color:{ACCENT_COLOR};">{packed}</span> / {len(df)} <span style="font-size:16px;color:#777;font-weight:600;">PKG</span></div></div>', unsafe_allow_html=True)
            c2.markdown(f'<div class="kpi-card"><div class="kpi-title">FR 컨테이너 ({fr_total} UNIT)</div><div style="margin-top:2px;">{fr_html}</div></div>', unsafe_allow_html=True)
            c3.markdown(f'<div class="kpi-card"><div class="kpi-title">DRY 컨테이너 ({dry_total} UNIT)</div><div style="margin-top:2px;">{dry_html}</div></div>', unsafe_allow_html=True)
            c4.markdown(f'<div class="kpi-card"><div class="kpi-title">배정 중량 / 전체 중량</div><div class="kpi-value"><span style="color:{ACCENT_COLOR};">{total_w:,.0f}</span> / {raw_total_w:,.0f} <span style="font-size:16px;color:#777;font-weight:600;">kg</span></div></div>', unsafe_allow_html=True)

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
                    sp=TYPE_SPECS[new_type]
                    all_it=[i for r in b['rows'] for i in r['items']]+b.get('stacked_items',[])
                    all_it = sorted(all_it, key=lambda x: (-x['W'], -x['H'], -x['L'], x['p_seq'], x['GROUP']))
                    nb={'id':b['id'],'rows':[],'used_L':0,'total_W':0,'max_W':0,'max_H':0,'stacked_items':[],'groups':set()}
                    ov=[]
                    for it in all_it:
                        if not pack_items_into_bin([it], nb, sp['mw'], sp['ml'], sp['mh']):
                            ov.append(it)
                    
                    h_ex=[i for i in ([x for r in nb['rows'] for x in r['items']]+nb.get('stacked_items',[])) if i['H']>sp['mh']]
                    nb['c_label']=f"{sp['lb']} #{b['id']}"; nb['forced_label']=True
                    
                    upd=[]
                    for bx in st.session_state.bins:
                        upd.append(nb if bx['id']==b['id'] else bx)
                    
                    if ov:
                        st.toast(f"⚠️ {len(ov)}개 화물 제원 초과 → 새 컨테이너로 분리됐습니다.")
                        ov = sorted(ov, key=lambda x: (-x['W'], -x['H'], -x['L'], x['p_seq'], x['GROUP']))
                        mid=max(bx['id'] for bx in upd)
                        curr_ob = None
                        abs_wt = max(max_40_wt, max_fr40_wt)
                        abs_len = max(max_40_len, max_fr40_len)
                        abs_h = max(max_hc_h, max_fr_h)
                        
                        for it in ov:
                            placed = False
                            if curr_ob:
                                if pack_items_into_bin([it], curr_ob, abs_wt, abs_len, abs_h):
                                    placed = True
                            if not placed:
                                if curr_ob and count_items(curr_ob) > 0: upd.append(curr_ob)
                                mid += 1
                                curr_ob = {'id':mid,'rows':[],'used_L':0,'total_W':0,'max_W':0,'max_H':0,'stacked_items':[],'groups':set()}
                                pack_items_into_bin([it], curr_ob, abs_wt, abs_len, abs_h)
                        if curr_ob and count_items(curr_ob) > 0: upd.append(curr_ob)
                            
                    if h_ex: st.toast(f"🚨 {len(h_ex)}개 화물 H 초과! CLP 작성 시 확인하세요.")
                    
                    st.session_state.bins=apply_labels(sorted(upd,key=lambda x:x['id']),
                        max_20_len,max_20_wt,max_dry_h,max_hc_h,max_fr20_len,max_fr20_wt,max_fr40_len,max_fr40_wt)
                    
                    for idx, bx in enumerate(st.session_state.bins):
                        if bx.get('forced_label'): bx['c_label'] = re.sub(r'#\d+', f"#{idx+1}", bx['c_label'])
                        bx['id'] = idx + 1
                    st.rerun()

                base_items=[item for r in b['rows'] for item in r['items']]
                t_data=[]
                for r in b['rows']:
                    for item in r['items']: t_data.append({**item,'위치':'바닥','이동':f"{b['id']}번",'⚠️':''})
                for s in b.get('stacked_items',[]):
                    t_data.append({**s,'위치':'단적/혼적','이동':f"{b['id']}번",'⚠️':''})
                    
                df_edit=pd.DataFrame(t_data)[['⚠️','위치','PKG NO','ITEM','L','W','H','WEIGHT','이동']]
                edited_df=st.data_editor(df_edit,hide_index=True,use_container_width=True,key=f"ed_{b['id']}",
                    column_config={"이동":st.column_config.SelectboxColumn("🚚 이동",options=topts)},
                    disabled=['⚠️','위치','PKG NO','ITEM','L','W','H','WEIGHT'])

                if st.button(f"🚀 {b['id']}번 변경사항 적용",key=f"btn_{b['id']}"):
                    moves=[(r['PKG NO'],r['이동']) for _,r in edited_df.iterrows() if r['이동']!=f"{b['id']}번"]
                    if moves:
                        from collections import defaultdict
                        move_dict = dict(moves)
                        source_bin_id = b['id']
                        target_bin_ids = set()
                        for mt in move_dict.values():
                            if "새" not in mt: target_bin_ids.add(int(mt.replace("번","")))
                        modified_ids = target_bin_ids | {source_bin_id}

                        new_bins = []
                        items_to_repack = defaultdict(list)
                        
                        for bx in st.session_state.bins:
                            if bx['id'] not in modified_ids:
                                new_bins.append(bx)
                            else:
                                items_in_bx = [i for r in bx['rows'] for i in r['items']] + bx.get('stacked_items', [])
                                for it in items_in_bx:
                                    pkg = str(it['PKG NO'])
                                    if bx['id'] == source_bin_id and pkg in move_dict:
                                        mt = move_dict[pkg]
                                        if "새" in mt: items_to_repack['NEW'].append(it)
                                        else: items_to_repack[int(mt.replace("번",""))].append(it)
                                    else:
                                        items_to_repack[bx['id']].append(it)

                        repacked_bins = []
                        overflow_items = []
                        
                        def get_limits(tid):
                            old_bx = next((x for x in st.session_state.bins if x['id'] == tid), None)
                            if old_bx:
                                lbl = old_bx['c_label']
                                if '20ft Dry' in lbl: return max_20_wt, max_20_len, max_dry_h
                                elif '40ft HC' in lbl: return max_40_wt, max_40_len, max_hc_h
                                elif '40ft Dry' in lbl: return max_40_wt, max_40_len, max_dry_h
                                elif '20ft Flat Rack' in lbl: return max_fr20_wt, max_fr20_len, max_fr_h
                                elif '40ft Flat Rack' in lbl: return max_fr40_wt, max_fr40_len, max_fr_h
                            return max_40_wt, max_40_len, max_hc_h

                        for tid in modified_ids:
                            if tid not in items_to_repack or not items_to_repack[tid]: continue
                            cmw, cml, cmh = get_limits(tid)
                            items = sorted(items_to_repack[tid], key=lambda x: (-x['W'], -x['H'], -x['L'], x['p_seq'], x['GROUP']))
                            nb = {'id': tid, 'rows': [], 'used_L': 0, 'total_W': 0, 'max_W': 0, 'max_H': 0, 'stacked_items': [], 'groups': set()}
                            for it in items:
                                if not pack_items_into_bin([it], nb, cmw, cml, cmh):
                                    overflow_items.append(it)
                            if count_items(nb) > 0: repacked_bins.append(nb)

                        new_items = items_to_repack.get('NEW', []) + overflow_items
                        if new_items:
                            if overflow_items: st.toast("⚠️ 공간 부족으로 넘친 화물을 새 컨테이너로 보호합니다.", icon="🚨")
                            new_items = sorted(new_items, key=lambda x: (-x['W'], -x['H'], -x['L'], x['p_seq'], x['GROUP']))
                            abs_wt = max(max_40_wt, max_fr40_wt)
                            abs_len = max(max_40_len, max_fr40_len)
                            abs_h = max(max_hc_h, max_fr_h)
                            mid = max([x['id'] for x in st.session_state.bins] + [x['id'] for x in repacked_bins] + [0])
                            curr_ob = None
                            
                            for it in new_items:
                                placed = False
                                if curr_ob:
                                    if pack_items_into_bin([it], curr_ob, abs_wt, abs_len, abs_h): placed = True
                                if not placed:
                                    if curr_ob and count_items(curr_ob) > 0: repacked_bins.append(curr_ob)
                                    mid += 1
                                    curr_ob = {'id': mid, 'rows': [], 'used_L': 0, 'total_W': 0, 'max_W': 0, 'max_H': 0, 'stacked_items': [], 'groups': set()}
                                    pack_items_into_bin([it], curr_ob, abs_wt, abs_len, abs_h)
                            if curr_ob and count_items(curr_ob) > 0: repacked_bins.append(curr_ob)

                        repacked_labeled = apply_labels(repacked_bins, max_20_len, max_20_wt, max_dry_h, max_hc_h, max_fr20_len, max_fr20_wt, max_fr40_len, max_fr40_wt)
                        final_bins = sorted(new_bins + repacked_labeled, key=lambda x: x['id'])
                        for idx, fb in enumerate(final_bins): fb['id'] = idx + 1
                        
                        st.session_state.bins = final_bins
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

                    fig=go.Figure()
                    fig.add_shape(type="rect",x0=b['used_L'],y0=0,x1=cml,y1=2340,fillcolor="#e1e4e8",opacity=0.4,line_width=0)
                    fig.add_shape(type="rect",x0=0,y0=0,x1=cml,y1=2340,line=dict(color=MAIN_COLOR,width=2))
                    cx=0
                    for r in b['rows']:
                        cy=(2340-r['used_W'])/2
                        for item in r['items']:
                            stacked_list = item.get('_stacked', [])
                            layers = 1 + len(stacked_list)
                            
                            if item['W']>2340 or item['H']>max_hc_h: ic=ALERT_COLOR
                            elif item['H']>max_dry_h: ic="#e67e22"
                            else: ic=ACCENT_COLOR
                            border=dict(color="#FFD700",width=3) if layers>1 else dict(color="white",width=1)
                            fig.add_shape(type="rect",x0=cx,y0=cy,x1=cx+item['L'],y1=cy+item['W'],fillcolor=ic,opacity=0.85,line=border)
                            
                            # 2D 화면에서 혼합적재(다른 화물이 섞여 올라간 경우) 내역 텍스트로 표시
                            if layers > 1: 
                                stack_txt = "<br>+".join([f"{s['PKG NO']}(H{s['H']})" for s in stacked_list])
                                lbl_txt = f"<b>{item['PKG NO']}</b><br>+ {stack_txt}"
                            else:        
                                lbl_txt=f"{item['PKG NO']}<br>H{item['H']}"
                                
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
                    rdl[item['row_idx']].append((item['PKG NO'], item['L'], item['W'], item['H'], item['WEIGHT'], bx['c_label']))
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
                
                box_det=[(p, l, w, h, wt, lb) for ri,dets in rdl.items() if len(dets)>1 for p, l, w, h, wt, lb in dets]
                
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
