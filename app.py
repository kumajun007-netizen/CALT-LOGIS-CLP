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

# PKG 번호에서 숫자만 추출하는 정렬용 함수
def extract_pkg_num(pkg_str):
    nums = re.findall(r'\d+', str(pkg_str))
    return int(nums[0]) if nums else 999999

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
        max_20_wt  = st.number_input("최대 중량 (20ft, kg)", 15000,40000,28250,key="i_20wt")
        max_20_len = st.number_input("최대 길이 (20ft, mm)", 5500,6500,5900,key="i_20len")
        max_40_wt  = st.number_input("최대 중량 (40ft, kg)", 20000,40000,29500,key="i_40wt")
        max_40_len = st.number_input("최대 길이 (40ft, mm)", 11000,13000,11900,key="i_40len")
        max_dry_h  = st.number_input("DRY 내부 높이 (mm)", 2000,3000,2370,key="i_dryh")
        max_hc_h   = st.number_input("HC 내부 높이 (mm)",  2000,3500,2670,key="i_hch")
        max_fr20_wt  = st.number_input("최대 중량 (20FR, kg)", 10000,40000,25000,key="i_fr20wt")
        max_fr20_len = st.number_input("최대 길이 (20FR, mm)", 4000,7000,5600,key="i_fr20len")
        max_fr40_wt  = st.number_input("최대 중량 (40FR, kg)", 10000,45000,30000,key="i_fr40wt")
        max_fr40_len = st.number_input("최대 길이 (40FR, mm)", 8000,14000,11600,key="i_fr40len")
        max_fr_h     = st.number_input("FR 내부 높이 (mm)", 1000,3000,2260,key="i_frh")

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
                        b['stacked_items'].append(piece); b['total_W'] += piece['WEIGHT']
                        b['groups'].add(piece['GROUP']); placed = True; continue
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
        if tags: b['c_label'] = f"{'20ft' if is_20fr else '40ft'} Flat Rack [{' + '.join(tags)}] #{b['id']}"
        else:
            base = "40ft HC" if b['max_H']>max_dry_h else ("20ft Dry" if is_20ft else "40ft Dry")
            b['c_label'] = f"{base} #{b['id']}"
    return bins

def calculate_expert_packing(df, max_40_wt, max_40_len, max_20_wt, max_20_len, max_dry_h, max_hc_h, load_mode):
    raw = []
    for _, row in df.iterrows():
        l,w,h,wt = int(row['L']+.5),int(row['W']+.5),int(row['H']+.5),int(row['WEIGHT']+.5)
        # PKG NO에서 숫자 추출하여 정렬 가중치 부여
        p_seq = extract_pkg_num(row['PKG NO'])
        raw.append({**row.to_dict(),'L':l,'W':w,'H':h,'WEIGHT':wt, 'p_seq': p_seq})

    all_pieces=[]
    for p in raw:
        rule = p.get('LOAD_KW', '') or load_mode
        if max(p['L'], p['W']) > 2340 or p['H'] > max_hc_h: rule = '4WAY'
        
        s, lg = min(p['L'], p['W']), max(p['L'], p['W'])
        if rule == 'FORK_L': el, ew = (s, lg) if lg <= 2340 else (lg, s)
        elif rule == 'FORK_W': el, ew = (lg, s) if s <= 2340 else (s, lg)
        else:
            if lg <= 2340: el, ew = (s, lg)
            else: el, ew = (lg, s)
        all_pieces.append({**p,'L':el,'W':ew})

    # 정렬 기준 최적화: 폭 -> 높이 -> 길이 -> PKG NO 숫자순
    sk=lambda x:(-x['W'],-x['H'],-x['L'],x['p_seq'])
    fr_p  = sorted([p for p in all_pieces if p['W']>2340 or p['H']>max_hc_h], key=sk)
    hc_p  = sorted([p for p in all_pieces if p['W']<=2340 and max_dry_h<p['H']<=max_hc_h], key=sk)
    dry_p = sorted([p for p in all_pieces if p['W']<=2340 and p['H']<=max_dry_h], key=sk)

    def _pack(pieces, c_no, mwt, mlen):
        if not pieces: return [],c_no
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

    fr_bins,c = _pack(fr_p,1,max_fr40_wt,max_fr40_len)
    rem_hc = _fill(fr_bins,hc_p,max_fr40_wt,max_fr40_len)
    rem_dry = _fill(fr_bins,dry_p,max_fr40_wt,max_fr40_len)
    hc_bins,c = _pack(rem_hc,c,max_40_wt,max_40_len)
    rem_dry = _fill(hc_bins,rem_dry,max_40_wt,max_40_len)
    dry_bins,c = _pack(rem_dry,c,max_40_wt,max_40_len)

    res = fr_bins + hc_bins + dry_bins
    for i,b in enumerate(res): b['id']=i+1
    return apply_labels(res,max_20_len,max_20_wt,max_dry_h,max_hc_h,max_fr20_len,max_fr20_wt,max_fr40_len,max_fr40_wt)

st.markdown("### 📤 기본 설정 및 패킹리스트 업로드")
load_mode_ui = st.radio("👉 **지게차(포크) 기본 진입 방향 설정**", ["FORK_L", "FORK_W", "4WAY"], horizontal=True, on_change=reset_data)

file = st.file_uploader("패킹리스트 파일을 업로드하세요 (xlsx, csv)", type=['csv','xlsx'], on_change=reset_data)

if file is not None:
    try:
        COL_PKG=1;COL_ITEM=3;COL_DESC=4;COL_QTY=5;COL_WEIGHT=8;COL_L=9;COL_W=11;COL_H=13;COL_REMARK=14;COL_LOAD=15
        raw_full=None; selected_sheet=None

        if file.name.lower().endswith('.xlsx'):
            file.seek(0); sheets=pd.read_excel(file,sheet_name=None,header=None)
            best=0
            for sn,sd in sheets.items():
                if sn=='사용설명서': continue
                if sd.shape[1] <= COL_H: continue
                cnt=sum(1 for _,r in sd.iterrows() if str(r.iloc[COL_PKG]).strip().lower() not in ['','nan','none','.'] and clean_num(r.iloc[COL_L])>0)
                if cnt>best: best=cnt; raw_full=sd; selected_sheet=sn
        else:
            file.seek(0); raw_full=pd.read_csv(file,header=None)

        p_data=[]
        for oi in range(len(raw_full)):
            row=raw_full.iloc[oi]
            if len(row) <= COL_H: continue
            pkg=str(row.iloc[COL_PKG]).replace('00:00:00','').replace('.0','').strip()
            lv, wv, hv = clean_num(row.iloc[COL_L]), clean_num(row.iloc[COL_W]), clean_num(row.iloc[COL_H])
            if pkg.lower() in ['','nan','none','.'] or lv==0: continue
            
            qty=int(clean_num(row.iloc[COL_QTY])) if clean_num(row.iloc[COL_QTY])>0 else 1
            wtv=clean_num(row.iloc[COL_WEIGHT])
            rem=str(row.iloc[COL_REMARK]).upper()
            load_val=str(row.iloc[COL_LOAD]).upper()
            
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
                with st.container():
                    st.markdown(f'<div class="container-box"><h3>📦 {b["c_label"]}</h3>', unsafe_allow_html=True)
                    
                    # 시각화 로직
                    fig = go.Figure()
                    cml = max_40_len if '40ft' in b['c_label'] else max_20_len
                    fig.add_shape(type="rect", x0=0, y0=0, x1=cml, y1=2340, line=dict(color=MAIN_COLOR, width=2))
                    
                    cx = 0
                    for r in b['rows']:
                        cy = (2340 - r['used_W']) / 2
                        for item in r['items']:
                            ic = ALERT_COLOR if item['W'] > 2340 or item['H'] > max_hc_h else (ACCENT_COLOR if item['H'] <= max_dry_h else "#e67e22")
                            fig.add_shape(type="rect", x0=cx, y0=cy, x1=cx+item['L'], y1=cy+item['W'], fillcolor=ic, line=dict(color="white", width=1))
                            fig.add_annotation(x=cx+item['L']/2, y=cy+item['W']/2, text=f"{item['PKG NO']}", showarrow=False, font=dict(color="white", size=9))
                            cy += item['W']
                        cx += r['max_L']
                    
                    fig.update_layout(height=300, margin=dict(l=10,r=10,t=10,b=10), xaxis=dict(visible=False), yaxis=dict(visible=False))
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # 상세 표
                    t_list = []
                    for r in b['rows']: 
                        for i in r['items']: t_list.append({**i, "위치": "바닥"})
                    for s in b.get('stacked_items', []): t_list.append({**s, "위치": "단적"})
                    st.dataframe(pd.DataFrame(t_list)[['PKG NO','ITEM','L','W','H','WEIGHT','위치']], use_container_width=True)
                    st.markdown('</div>', unsafe_allow_html=True)

            # 결과 다운로드
            if selected_sheet:
                file.seek(0); wb=load_workbook(file); ws=wb[selected_sheet]
                tc = ws.max_column + 1
                ws.cell(row=1, column=tc).value = "배정 결과"
                
                # 매핑 데이터 생성
                mapping = {}
                for bx in bins:
                    for it in ([i for r in bx['rows'] for i in r['items']] + bx.get('stacked_items',[])):
                        rid = it['row_idx'] + 1
                        mapping[rid] = mapping.get(rid, []) + [bx['c_label']]
                
                for rid, lbls in mapping.items():
                    ws.cell(row=rid, column=tc).value = ", ".join(set(lbls))
                
                out=io.BytesIO(); wb.save(out)
                st.download_button("📥 최종 결과 엑셀 다운로드", out.getvalue(), "CLP_RESULT.xlsx", use_container_width=True)

    except Exception as e:
        st.error(f"오류 발생: {e}")
