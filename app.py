import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import math
import os
import io

# --- 1. 화면 스타일 및 테마 설정 ---
st.set_page_config(page_title="CALT-LOGIS CLP System", layout="wide")

# 칼트로지스 공식 컬러 가이드 적용 (Deep Navy & Soft Gray)
MAIN_COLOR = "#001f3f"
SUB_COLOR = "#f0f2f6"
ACCENT_COLOR = "#007bff"

st.markdown(f"""
    <style>
    /* 전체 배경 및 폰트 */
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    * {{ font-family: 'Pretendard', sans-serif; }}
    
    .main {{ background-color: #f8f9fa; }}
    
    /* 상단 타이틀 바 (가운데 정렬로 수정) */
    .header-container {{
        background-color: {MAIN_COLOR};
        padding: 25px;
        border-radius: 10px;
        color: white;
        margin-bottom: 25px;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }}
    
    /* 카드 스타일 섹션 */
    .st-emotion-cache-1r6slb0 {{ border-radius: 10px; border: 1px solid #e1e4e8; background-color: white; padding: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }}
    
    /* 메트릭 카드 */
    [data-testid="stMetric"] {{
        background-color: white;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid {MAIN_COLOR};
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }}

    /* 버튼 스타일 커스텀 */
    .stButton>button {{
        width: 100%;
        font-weight: 600;
        color: white;
        background-color: {MAIN_COLOR};
        border-radius: 8px;
        border: none;
        padding: 0.5rem 1rem;
        transition: 0.3s;
    }}
    .stButton>button:hover {{
        background-color: {ACCENT_COLOR};
        color: white;
    }}
    
    /* 표(Table) 가독성 향상 */
    div[data-testid="stTable"] {{ font-size: 12px !important; }}
    th {{ background-color: {SUB_COLOR} !important; color: {MAIN_COLOR}; }}
    </style>
    """, unsafe_allow_html=True)

# --- 상단 헤더 영역 (로고 삭제 및 타이틀 가운데 정렬) ---
with st.container():
    st.markdown(f"""
        <div class="header-container">
            <div style="font-size: 30px; font-weight: 800; letter-spacing: -1px;">CALT-LOGIS CLP SYSTEM</div>
            <div style="font-size: 15px; opacity: 0.8; margin-top: 5px;">Busan New Port Center | Logistics Management</div>
        </div>
    """, unsafe_allow_html=True)

def clean_num(val):
    try:
        if pd.isna(val) or str(val).strip() in ['', '.', 'X', 'x']: return None 
        return float(str(val).replace(',', '').strip())
    except: return None

# --- 2. 사이드바: 설정 영역 (로고 유지) ---
with st.sidebar:
    # 요청하신 사이드바 로고는 그대로 유지됩니다.
    logo_path_sidebar = "칼트로지스로고.png"
    if os.path.exists(logo_path_sidebar):
        st.image(logo_path_sidebar, use_column_width=True)

    st.header("⚙️ 배정 옵션 설정")
    
    with st.expander("⚖️ 컨테이너 제원", expanded=True):
        max_40_wt = st.number_input("40ft 중량 (kg)", 20000, 40000, 29500)
        max_40_len = st.number_input("40ft 길이 (mm)", 11000, 13000, 12034)
        st.markdown("---")
        max_20_wt = st.number_input("20ft 중량 (kg)", 15000, 40000, 28250)
        max_20_len = st.number_input("20ft 길이 (mm)", 5500, 6500, 5899)
        st.markdown("---")
        max_dry_h = st.number_input("DRY 높이 (mm)", 2000, 3000, 2390)
        max_hc_h = st.number_input("HC 높이 (mm)", 2000, 3500, 2695)

    with st.expander("🛠 적재 로직", expanded=True):
        use_balancing = st.checkbox("⚖️ 균분적재 (Balancing)", value=True)
        allow_stacking = st.checkbox("🏢 다단적재 (Stacking)", value=False)
    
    if st.button("🔄 AI 재계산 실행"):
        st.session_state['manual_mode'] = False
    
    st.markdown("---")
    st.info("파일 수정 후 이 버튼을 눌러 설정을 적용하세요.")

# --- 짐 배치 핵심 로직 ---
def pack_items_into_bin(pieces, b, max_40_wt, max_40_len):
    for piece in pieces:
        placed = False
        if allow_stacking and piece['STACK_OK'] and piece['WEIGHT'] <= 1000 and b['total_W'] + piece['WEIGHT'] <= max_40_wt:
            if 'stacked_items' not in b: b['stacked_items'] = []
            b['stacked_items'].append(piece); b['total_W'] += piece['WEIGHT']; b['groups'].add(piece['GROUP']); placed = True
            continue
        row_found = False
        for r in b['rows']:
            temp_max_L = max(r['max_L'], piece['L'])
            if r['used_W'] + piece['W'] <= 2350 and b['used_L'] + (temp_max_L - r['max_L']) <= max_40_len and b['total_W'] + piece['WEIGHT'] <= max_40_wt:
                r['items'].append(piece); r['used_W'] += piece['W']; b['used_L'] += (temp_max_L - r['max_L'])
                r['max_L'] = temp_max_L; b['total_W'] += piece['WEIGHT']
                b['max_W'] = max(b['max_W'], piece['W']); b['max_H'] = max(b['max_H'], piece['H'])
                b['groups'].add(piece['GROUP']); row_found = True; placed = True; break
        if row_found: continue
        if b['used_L'] + piece['L'] <= max_40_len and b['total_W'] + piece['WEIGHT'] <= max_40_wt:
            b['rows'].append({'items': [piece], 'used_W': piece['W'], 'max_L': piece['L']})
            b['used_L'] += piece['L']; b['total_W'] += piece['WEIGHT']
            b['max_W'] = max(b['max_W'], piece['W']); b['max_H'] = max(b['max_H'], piece['H'])
            b['groups'].add(piece['GROUP']); placed = True

def apply_labels(bins, max_20_len, max_20_wt, fr_max_len, max_dry_h):
    for b in bins:
        is_20ft_size = b['used_L'] <= max_20_len and b['total_W'] <= max_20_wt
        is_ow, is_oh, is_ol = b['max_W'] > 2300, b['max_H'] > 1900, b['used_L'] > fr_max_len
        is_fv = is_ol or (is_ow and is_oh)
        tags = []
        if is_fv: tags.append("FV")
        else:
            if is_oh: tags.append("OH")
            if is_ow: tags.append("OW")
            if is_ol: tags.append("OL")
        if b['type'] == "FR":
            base = "20ft Flat Rack" if is_20ft_size else "40ft Flat Rack"
            b['c_label'] = f"{base} [{' + '.join(tags)}] #{b['id']}" if tags else f"{base} #{b['id']}"
        else:
            if b['max_H'] > max_dry_h: base = "40ft HC"
            else: base = "20ft Dry" if is_20ft_size else "40ft Dry"
            b['c_label'] = f"{base} #{b['id']}"
    return bins

def calculate_expert_packing(df, max_40_wt, max_40_len, max_20_wt, max_20_len, max_dry_h, max_hc_h, use_balancing):
    all_pieces = []
    fr_max_len = max_40_len - 430
    for _, row in df.iterrows():
        l, w, h, weight = int(row['L'] + 0.5), int(row['W'] + 0.5), int(row['H'] + 0.5), int(row['WEIGHT'] + 0.5)
        if max(l, w) <= 2350: el, ew = min(l, w), max(l, w)
        else: el, ew = max(l, w), min(l, w)
        req = "FR" if ew > 2350 or el > fr_max_len or h > max_hc_h else "DRY"
        all_pieces.append({**row, 'L': el, 'W': ew, 'H': h, 'WEIGHT': weight, 'REQ_TYPE': req})
    all_pieces.sort(key=lambda x: (x['REQ_TYPE'], x['GROUP'], x['PKG NO']))
    bins = []; c_no = 1
    req_types = sorted(list(set(p['REQ_TYPE'] for p in all_pieces)))
    for r_type in req_types:
        r_pieces = [p for p in all_pieces if p['REQ_TYPE'] == r_type]
        type_bins = []
        if use_balancing:
            total_l = sum(p['L'] for p in r_pieces)
            total_w = sum(p['WEIGHT'] for p in r_pieces)
            est_bins = max(1, math.ceil(total_l / max_40_len), math.ceil(total_w / max_40_wt))
            for _ in range(est_bins):
                type_bins.append({'id': c_no, 'type': r_type, 'rows': [], 'used_L': 0, 'total_W': 0, 'max_W': 0, 'max_H': 0, 'stacked_items': [], 'groups': set()})
                c_no += 1
        for piece in r_pieces:
            placed = False
            if use_balancing: type_bins.sort(key=lambda b: (0 if piece['GROUP'] in b['groups'] else 1, b['used_L']))
            for b in (type_bins if use_balancing else bins):
                if not use_balancing and b['type'] != piece['REQ_TYPE']: continue
                pack_items_into_bin([piece], b, max_40_wt, max_40_len)
                if piece in b.get('stacked_items', []) or any(piece in r['items'] for r in b['rows']): placed = True; break
            if not placed:
                new_bin = {'id': c_no, 'type': piece['REQ_TYPE'], 'rows': [], 'used_L': 0, 'total_W': 0, 'max_W': 0, 'max_H': 0, 'stacked_items': [], 'groups': set()}
                pack_items_into_bin([piece], new_bin, max_40_wt, max_40_len)
                if use_balancing: type_bins.append(new_bin)
                else: bins.append(new_bin)
                c_no += 1
        if use_balancing: bins.extend([b for b in type_bins if b['used_L'] > 0])
    bins.sort(key=lambda x: x['id'])
    return apply_labels(bins, max_20_len, max_20_wt, fr_max_len, max_dry_h)

# --- 3. 메인 화면 로직 ---
file = st.file_uploader("📂 패킹리스트 엑셀 파일을 업로드하세요", type=['csv', 'xlsx'])

if file is not None:
    try:
        raw_full = pd.read_excel(file, header=None) if file.name.endswith('.xlsx') else pd.read_csv(file, header=None)
        raw_process = pd.read_excel(file, skiprows=4, header=None) if file.name.endswith('.xlsx') else pd.read_csv(file, skiprows=4, header=None)
        
        p_data = []
        for i in range(len(raw_process)):
            row = raw_process.iloc[i]; s_ok = row.astype(str).str.contains('단적허용').any() if allow_stacking else False
            pkg_v = str(row[0]).replace('00:00:00','').replace('.0','').strip()
            if pkg_v in ['nan', '.', '']: continue
            p_data.append({'PKG NO': pkg_v, 'GROUP': str(row[4]), 'ITEM': str(row[5]), 'L': clean_num(row[10]), 'W': clean_num(row[12]), 'H': clean_num(row[14]), 'WEIGHT': clean_num(row[9]), 'STACK_OK': s_ok, 'row_idx': i+4})
        df = pd.DataFrame(p_data).dropna(subset=['L', 'W', 'H', 'WEIGHT'])
        df = df[~df['ITEM'].str.upper().str.contains('TOTAL', na=False)]

        if 'bins' not in st.session_state or not st.session_state.get('manual_mode', False):
            st.session_state.bins = calculate_expert_packing(df, max_40_wt, max_40_len, max_20_wt, max_20_len, max_dry_h, max_hc_h, use_balancing)
            st.session_state.manual_mode = True

        bins = st.session_state.bins
        target_options = [f"{b_opt['id']}번" for b_opt in bins] + ["✨ 새 컨테이너"]

        # --- 대시보드 요약 지표 ---
        st.subheader("📊 실시간 적재 요약")
        m1, m2, m3, m4 = st.columns(4)
        packed_qty = sum([len(r['items']) for b in bins for r in b['rows']]) + sum([len(b.get('stacked_items', [])) for b in bins])
        m1.metric("총 화물 수량", f"{len(df)} PKG")
        m2.metric("배정 완료", f"{packed_qty} PKG", f"{packed_qty - len(df)}")
        m3.metric("필요 컨테이너", f"{len(bins)} UNIT")
        m4.metric("평균 중량 효율", f"{sum(b['total_W'] for b in bins)/len(bins):,.0f} kg/UNIT")

        st.markdown("---")

        # --- 컨테이너별 상세 정보 ---
        for b in bins:
            if b['used_L'] == 0 and not b.get('stacked_items'): continue
            
            with st.container():
                st.markdown(f"### 📦 {b['c_label']}")
                c_info1, c_info2 = st.columns([2, 1])
                
                with c_info1:
                    t_data = []
                    for r in b['rows']:
                        for item in r['items']: t_data.append({**item, '상태': '바닥', '이동': f"{b['id']}번"})
                    for s in b.get('stacked_items', []): t_data.append({**s, '상태': '단적', '이동': f"{b['id']}번"})
                    
                    df_edit = pd.DataFrame(t_data)[['상태', 'PKG NO', 'GROUP', 'ITEM', 'L', 'W', 'WEIGHT', '이동']]
                    edited_df = st.data_editor(df_edit, hide_index=True, use_container_width=True, key=f"ed_{b['id']}",
                                            column_config={"이동": st.column_config.SelectboxColumn("🚚 이동", options=target_options)},
                                            disabled=['상태', 'PKG NO', 'GROUP', 'ITEM', 'L', 'W', 'WEIGHT'])
                    
                    if st.button(f"🚀 {b['id']}번 컨테이너 변경사항 적용", key=f"btn_{b['id']}"):
                        moves = [(r['PKG NO'], r['이동']) for _, r in edited_df.iterrows() if r['이동'] != f"{b['id']}번"]
                        if moves:
                            new_alloc = []
                            max_id = max([bx['id'] for bx in st.session_state.bins])
                            for bx in st.session_state.bins:
                                for item in ([i for r in bx['rows'] for i in r['items']] + bx.get('stacked_items', [])):
                                    target = bx['id']
                                    for m_pkg, m_tgt in moves:
                                        if str(item['PKG NO']) == str(m_pkg): target = max_id + 1 if "새" in m_tgt else int(m_tgt.replace("번",""))
                                    new_alloc.append((item, target))
                            new_bins_dict = {}
                            for item, t_id in new_alloc:
                                if t_id not in new_bins_dict: new_bins_dict[t_id] = {'id': t_id, 'type': item['REQ_TYPE'], 'rows': [], 'used_L': 0, 'total_W': 0, 'max_W': 0, 'max_H': 0, 'stacked_items': [], 'groups': set()}
                                pack_items_into_bin([item], new_bins_dict[t_id], max_40_wt, max_40_len)
                            st.session_state.bins = apply_labels(sorted(list(new_bins_dict.values()), key=lambda x: x['id']), max_20_len, max_20_wt, max_40_len-430, max_dry_h)
                            st.rerun()

                with c_info2:
                    st.markdown(f"**길이 점유:** {b['used_L']:,} / {max_40_len if '40ft' in b['c_label'] else max_20_len:,} mm")
                    st.progress(min(1.0, b['used_L']/(max_40_len if '40ft' in b['c_label'] else max_20_len)))
                    st.markdown(f"**중량 점유:** {b['total_W']:,} / {max_40_wt if '40ft' in b['c_label'] else max_20_wt:,} kg")
                    st.progress(min(1.0, b['total_W']/(max_40_wt if '40ft' in b['c_label'] else max_20_wt)))
                    
                    # 도면 시각화
                    d_len = max_20_len if "20ft" in b['c_label'] else max_40_len
                    fig = go.Figure()
                    fig.add_shape(type="rect", x0=0, y0=0, x1=d_len, y1=2350, line=dict(color=MAIN_COLOR, width=2))
                    cx = 0
                    for r in b['rows']:
                        cy = (2350 - r['used_W']) / 2
                        for item in r['items']:
                            fig.add_shape(type="rect", x0=cx, y0=cy, x1=cx+item['L'], y1=cy+item['W'], fillcolor=ACCENT_COLOR, opacity=0.7, line=dict(color="white", width=1))
                            cy += item['W']
                        cx += r['max_L'] + 50
                    fig.update_layout(xaxis=dict(visible=False), yaxis=dict(visible=False), height=180, margin=dict(l=5,r=5,t=5,b=5), paper_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig, use_container_width=True)

        # --- 데이터 다운로드 ---
        mapping = {}
        for bx in bins:
            for item in ([i for r in bx['rows'] for i in r['items']] + bx.get('stacked_items', [])):
                mapping[item['row_idx']] = bx['c_label']
        
        export_df = raw_full.copy()
        target_col = export_df.shape[1] 
        export_df[target_col] = "" 
        export_df.iloc[3, target_col] = "배정 컨테이너" 
        for r_idx, label in mapping.items(): export_df.iloc[r_idx, target_col] = label
            
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            export_df.to_excel(writer, index=False, header=False, sheet_name='CLP_Result')
        
        st.markdown("---")
        st.download_button(label="📥 최종 배정 결과 엑셀 다운로드", data=output.getvalue(), file_name="CALT_CLP_RESULT.xlsx", use_container_width=True)

    except Exception as e: st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
