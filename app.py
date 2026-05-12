import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import math
import os
import io
import re  # 강력한 숫자 추출기

# 💡 안전장치: 코드가 일부 누락되어도 시스템이 뻗지 않도록 기본값을 최상단에 강제 선언
allow_stacking = False
use_balancing = True

# --- 1. 화면 스타일 및 테마 설정 ---
st.set_page_config(page_title="CALT-LOGIS CLP System", layout="wide")

MAIN_COLOR = "#001f3f"
SUB_COLOR = "#f0f2f6"
ACCENT_COLOR = "#007bff"
ALERT_COLOR = "#e74c3c"

st.markdown(f"""
    <style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    * {{ font-family: 'Pretendard', sans-serif; }}
    .main {{ background-color: #f8f9fa; }}
    .header-container {{
        background-color: {MAIN_COLOR};
        padding: 25px;
        border-radius: 10px;
        color: white;
        margin-bottom: 25px;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }}
    [data-testid="stFileUploadDropzone"] {{
        border: 2px dashed {MAIN_COLOR};
        background-color: #eaf1fb;
        padding: 40px;
        border-radius: 12px;
    }}
    [data-testid="stMetric"] {{
        background-color: white;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid {MAIN_COLOR};
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }}
    .stButton>button {{
        width: 100%;
        font-weight: 600;
        color: white;
        background-color: {MAIN_COLOR};
        border-radius: 8px;
        border: none;
        padding: 0.5rem 1rem;
    }}
    div[data-testid="stTable"] {{ font-size: 12px !important; }}
    th {{ background-color: {SUB_COLOR} !important; color: {MAIN_COLOR}; }}
    .container-box {{
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #e1e4e8;
        margin-bottom: 20px;
    }}
    .guide-table {{ font-size: 11px; width: 100%; border-collapse: collapse; }}
    .guide-table th, .guide-table td {{ border: 1px solid #ddd; padding: 5px; text-align: center; }}
    .guide-table th {{ background-color: #eee; }}
    .essential {{ color: {ALERT_COLOR}; font-weight: bold; }}
    </style>
    """, unsafe_allow_html=True)

with st.container():
    st.markdown(f"""
        <div class="header-container">
            <div style="font-size: 30px; font-weight: 800; letter-spacing: -1px;">CALT-LOGIS CLP SYSTEM</div>
            <div style="font-size: 15px; opacity: 0.8; margin-top: 5px;">Busan New Port Center | Logistics Management</div>
        </div>
    """, unsafe_allow_html=True)

# 💡 무적 파싱 로직: 글자, 쉼표, 기호가 섞여있어도 순수 숫자만 강제 추출
def clean_num(val):
    try:
        if pd.isna(val): return 0.0
        s = str(val).replace(',', '').strip()
        if s.upper() in ['', '.', 'X', 'NAN', 'NONE']: return 0.0 
        match = re.search(r'[-+]?\d*\.?\d+', s)
        if match: return float(match.group())
        return 0.0
    except: return 0.0

def get_col_idx(letter):
    try:
        letter = str(letter).upper().strip()
        if not letter: return 0
        ans = 0
        for c in letter: ans = ans * 26 + (ord(c) - ord('A') + 1)
        return ans - 1
    except: return 0

def safe_get(row, idx):
    if 0 <= idx < len(row): return row.iloc[idx]
    return None

def reset_data():
    if 'bins' in st.session_state: del st.session_state['bins']
    if 'manual_mode' in st.session_state: del st.session_state['manual_mode']

# --- 2. 사이드바: 엑셀 열 매핑 가이드 및 설정 ---
with st.sidebar:
    logo_path = "칼트로지스로고.png"
    if os.path.exists(logo_path):
        st.image(logo_path, use_column_width=True)

    with st.expander("📄 엑셀 열 매핑 (알파벳 지정)", expanded=True):
        st.markdown("<p style='font-size:12px; color:gray;'>양식이 바뀌면 아래 알파벳만 변경하세요.</p>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        col_pkg = c1.text_input("PKG NO (필수)", value="B", on_change=reset_data)
        col_wt = c2.text_input("WEIGHT (선택)", value="I", on_change=reset_data)
        col_l = c1.text_input("LENGTH (필수)", value="J", on_change=reset_data)
        col_w = c2.text_input("WIDTH (필수)", value="L", on_change=reset_data)
        col_h = c1.text_input("HEIGHT (필수)", value="N", on_change=reset_data)
        col_grp = c2.text_input("ITEM (참고)", value="D", on_change=reset_data)
        col_desc = c1.text_input("DESC (참고)", value="E", on_change=reset_data)

    st.markdown("---")
    template_data = {
        "Invoice No": [""], "No.of PKG": ["PKG-001"], "LOCATION": [""], "ITEM": ["SAMPLE ITEM"],
        "Description of Goods": ["DETAIL DESC"], "Q'ty": [1], "UNIT": ["EA"], "Net Weight (kg)": [""],
        "Gross Weight (kg)": [550], "Dimension L (mm)": [1200], "X1": ["X"], "Dimension W (mm)": [1000],
        "X2": ["X"], "Dimension H (mm)": [2300]
    }
    df_template = pd.DataFrame(template_data)
    tow = io.BytesIO()
    df_template.to_excel(tow, index=False, header=True, engine='openpyxl')
    st.download_button(
        label="📥 신규 화주용 표준 양식 다운로드", data=tow.getvalue(),
        file_name="CALT_CLP_TEMPLATE.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

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

# --- 짐 배치 핵심 로직 ---
def pack_items_into_bin(pieces, b, max_40_wt, max_40_len, is_stack_allowed):
    for piece in pieces:
        placed = False
        if is_stack_allowed and piece['STACK_OK'] and piece['WEIGHT'] <= 1000 and b['total_W'] + piece['WEIGHT'] <= max_40_wt:
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

def apply_labels(bins, max_20_len, max_20_wt, fr_max_len, max_dry_h, max_hc_h):
    for b in bins:
        is_20ft_size = b['used_L'] <= max_20_len and b['total_W'] <= max_20_wt
        is_ow, is_oh = b['max_W'] > 2350, b['max_H'] > max_hc_h
        is_ol = b['used_L'] > fr_max_len
        tags = []
        if is_oh: tags.append("OH")
        if is_ow: tags.append("OW")
        if is_ol: tags.append("OL")
        if tags:
            base = "20ft Flat Rack" if is_20ft_size else "40ft Flat Rack"
            b['c_label'] = f"{base} [{' + '.join(tags)}] #{b['id']}"
        else:
            if b['max_H'] > max_dry_h: base = "40ft HC"
            else: base = "20ft Dry" if is_20ft_size else "40ft Dry"
            b['c_label'] = f"{base} #{b['id']}"
    return bins

def calculate_expert_packing(df, max_40_wt, max_40_len, max_20_wt, max_20_len, max_dry_h, max_hc_h, is_bal_allowed, is_stack_allowed):
    all_pieces = []
    for _, row in df.iterrows():
        # 💡 에러 방지 1: 크기가 0.X mm라서 0이 되어버리는 유령 화물 방지 (무조건 최소 1mm 보장)
        l = max(1, int(row['L'] + 0.5))
        w = max(1, int(row['W'] + 0.5))
        h = max(1, int(row['H'] + 0.5))
        weight = max(0, int(row['WEIGHT'] + 0.5))
        
        if max(l, w) <= 2350: el, ew = min(l, w), max(l, w)
        else: el, ew = max(l, w), min(l, w)
        all_pieces.append({**row, 'L': el, 'W': ew, 'H': h, 'WEIGHT': weight})
        
    all_pieces.sort(key=lambda x: (-x['W'], -x['H'], -x['L'], x['GROUP']))
    bins = []; c_no = 1
    if is_bal_allowed:
        total_l, total_w = sum(p['L'] for p in all_pieces), sum(p['WEIGHT'] for p in all_pieces)
        est_bins = max(1, math.ceil(total_l / max_40_len), math.ceil(total_w / max_40_wt))
        for _ in range(est_bins):
            bins.append({'id': c_no, 'rows': [], 'used_L': 0, 'total_W': 0, 'max_W': 0, 'max_H': 0, 'stacked_items': [], 'groups': set()})
            c_no += 1
    for piece in all_pieces:
        placed = False
        if is_bal_allowed: bins.sort(key=lambda b: (0 if piece['GROUP'] in b['groups'] else 1, b['used_L']))
        for b in bins:
            pack_items_into_bin([piece], b, max_40_wt, max_40_len, is_stack_allowed)
            if piece in b.get('stacked_items', []) or any(piece in r['items'] for r in b['rows']): 
                placed = True; break
        if not placed:
            new_bin = {'id': c_no, 'rows': [], 'used_L': 0, 'total_W': 0, 'max_W': 0, 'max_H': 0, 'stacked_items': [], 'groups': set()}
            pack_items_into_bin([piece], new_bin, max_40_wt, max_40_len, is_stack_allowed)
            bins.append(new_bin); c_no += 1
    bins = [b for b in bins if b['used_L'] > 0 or b.get('stacked_items')]
    bins.sort(key=lambda x: x['id'])
    for idx, b in enumerate(bins): b['id'] = idx + 1
    return apply_labels(bins, max_20_len, max_20_wt, max_40_len - 430, max_dry_h, max_hc_h)

# --- 3. 메인 화면 로직 ---
st.markdown("### 📤 패킹리스트 업로드 (드래그 앤 드롭)")
file = st.file_uploader("이곳에 파일을 끌어다 놓으세요.", type=['csv', 'xlsx'], on_change=reset_data)

if file is not None:
    try:
        raw_full = pd.read_excel(file, header=None) if file.name.endswith('.xlsx') else pd.read_csv(file, header=None)
        
        idx_pkg = get_col_idx(col_pkg)
        idx_wt = get_col_idx(col_wt)
        idx_l = get_col_idx(col_l)
        idx_w = get_col_idx(col_w)
        idx_h = get_col_idx(col_h)
        idx_grp = get_col_idx(col_grp)
        idx_desc = get_col_idx(col_desc)

        p_data = []
        debug_logs = []  
        last_pkg_no = "UNKNOWN"

        for i in range(len(raw_full)):
            row = raw_full.iloc[i]
            raw_pkg = safe_get(row, idx_pkg)
            str_pkg = str(raw_pkg).strip()
            
            if pd.notna(raw_pkg) and str_pkg not in ['', 'None', 'nan', 'NaN', '.', 'No.of PKG', 'PKG NO', 'Invoice No']:
                last_pkg_no = str_pkg
                
            pkg_v = last_pkg_no
            raw_l = safe_get(row, idx_l)
            raw_w = safe_get(row, idx_w)
            raw_h = safe_get(row, idx_h)
            raw_wt = safe_get(row, idx_wt)
            
            l_v = clean_num(raw_l)
            w_v = clean_num(raw_w)
            h_v = clean_num(raw_h)
            weight_v = clean_num(raw_wt)
            
            debug_logs.append({
                '행번호': i + 1, 'PKG (원본)': raw_pkg, 'PKG (인식)': pkg_v,
                'L(원본)': raw_l, 'L(숫자)': l_v, 'W(원본)': raw_w, 'W(숫자)': w_v, 'H(원본)': raw_h, 'H(숫자)': h_v
            })
            
            if l_v == 0 or w_v == 0 or h_v == 0 or pkg_v == "UNKNOWN":
                continue
            
            s_ok = row.astype(str).str.contains('단적허용').any() if allow_stacking else False
            val_group_item = str(safe_get(row, idx_grp)).strip() if pd.notna(safe_get(row, idx_grp)) else "-"
            val_desc = str(safe_get(row, idx_desc)).strip() if pd.notna(safe_get(row, idx_desc)) else "-"

            p_data.append({
                'PKG NO': pkg_v, 'GROUP': val_group_item, 'ITEM': val_group_item, 'DESC': val_desc, 
                'L': l_v, 'W': w_v, 'H': h_v, 'WEIGHT': weight_v, 'STACK_OK': s_ok, 'row_idx': i
            })
        
        df = pd.DataFrame(p_data)
        
        if df.empty:
            st.error("⚠️ 설정된 열(Column)에서 필수 항목(L, W, H)을 0보다 큰 숫자로 변환하지 못했습니다.")
            with st.expander("🔍 AI 데이터 스캔 진단표 펼쳐보기", expanded=True):
                st.dataframe(pd.DataFrame(debug_logs).head(30))
        else:
            df = df[~df['ITEM'].str.upper().str.contains('TOTAL', na=False)]

            if 'bins' not in st.session_state or not st.session_state.get('manual_mode', False):
                st.session_state.bins = calculate_expert_packing(df, max_40_wt, max_40_len, max_20_wt, max_20_len, max_dry_h, max_hc_h, use_balancing, allow_stacking)
                st.session_state.manual_mode = True

            bins = st.session_state.bins
            target_options = [f"{b_opt['id']}번" for b_opt in bins] + ["✨ 새 컨테이너"]

            st.subheader("📊 실시간 적재 요약")
            m1, m2, m3, m4 = st.columns(4)
            packed_qty = sum([len(r['items']) for b in bins for r in b['rows']]) + sum([len(b.get('stacked_items', [])) for b in bins])
            m1.metric("전체 화물", f"{len(df)} PKG")
            m2.metric("배정 완료", f"{packed_qty} PKG") 
            m3.metric("컨테이너 수", f"{len(bins)} UNIT")
            m4.metric("평균 중량", f"{sum(b['total_W'] for b in bins)/len(bins):,.0f} kg")

            for b in bins:
                if b['used_L'] == 0 and not b.get('stacked_items'): continue
                st.markdown(f'<div class="container-box">', unsafe_allow_html=True)
                st.markdown(f"### 📦 {b['c_label']}")
                t_data = []
                for r in b['rows']:
                    for item in r['items']: t_data.append({**item, '위치': '바닥', '이동': f"{b['id']}번"})
                for s in b.get('stacked_items', []): t_data.append({**s, '위치': '단적', '이동': f"{b['id']}번"})
                
                df_edit = pd.DataFrame(t_data)[['위치', 'PKG NO', 'ITEM', 'L', 'W', 'H', 'WEIGHT', '이동']]
                edited_df = st.data_editor(df_edit, hide_index=True, use_container_width=True, key=f"ed_{b['id']}",
                                        column_config={"이동": st.column_config.SelectboxColumn("🚚 이동", options=target_options)},
                                        disabled=['위치', 'PKG NO', 'ITEM', 'L', 'W', 'H', 'WEIGHT'])
                
                if st.button(f"🚀 {b['id']}번 변경사항 적용", key=f"btn_{b['id']}"):
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
                            if t_id not in new_bins_dict: new_bins_dict[t_id] = {'id': t_id, 'rows': [], 'used_L': 0, 'total_W': 0, 'max_W': 0, 'max_H': 0, 'stacked_items': [], 'groups': set()}
                            pack_items_into_bin([item], new_bins_dict[t_id], max_40_wt, max_40_len, allow_stacking)
                        st.session_state.bins = apply_labels(sorted(list(new_bins_dict.values()), key=lambda x: x['id']), max_20_len, max_20_wt, max_40_len-430, max_dry_h, max_hc_h)
                        st.rerun()

                with st.expander("👁️ 적재 단면도 및 제원 확인", expanded=False):
                    try:
                        cur_max_l = float(max_20_len or 5899) if "20ft" in b['c_label'] else float(max_40_len or 12034)
                        cur_max_w = float(max_20_wt or 28250) if "20ft" in b['c_label'] else float(max_40_wt or 29500)
                        cur_max_h = float(max_hc_h or 2695) if "HC" in b['c_label'] else float(max_dry_h or 2390)
                        used_width = max([r['used_W'] for r in b['rows']] + [0])
                        max_stacked_h = max([s['H'] for s in b.get('stacked_items', [])] + [0])
                        used_height = b['max_H'] + max_stacked_h if b.get('stacked_items') else b['max_H']
                        
                        c1, c2, c3, c4 = st.columns(4)
                        c1.markdown(f"**📏 길이:** {b['used_L']:,}/{cur_max_l:,.0f}mm"); c1.progress(min(1.0, b['used_L']/cur_max_l))
                        c2.markdown(f"**⚖️ 중량:** {b['total_W']:,}/{cur_max_w:,.0f}kg"); c2.progress(min(1.0, b['total_W']/cur_max_w))
                        
                        if used_width > 2350: c3.markdown(f"**↔️ 폭:** {used_width:,}/2350mm <span style='color:{ALERT_COLOR};font-weight:bold;'>[OW +{used_width-2350:,}]</span>", unsafe_allow_html=True)
                        else: c3.markdown(f"**↔️ 폭:** {used_width:,}/2350mm"); c3.progress(min(1.0, used_width/2350))
                        
                        if used_height > cur_max_h: c4.markdown(f"**↕️ 높이:** {used_height:,}/{cur_max_h:,.0f}mm <span style='color:{ALERT_COLOR};font-weight:bold;'>[OH +{used_height-cur_max_h:,.0f}]</span>", unsafe_allow_html=True)
                        else: c4.markdown(f"**↕️ 높이:** {used_height:,}/{cur_max_h:,.0f}mm"); c4.progress(min(1.0, used_height/cur_max_h))
                        
                        fig = go.Figure()
                        
                        # 💡 에러 방지 2: 속성 오류를 방지하기 위해 엄격한 dict 속성(line=dict(width=0)) 적용
                        bg_x1 = max(cur_max_l, b['used_L'])
                        if bg_x1 > b['used_L']:
                            fig.add_shape(type="rect", x0=b['used_L'], y0=0, x1=bg_x1, y1=2350, fillcolor="#e1e4e8", opacity=0.4, line=dict(width=0))
                        
                        fig.add_shape(type="rect", x0=0, y0=0, x1=bg_x1, y1=2350, line=dict(color=MAIN_COLOR, width=2))
                        
                        cx = 0
                        for r in b['rows']:
                            cy = (2350 - r['used_W']) / 2
                            for item in r['items']:
                                fig.add_shape(type="rect", x0=cx, y0=cy, x1=cx+item['L'], y1=cy+item['W'], fillcolor=ACCENT_COLOR, opacity=0.8, line=dict(color="white", width=1))
                                fig.add_annotation(x=cx+item['L']/2, y=cy+item['W']/2, text=str(item['PKG NO']), showarrow=False, font=dict(color="white", size=10))
                                cy += item['W']
                            cx += r['max_L']
                            
                        fig.add_shape(type="line", x0=b['used_L'], y0=-200, x1=b['used_L'], y1=2800, line=dict(color=ALERT_COLOR, width=2, dash="dash"))
                        
                        if b['used_L'] > 100: fig.add_annotation(x=b['used_L']/2, y=2650, text=f"적재: {b['used_L']:,}mm", showarrow=False, font=dict(color=MAIN_COLOR, size=13, weight="bold"))
                        if cur_max_l - b['used_L'] > 100: fig.add_annotation(x=b['used_L'] + (cur_max_l - b['used_L'])/2, y=2650, text=f"잔여: {cur_max_l - b['used_L']:,.0f}mm", showarrow=False, font=dict(color=ALERT_COLOR, size=13, weight="bold"))
                        
                        fig.update_layout(xaxis=dict(visible=False, range=[-200, bg_x1+400]), yaxis=dict(visible=False, range=[-300, 3100]), height=280, margin=dict(l=10, r=10, t=30, b=10), paper_bgcolor="rgba(0,0,0,0)")
                        st.plotly_chart(fig, use_container_width=True, key=f"plot_{b['id']}")
                        
                    # 💡 에러 방지 3: 특정 컨테이너 도면이 뻗어도 앱 전체가 다운되지 않도록 예외 처리
                    except Exception as e:
                        st.error(f"🚨 이 컨테이너({b['c_label']})의 도면을 그리는 중 내부 오류가 발생했습니다.")
                        st.info(f"기술적 오류 코드: {e}")

                st.markdown('</div>', unsafe_allow_html=True)

            export_df = raw_full.copy()
            target_col = export_df.shape[1]
            export_df[target_col] = ""; export_df.iloc[0, target_col] = "배정 컨테이너"
            for r_idx, label in {item['row_idx']: bx['c_label'] for bx in bins for item in ([i for r in bx['rows'] for i in r['items']] + bx.get('stacked_items', []))}.items(): export_df.iloc[r_idx, target_col] = label
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer: export_df.to_excel(writer, index=False, header=False)
            st.markdown("---")
            st.download_button(label="📥 최종 결과 다운로드 (배정 정보 포함)", data=output.getvalue(), file_name="CLP_RESULT_FINAL.xlsx", use_container_width=True)

    except Exception as e: st.error(f"데이터 처리 중 오류 발생: {e}")
