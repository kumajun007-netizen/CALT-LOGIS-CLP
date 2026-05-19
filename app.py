import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import math
import os
import io
from openpyxl import load_workbook
from copy import copy
from collections import defaultdict

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

def clean_num(val):
    try:
        if pd.isna(val) or str(val).strip() in ['', '.', 'X', 'x', 'NaN', 'nan']: return 0.0
        return float(str(val).replace(',', '').strip())
    except: return 0.0

# --- 2. 사이드바: 엑셀 열 가이드 및 설정 ---
with st.sidebar:
    logo_path = "칼트로지스로고.png"
    if os.path.exists(logo_path):
        st.image(logo_path, use_column_width=True)

    with st.expander("📄 엑셀 업로드 표준 규격 (열 고정)", expanded=True):
        st.markdown(f"""
        <table class="guide-table">
            <tr><th>항목</th><th>엑셀 열</th><th>구분</th></tr>
            <tr class="essential"><td>No.of PKG</td><td>B 열</td><td>[필수]</td></tr>
            <tr class="essential"><td>Q'ty</td><td>F 열</td><td>[필수]</td></tr>
            <tr class="essential"><td>LENGTH (L)</td><td>J 열</td><td>[필수]</td></tr>
            <tr class="essential"><td>WIDTH (W)</td><td>L 열</td><td>[필수]</td></tr>
            <tr class="essential"><td>HEIGHT (H)</td><td>N 열</td><td>[필수]</td></tr>
            <tr><td>NET WEIGHT</td><td>I 열</td><td>선택</td></tr>
            <tr><td>REMARK (2단/3단)</td><td>O 열</td><td>선택</td></tr>
            <tr><td>LOAD (FORK_L/W/4WAY)</td><td>P 열</td><td>선택</td></tr>
            <tr><td>ITEM</td><td>D 열</td><td>참고</td></tr>
            <tr><td>DESCRIPTION</td><td>E 열</td><td>참고</td></tr>
        </table>
        <p style='font-size:11px; color:gray; margin-top:10px;'>* 데이터는 6번째 줄부터 시작하는 것을 권장합니다.</p>
        """, unsafe_allow_html=True)

    st.markdown("---")
    template_data = {
        "Invoice No": [""],
        "No.of PKG": ["PKG-001"],
        "LOCATION": [""],
        "ITEM": ["SAMPLE ITEM"],
        "Description of Goods": ["DETAIL DESC"],
        "Q'ty": [1],
        "UNIT": ["EA"],
        "Net Weight (kg)": [500],
        "Gross Weight (kg)": [550],
        "Dimension L (mm)": [1200],
        "X1": ["X"],
        "Dimension W (mm)": [1000],
        "X2": ["X"],
        "Dimension H (mm)": [2300],
        "REMARK": ["BOX,2단"],
        "LOAD": [""]
    }
    df_template = pd.DataFrame(template_data)
    tow = io.BytesIO()
    df_template.to_excel(tow, index=False, header=True, engine='openpyxl')
    st.download_button(
        label="📥 신규 화주용 양식 다운로드",
        data=tow.getvalue(),
        file_name="CALT_CLP_TEMPLATE.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
        max_width = st.number_input("컨테이너 폭 (mm)", 2200, 2500, 2350)

    with st.expander("🛠 적재 로직", expanded=True):
        use_balancing = st.checkbox("⚖️ 균분적재 (Balancing) ※ 끄면 컨테이너 수 최소화", value=False)
        balancing_eff = st.slider("균분적재 면적 효율 가정 (%)", 70, 100, 95, 1,
                                  help="높일수록 컨테이너 수 최소화. 낮추면 여유분 확보.")
        allow_stacking = st.checkbox("🏢 다단적재 (REMARK 2단/3단 반영)", value=True)
        force_container_type = st.selectbox(
            "🏗 컨테이너 타입 강제 선정",
            ["자동 (높이 기준)", "40ft HC 강제", "40ft DRY 강제"],
            index=1,  # 기본값 HC 강제
            help="자동: 누적 H가 DRY 한계(2390mm) 초과 시에만 HC. 강제: 화주 요청 등 실무 케이스."
        )
        global_load_mode = st.selectbox(
            "🚜 전체 LOAD 모드 (개별 P열 키워드 우선)",
            ["일반 (폭 활용 최대화)", "FORK_W (긴쪽→W)", "FORK_L (긴쪽→L)", "4WAY (자유)"],
            index=2  # 기본값 FORK_L (사용자 케이스)
        )

    if st.button("🔄 AI 재계산 실행"):
        st.session_state['manual_mode'] = False

# ===================================================================
# --- REMARK / LOAD 파싱 ---
# ===================================================================
def parse_remark(remark_str):
    """O열 REMARK 파싱: BOX, 2단, 3단"""
    if remark_str is None or pd.isna(remark_str):
        return {'layers': 1, 'is_box': True}
    s = str(remark_str).strip()
    layers = 1
    if '3단' in s: layers = 3
    elif '2단' in s: layers = 2
    is_box = 'BOX' in s.upper() or layers > 1 or s == ''
    return {'layers': layers, 'is_box': is_box}

def parse_load(load_str, global_mode):
    """P열 LOAD 파싱: FORK_L / FORK_W / 4WAY. 개별 키워드 우선, 없으면 전체 모드"""
    if load_str is not None and not pd.isna(load_str):
        s = str(load_str).upper().strip()
        if 'FORK_L' in s: return 'FORK_L'
        if 'FORK_W' in s: return 'FORK_W'
        if '4WAY' in s: return '4WAY'
    # 전체 모드 적용
    if 'FORK_L' in global_mode: return 'FORK_L'
    if 'FORK_W' in global_mode: return 'FORK_W'
    if '4WAY' in global_mode: return '4WAY'
    return None  # 일반

def apply_rotation(l_raw, w_raw, load_mode, max_w_container):
    """LOAD 키워드 기준으로 L/W 회전 결정"""
    if load_mode == 'FORK_L':
        # 긴쪽을 L방향 (FR 빈공간 채우기)
        return max(l_raw, w_raw), min(l_raw, w_raw)
    elif load_mode == 'FORK_W':
        # 긴쪽을 W방향 (LSE 기본)
        return min(l_raw, w_raw), max(l_raw, w_raw)
    else:
        # 일반/4WAY: 폭 활용 최대화 (긴쪽이 컨테이너 폭에 들어가면 W로)
        if max(l_raw, w_raw) <= max_w_container:
            return min(l_raw, w_raw), max(l_raw, w_raw)
        else:
            return max(l_raw, w_raw), min(l_raw, w_raw)

# ===================================================================
# --- 다단 묶음 → 풋프린트 단위 생성 ---
# ===================================================================
def build_footprints(df, max_w_container, allow_stack):
    """
    같은 (PKG, 치수, 단수) 박스를 N단씩 묶어 풋프린트 단위로 만든다.
    풋프린트 1개 = 컨테이너 바닥에 차지하는 면적 1개 = N단 누적된 박스 묶음.
    """
    # 박스 단위 리스트 (Q'ty 복제 완료 후 들어옴)
    grouped = defaultdict(list)
    for _, row in df.iterrows():
        key = (row['PKG NO'], int(row['L_raw']), int(row['W_raw']), int(row['H_raw']),
               row['layers'] if allow_stack else 1, row['load_mode'])
        grouped[key].append(row.to_dict())

    footprints = []
    for key, items in grouped.items():
        pkg, l_raw, w_raw, h_raw, layers, load_mode = key
        i = 0
        while i < len(items):
            stack = items[i:i+layers]
            i += layers
            n = len(stack)
            l_fit, w_fit = apply_rotation(l_raw, w_raw, load_mode, max_w_container)
            fp = {
                'PKG NO': stack[0]['PKG NO'],
                'ITEM': stack[0].get('ITEM', '-'),
                'GROUP': stack[0].get('GROUP', '-'),
                'L': l_fit,  # 회전 후 L
                'W': w_fit,  # 회전 후 W
                'H': h_raw * n,  # 누적 높이
                'H_per_layer': h_raw,
                'stack_count': n,
                'WEIGHT': sum(s['WEIGHT'] for s in stack),
                'load_mode': load_mode if load_mode else '-',
                'STACK_OK': n > 1,
                'row_idx': stack[0]['row_idx'],  # 대표값
                'sub_items': stack,  # 원본 박스들
            }
            footprints.append(fp)
    return footprints

# ===================================================================
# --- 컨테이너 타입 사전 선정 ---
# ===================================================================
def decide_container_height(footprints, max_dry_h, max_hc_h, force_type='자동 (높이 기준)'):
    """전체 풋프린트의 최대 누적 높이로 HC/Dry 선정. 강제 옵션 우선."""
    if 'HC' in force_type:
        return max_hc_h, 'HC'
    if 'DRY' in force_type:
        return max_dry_h, 'DRY'
    # 자동
    if not footprints:
        return max_dry_h, 'DRY'
    max_h = max(fp['H'] for fp in footprints)
    if max_h > max_dry_h:
        return max_hc_h, 'HC'
    else:
        return max_dry_h, 'DRY'

# ===================================================================
# --- 풋프린트 패킹 로직 (개선) ---
# ===================================================================
def pack_footprint_into_bin(fp, b, max_wt, max_len, max_h_container, max_w_container):
    """
    한 풋프린트(= 다단 묶음)를 컨테이너에 적재.
    기존 row(=L방향 슬라이스) 안에 W방향으로 추가 시도 후,
    안 되면 새 row(L방향 신규 슬라이스) 생성.
    """
    # 중량/높이 한계 사전 체크
    if b['total_W'] + fp['WEIGHT'] > max_wt: return False
    if fp['H'] > max_h_container: return False
    if fp['L'] > max_len: return False
    if fp['W'] > max_w_container: return False

    # 1. 기존 row에 W방향으로 추가 시도
    for r in b['rows']:
        temp_max_L = max(r['max_L'], fp['L'])
        new_used_L = b['used_L'] + (temp_max_L - r['max_L'])
        if (r['used_W'] + fp['W'] <= max_w_container
            and new_used_L <= max_len):
            r['items'].append(fp)
            r['used_W'] += fp['W']
            b['used_L'] = new_used_L
            r['max_L'] = temp_max_L
            b['total_W'] += fp['WEIGHT']
            b['max_W'] = max(b['max_W'], fp['W'])
            b['max_H'] = max(b['max_H'], fp['H'])
            b['groups'].add(fp['GROUP'])
            return True

    # 2. 새 row 생성 (L방향 신규 슬라이스)
    if b['used_L'] + fp['L'] <= max_len:
        b['rows'].append({'items': [fp], 'used_W': fp['W'], 'max_L': fp['L']})
        b['used_L'] += fp['L']
        b['total_W'] += fp['WEIGHT']
        b['max_W'] = max(b['max_W'], fp['W'])
        b['max_H'] = max(b['max_H'], fp['H'])
        b['groups'].add(fp['GROUP'])
        return True

    return False

# ===================================================================
# --- 라벨 적용 (HC 사전 선정 반영) ---
# ===================================================================
def apply_labels(bins, max_20_len, max_20_wt, fr_max_len, max_dry_h, max_hc_h, max_w_container, container_h_used):
    for b in bins:
        is_20ft_size = b['used_L'] <= max_20_len and b['total_W'] <= max_20_wt
        is_ow = b['max_W'] > max_w_container
        is_oh = b['max_H'] > max_hc_h
        is_ol = b['used_L'] > fr_max_len
        tags = []
        if is_oh: tags.append("OH")
        if is_ow: tags.append("OW")
        if is_ol: tags.append("OL")
        if tags:
            base = "20ft Flat Rack" if is_20ft_size else "40ft Flat Rack"
            b['c_label'] = f"{base} [{' + '.join(tags)}] #{b['id']}"
        else:
            # HC 사전 선정 결과 우선 반영
            if container_h_used == 'HC' or b['max_H'] > max_dry_h:
                base = "40ft HC"
            else:
                base = "20ft Dry" if is_20ft_size else "40ft Dry"
            b['c_label'] = f"{base} #{b['id']}"
    return bins

# ===================================================================
# --- 메인 패킹 계산 ---
# ===================================================================
def calculate_expert_packing(df, max_40_wt, max_40_len, max_20_wt, max_20_len,
                             max_dry_h, max_hc_h, max_w_container,
                             use_balancing, allow_stack, force_container_type,
                             global_load_mode, balancing_eff=95):
    # 1. 풋프린트 생성 (다단 묶음)
    fps = build_footprints(df, max_w_container, allow_stack)

    # 2. 컨테이너 타입 사전 선정 (강제 옵션 우선)
    max_h_container, h_type = decide_container_height(fps, max_dry_h, max_hc_h, force_container_type)

    # 3. 풋프린트 정렬: W 큰순 → H 큰순 → L 큰순 → GROUP
    fps.sort(key=lambda x: (-x['W'], -x['H'], -x['L'], str(x['GROUP'])))

    bins = []
    c_no = 1

    # 4. 균분적재용 사전 컨테이너 생성
    if use_balancing and fps:
        total_w_kg = sum(fp['WEIGHT'] for fp in fps)
        floor_area_needed = sum(fp['L'] * fp['W'] for fp in fps)
        floor_area_per_bin = max_40_len * max_w_container * (balancing_eff / 100.0)
        est_by_area = math.ceil(floor_area_needed / floor_area_per_bin)
        est_by_wt = math.ceil(total_w_kg / max_40_wt) if total_w_kg > 0 else 1
        est_bins = max(1, est_by_area, est_by_wt)
        for _ in range(est_bins):
            bins.append({
                'id': c_no, 'rows': [], 'used_L': 0, 'total_W': 0,
                'max_W': 0, 'max_H': 0, 'groups': set()
            })
            c_no += 1

    # 5. 풋프린트 배치
    for fp in fps:
        if use_balancing and bins:
            # 같은 GROUP 우선 + 적재 길이 적은 컨테이너 우선
            bins.sort(key=lambda b: (0 if fp['GROUP'] in b['groups'] else 1, b['used_L']))

        placed = False
        for b in bins:
            if pack_footprint_into_bin(fp, b, max_40_wt, max_40_len, max_h_container, max_w_container):
                placed = True
                break

        if not placed:
            new_bin = {
                'id': c_no, 'rows': [], 'used_L': 0, 'total_W': 0,
                'max_W': 0, 'max_H': 0, 'groups': set()
            }
            if pack_footprint_into_bin(fp, new_bin, max_40_wt, max_40_len, max_h_container, max_w_container):
                bins.append(new_bin)
                c_no += 1
            else:
                # 한 풋프린트가 너무 커서 못 들어감 → FR 케이스
                new_bin['rows'].append({'items': [fp], 'used_W': fp['W'], 'max_L': fp['L']})
                new_bin['used_L'] = fp['L']
                new_bin['total_W'] = fp['WEIGHT']
                new_bin['max_W'] = fp['W']
                new_bin['max_H'] = fp['H']
                new_bin['groups'].add(fp['GROUP'])
                bins.append(new_bin)
                c_no += 1

    # 6. 빈 컨테이너 정리
    bins = [b for b in bins if b['used_L'] > 0]
    bins.sort(key=lambda x: x['id'])
    for idx, b in enumerate(bins):
        b['id'] = idx + 1

    return apply_labels(bins, max_20_len, max_20_wt, max_40_len - 430,
                        max_dry_h, max_hc_h, max_w_container, h_type)

# ===================================================================
# --- 3. 메인 화면 로직 ---
# ===================================================================
def reset_data():
    if 'bins' in st.session_state: del st.session_state['bins']
    if 'manual_mode' in st.session_state: del st.session_state['manual_mode']

st.markdown("### 📤 패킹리스트 업로드 (드래그 앤 드롭)")
file = st.file_uploader("이곳에 파일을 끌어다 놓으세요.", type=['csv', 'xlsx'], on_change=reset_data)

if file is not None:
    try:
        selected_sheet = None

        # 엑셀 고정 열 매핑 (0-base)
        COL_PKG = 1       # B열: No.of PKG
        COL_ITEM = 3      # D열: ITEM
        COL_DESC = 4      # E열: DESCRIPTION
        COL_QTY = 5       # F열: Q'ty  ★ 신규 반영
        COL_WEIGHT = 8    # I열: WEIGHT
        COL_L = 9         # J열: LENGTH
        COL_W = 11        # L열: WIDTH
        COL_H = 13        # N열: HEIGHT
        COL_REMARK = 14   # O열: REMARK ★ 신규 반영
        COL_LOAD = 15     # P열: LOAD   ★ 신규 반영

        def count_valid_rows(sheet_df):
            valid_count = 0
            for _, r in sheet_df.iterrows():
                try:
                    pkg = str(r.iloc[COL_PKG]).replace('00:00:00', '').replace('.0', '').strip() if pd.notna(r.iloc[COL_PKG]) else ""
                    l = clean_num(r.iloc[COL_L])
                    w = clean_num(r.iloc[COL_W])
                    h = clean_num(r.iloc[COL_H])
                    if pkg.lower() not in ['', 'nan', 'none', '.'] and l > 0 and w > 0 and h > 0:
                        valid_count += 1
                except:
                    continue
            return valid_count

        if file.name.lower().endswith('.xlsx'):
            file.seek(0)
            excel_sheets = pd.read_excel(file, sheet_name=None, header=None)

            raw_full = None
            best_count = 0

            for sheet_name, sheet_df in excel_sheets.items():
                valid_count = count_valid_rows(sheet_df)
                if valid_count > best_count:
                    best_count = valid_count
                    raw_full = sheet_df
                    selected_sheet = sheet_name

            if raw_full is None or best_count == 0:
                st.error("엑셀에서 B/J/L/N 열 기준의 필수 데이터를 찾지 못했습니다.")
                with st.expander("🔎 확인 필요 사항", expanded=True):
                    st.write("- B열: No.of PKG / F열: Q'ty / J열: LENGTH / L열: WIDTH / N열: HEIGHT")
                    st.write("- O열(REMARK)에 '2단','3단' 키워드 / P열(LOAD)에 FORK_L/W/4WAY 가능")
                st.stop()

            st.info(f"읽은 시트: {selected_sheet} / 인식된 화물 행 수: {best_count} PKG (Q'ty 복제 전)")

        else:
            file.seek(0)
            raw_full = pd.read_csv(file, header=None)
            selected_sheet = None

        p_data = []
        total_qty = 0

        # ★ 핵심 변경: Q'ty 만큼 박스 복제
        for orig_idx in range(len(raw_full)):
            row = raw_full.iloc[orig_idx]

            try:
                pkg_v = str(row.iloc[COL_PKG]).replace('00:00:00', '').replace('.0', '').strip() if pd.notna(row.iloc[COL_PKG]) else ""
                l_v = clean_num(row.iloc[COL_L])
                w_v = clean_num(row.iloc[COL_W])
                h_v = clean_num(row.iloc[COL_H])
                weight_v = clean_num(row.iloc[COL_WEIGHT])
                qty_v = int(clean_num(row.iloc[COL_QTY])) if row.shape[0] > COL_QTY else 1
                if qty_v <= 0: qty_v = 1

                # REMARK / LOAD 파싱
                remark_raw = row.iloc[COL_REMARK] if row.shape[0] > COL_REMARK else None
                load_raw = row.iloc[COL_LOAD] if row.shape[0] > COL_LOAD else None
                r_info = parse_remark(remark_raw)
                load_mode = parse_load(load_raw, global_load_mode)

                if pkg_v.lower() in ['', 'nan', 'none', '.'] or l_v == 0 or w_v == 0 or h_v == 0:
                    continue

                # Q'ty 만큼 복제
                for q_i in range(qty_v):
                    p_data.append({
                        'PKG NO': pkg_v,
                        'GROUP': str(row.iloc[COL_DESC]) if pd.notna(row.iloc[COL_DESC]) else "-",
                        'ITEM': str(row.iloc[COL_ITEM]) if pd.notna(row.iloc[COL_ITEM]) else "-",
                        'DESC': str(row.iloc[COL_DESC]) if pd.notna(row.iloc[COL_DESC]) else "-",
                        'L_raw': int(l_v + 0.5),
                        'W_raw': int(w_v + 0.5),
                        'H_raw': int(h_v + 0.5),
                        'WEIGHT': int(weight_v + 0.5),
                        'layers': r_info['layers'],
                        'load_mode': load_mode,
                        'row_idx': int(orig_idx),
                    })
                total_qty += qty_v
            except:
                continue

        df = pd.DataFrame(p_data)
        if df.empty:
            st.warning("⚠️ 지정된 열(B, F, J, L, N)에서 필수 데이터를 찾을 수 없습니다. 양식을 확인하세요.")
        else:
            st.success(f"✅ 총 {len(df):,} 박스 복제 완료 (Q'ty 반영)")

            if 'bins' not in st.session_state or not st.session_state.get('manual_mode', False):
                st.session_state.bins = calculate_expert_packing(
                    df, max_40_wt, max_40_len, max_20_wt, max_20_len,
                    max_dry_h, max_hc_h, max_width,
                    use_balancing, allow_stacking, force_container_type,
                    global_load_mode, balancing_eff
                )
                st.session_state.manual_mode = True

            bins = st.session_state.bins
            target_options = [f"{b_opt['id']}번" for b_opt in bins] + ["✨ 새 컨테이너"]

            st.subheader("📊 실시간 적재 요약")
            m1, m2, m3, m4 = st.columns(4)
            packed_fp = sum(len(r['items']) for b in bins for r in b['rows'])
            packed_boxes = sum(
                sum(fp['stack_count'] for r in b['rows'] for fp in r['items'])
                for b in bins
            )
            m1.metric("전체 화물(박스)", f"{len(df):,} PKG")
            m2.metric("배정 박스", f"{packed_boxes:,} PKG")
            m3.metric("컨테이너 수", f"{len(bins)} UNIT")
            avg_wt = sum(b['total_W'] for b in bins) / len(bins) if bins else 0
            m4.metric("평균 중량", f"{avg_wt:,.0f} kg")

            for b in bins:
                if b['used_L'] == 0: continue
                st.markdown(f'<div class="container-box">', unsafe_allow_html=True)
                st.markdown(f"### 📦 {b['c_label']}")

                # 풋프린트 상세
                t_data = []
                for r in b['rows']:
                    for fp in r['items']:
                        t_data.append({
                            'PKG NO': fp['PKG NO'],
                            'ITEM': fp['ITEM'],
                            'L(회전후)': fp['L'],
                            'W(회전후)': fp['W'],
                            'H(누적)': fp['H'],
                            '단수': fp['stack_count'],
                            'LOAD': fp['load_mode'],
                            'WEIGHT': fp['WEIGHT'],
                            '이동': f"{b['id']}번"
                        })

                df_edit = pd.DataFrame(t_data)
                edited_df = st.data_editor(
                    df_edit, hide_index=True, use_container_width=True, key=f"ed_{b['id']}",
                    column_config={"이동": st.column_config.SelectboxColumn("🚚 이동", options=target_options)},
                    disabled=['PKG NO', 'ITEM', 'L(회전후)', 'W(회전후)', 'H(누적)', '단수', 'LOAD', 'WEIGHT']
                )

                with st.expander("👁️ 적재 단면도 및 제원 확인", expanded=True):
                    cur_max_l = max_20_len if "20ft" in b['c_label'] else max_40_len
                    cur_max_w = max_20_wt if "20ft" in b['c_label'] else max_40_wt
                    cur_max_h = max_hc_h if "HC" in b['c_label'] else max_dry_h
                    used_width = max([r['used_W'] for r in b['rows']] + [0])
                    used_height = b['max_H']

                    c1, c2, c3, c4 = st.columns(4)
                    c1.markdown(f"**📏 길이:** {b['used_L']:,}/{cur_max_l:,}mm")
                    c1.progress(min(1.0, b['used_L']/cur_max_l))
                    c2.markdown(f"**⚖️ 중량:** {b['total_W']:,}/{cur_max_w:,}kg")
                    c2.progress(min(1.0, b['total_W']/cur_max_w))

                    if used_width > max_width:
                        c3.markdown(f"**↔️ 폭:** {used_width:,}/{max_width:,}mm <span style='color:{ALERT_COLOR};font-weight:bold;'>[OW +{used_width-max_width:,}]</span>", unsafe_allow_html=True)
                    else:
                        c3.markdown(f"**↔️ 폭:** {used_width:,}/{max_width:,}mm")
                        c3.progress(min(1.0, used_width/max_width))

                    if used_height > cur_max_h:
                        c4.markdown(f"**↕️ 높이:** {used_height:,}/{cur_max_h:,}mm <span style='color:{ALERT_COLOR};font-weight:bold;'>[OH +{used_height-cur_max_h:,}]</span>", unsafe_allow_html=True)
                    else:
                        c4.markdown(f"**↕️ 높이:** {used_height:,}/{cur_max_h:,}mm")
                        c4.progress(min(1.0, used_height/cur_max_h))

                    # 평면도
                    fig = go.Figure()
                    fig.add_shape(type="rect", x0=b['used_L'], y0=0, x1=cur_max_l, y1=max_width,
                                  fillcolor="#e1e4e8", opacity=0.4, line_width=0)
                    fig.add_shape(type="rect", x0=0, y0=0, x1=cur_max_l, y1=max_width,
                                  line=dict(color=MAIN_COLOR, width=2))
                    cx = 0
                    for r in b['rows']:
                        cy = (max_width - r['used_W']) / 2
                        for fp in r['items']:
                            fig.add_shape(type="rect",
                                          x0=cx, y0=cy, x1=cx+fp['L'], y1=cy+fp['W'],
                                          fillcolor=ACCENT_COLOR, opacity=0.8,
                                          line=dict(color="white", width=1))
                            label_text = f"{fp['PKG NO']}<br>{fp['stack_count']}단" if fp['stack_count'] > 1 else str(fp['PKG NO'])
                            fig.add_annotation(x=cx+fp['L']/2, y=cy+fp['W']/2,
                                               text=label_text, showarrow=False,
                                               font=dict(color="white", size=9))
                            cy += fp['W']
                        cx += r['max_L']
                    fig.add_shape(type="line", x0=b['used_L'], y0=-200, x1=b['used_L'], y1=max_width+450,
                                  line=dict(color=ALERT_COLOR, width=2, dash="dash"))
                    if b['used_L'] > 100:
                        fig.add_annotation(x=b['used_L']/2, y=max_width+300,
                                           text=f"적재: {b['used_L']:,}mm",
                                           showarrow=False, font=dict(color=MAIN_COLOR, size=13))
                    if cur_max_l - b['used_L'] > 100:
                        fig.add_annotation(x=b['used_L'] + (cur_max_l - b['used_L'])/2, y=max_width+300,
                                           text=f"잔여: {cur_max_l - b['used_L']:,}mm",
                                           showarrow=False, font=dict(color=ALERT_COLOR, size=13))
                    fig.update_layout(xaxis=dict(visible=False, range=[-200, max_40_len+400]),
                                      yaxis=dict(visible=False, range=[-300, max_width+750]),
                                      height=280, margin=dict(l=10, r=10, t=30, b=10),
                                      paper_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig, use_container_width=True, key=f"plot_{b['id']}")
                st.markdown('</div>', unsafe_allow_html=True)

            # --- 최종 결과 다운로드 ---
            st.markdown("---")

            mapping = {}
            for bx in bins:
                for r in bx['rows']:
                    for fp in r['items']:
                        for sub in fp['sub_items']:
                            mapping[sub['row_idx']] = bx['c_label']

            if file.name.endswith('.xlsx'):
                file.seek(0)
                wb = load_workbook(file)
                ws = wb[selected_sheet]

                target_col = ws.max_column + 1
                target_letter = ws.cell(row=1, column=target_col).column_letter
                ws.cell(row=4, column=target_col).value = "배정 컨테이너"

                if target_col > 1:
                    src_cell = ws.cell(row=4, column=target_col - 1)
                    dst_cell = ws.cell(row=4, column=target_col)
                    dst_cell.font = copy(src_cell.font)
                    dst_cell.fill = copy(src_cell.fill)
                    dst_cell.border = copy(src_cell.border)
                    dst_cell.alignment = copy(src_cell.alignment)
                    dst_cell.number_format = src_cell.number_format

                for r_idx, label in mapping.items():
                    excel_row = int(r_idx) + 1
                    ws.cell(row=excel_row, column=target_col).value = label
                    if target_col > 1:
                        src_cell = ws.cell(row=excel_row, column=target_col - 1)
                        dst_cell = ws.cell(row=excel_row, column=target_col)
                        dst_cell.font = copy(src_cell.font)
                        dst_cell.fill = copy(src_cell.fill)
                        dst_cell.border = copy(src_cell.border)
                        dst_cell.alignment = copy(src_cell.alignment)
                        dst_cell.number_format = src_cell.number_format

                ws.column_dimensions[target_letter].width = 25

                output = io.BytesIO()
                wb.save(output)

                st.download_button(
                    label="📥 최종 결과 다운로드 (원본 양식 유지)",
                    data=output.getvalue(),
                    file_name="CLP_RESULT_FINAL.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            else:
                export_df = raw_full.copy()
                target_col = export_df.shape[1]
                export_df[target_col] = ""
                export_df.iloc[3, target_col] = "배정 컨테이너"

                for r_idx, label in mapping.items():
                    export_df.iloc[int(r_idx), target_col] = label

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    export_df.to_excel(writer, index=False, header=False)

                st.download_button(
                    label="📥 최종 결과 다운로드",
                    data=output.getvalue(),
                    file_name="CLP_RESULT_FINAL.xlsx",
                    use_container_width=True
                )

    except Exception as e:
        st.error(f"데이터 처리 중 오류 발생: {e}")
        import traceback
        st.code(traceback.format_exc())
