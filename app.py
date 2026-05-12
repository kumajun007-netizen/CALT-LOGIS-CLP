import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import math
import os
import io
from openpyxl import load_workbook
from copy import copy

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

    # 💡 수정: 최신 양식(B, I, J, L, N) 기준 열 안내
    with st.expander("📄 엑셀 업로드 표준 규격 (열 고정)", expanded=True):
        st.markdown(f"""
        <table class="guide-table">
            <tr><th>항목</th><th>엑셀 열</th><th>구분</th></tr>
            <tr class="essential"><td>No.of PKG</td><td>B 열</td><td>[필수]</td></tr>
            <tr class="essential"><td>LENGTH (L)</td><td>J 열</td><td>[필수]</td></tr>
            <tr class="essential"><td>WIDTH (W)</td><td>L 열</td><td>[필수]</td></tr>
            <tr class="essential"><td>HEIGHT (H)</td><td>N 열</td><td>[필수]</td></tr>
            <tr><td>NET WEIGHT</td><td>I 열</td><td>선택</td></tr>
            <tr><td>ITEM</td><td>D 열</td><td>참고</td></tr>
            <tr><td>DESCRIPTION</td><td>E 열</td><td>참고</td></tr>
        </table>
        <p style='font-size:11px; color:gray; margin-top:10px;'>* 데이터는 6번째 줄부터 시작하는 것을 권장합니다.</p>
        """, unsafe_allow_html=True)

    # 💡 신규 추가: 표준 양식 다운로드 기능
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
        "Dimension H (mm)": [2300]
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

    with st.expander("🛠 적재 로직", expanded=True):
        use_balancing = st.checkbox("⚖️ 균분적재 (Balancing)", value=False)
        allow_stacking = st.checkbox("🏢 다단적재 (Stacking)", value=False)
    
    if st.button("🔄 AI 재계산 실행"):
        st.session_state['manual_mode'] = False

# --- 짐 배치 로직 (이전과 동일하지만 입력 정렬 및 혼적 유지) ---
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

def calculate_expert_packing(df, max_40_wt, max_40_len, max_20_wt, max_20_len, max_dry_h, max_hc_h, use_balancing):
    # ── Step 1: 치수 정규화
    raw_pieces = []
    for _, row in df.iterrows():
        l, w, h, weight = int(row['L'] + 0.5), int(row['W'] + 0.5), int(row['H'] + 0.5), int(row['WEIGHT'] + 0.5)
        raw_pieces.append({**row.to_dict(), 'L': l, 'W': w, 'H': h, 'WEIGHT': weight})

    # ── Step 2: 동일 크기 그룹핑 + 나란히 최적 방향 결정
    # 같은 치수의 화물 N개를 배치할 때 컨테이너 L 소비가 최소인 방향 선택
    size_groups = {}
    for p in raw_pieces:
        # HC 필요 여부를 키에 포함 → HC끼리, Dry끼리 각각 나란히 방향 최적화
        is_hc = p['H'] > max_dry_h
        key = (min(p['L'], p['W']), max(p['L'], p['W']), is_hc)
        size_groups.setdefault(key, []).append(p)

    all_pieces = []
    for (s, lg, _), items in size_groups.items():
        n = len(items)
        # 방향 A: el=s(짧은쪽→컨테이너L), ew=lg(긴쪽→컨테이너W)
        # 방향 B: el=lg(긴쪽→컨테이너L), ew=s(짧은쪽→컨테이너W) ← 회전
        can_a = lg <= 2350
        can_b = s <= 2350
        if can_a and can_b:
            slots_a = max(1, int(2350 // lg))
            slots_b = max(1, int(2350 // s))
            L_used_a = math.ceil(n / slots_a) * s
            L_used_b = math.ceil(n / slots_b) * lg
            # 나란히 2개 이상 가능하고 L 소비가 더 적은 방향으로
            use_el, use_ew = (lg, s) if (L_used_b < L_used_a and slots_b >= 2) else (s, lg)
        elif can_b:
            use_el, use_ew = lg, s
        else:
            use_el, use_ew = s, lg

        for p in items:
            all_pieces.append({**p, 'L': use_el, 'W': use_ew})
    # ── Step 3: 화물 등급 분류 (비용 절감 배치 원칙)
    # 우선순위: FR(OW/OH) → HC → Dry
    # 각 상위 등급 bin의 남는 공간에 하위 등급 화물을 채워 컨테이너 수 최소화
    sk = lambda x: (-x['W'], -x['H'], -x['L'], x['GROUP'])
    fr_pieces  = sorted([p for p in all_pieces if p['W'] > 2350 or p['H'] > max_hc_h],  key=sk)
    hc_pieces  = sorted([p for p in all_pieces if p['W'] <= 2350 and max_dry_h < p['H'] <= max_hc_h], key=sk)
    dry_pieces = sorted([p for p in all_pieces if p['W'] <= 2350 and p['H'] <= max_dry_h], key=sk)

    def _pack_group(pieces, c_no_start):
        if not pieces:
            return [], c_no_start
        c_no = c_no_start
        group_bins = []
        if use_balancing:
            from collections import Counter
            real_total_l = 0
            for (el, ew), cnt in Counter((p['L'], p['W']) for p in pieces).items():
                slots = max(1, int(2350 // ew))
                rows  = math.ceil(cnt / slots)
                real_total_l += rows * el
            total_w = sum(p['WEIGHT'] for p in pieces)
            est = max(1, math.ceil(real_total_l / max_40_len), math.ceil(total_w / max_40_wt))
            for _ in range(est):
                group_bins.append({'id': c_no, 'rows': [], 'used_L': 0, 'total_W': 0,
                                   'max_W': 0, 'max_H': 0, 'stacked_items': [], 'groups': set()})
                c_no += 1
        for piece in pieces:
            placed = False
            if use_balancing:
                group_bins.sort(key=lambda b: (0 if piece['GROUP'] in b['groups'] else 1, b['used_L']))
            for b in group_bins:
                pack_items_into_bin([piece], b, max_40_wt, max_40_len)
                if piece in b.get('stacked_items', []) or any(piece in r['items'] for r in b['rows']):
                    placed = True; break
            if not placed:
                new_bin = {'id': c_no, 'rows': [], 'used_L': 0, 'total_W': 0,
                           'max_W': 0, 'max_H': 0, 'stacked_items': [], 'groups': set()}
                pack_items_into_bin([piece], new_bin, max_40_wt, max_40_len)
                group_bins.append(new_bin); c_no += 1
        group_bins = [b for b in group_bins if b['used_L'] > 0 or b.get('stacked_items')]
        return group_bins, c_no

    def _fill_gaps(bins, candidates):
        """bin의 남는 L 공간에 candidates를 채워 넣고, 배치된 화물을 제거한 나머지 반환"""
        remaining = list(candidates)
        for b in sorted(bins, key=lambda x: x['used_L']):  # 덜 찬 bin부터
            for piece in list(remaining):
                before = sum(len(r['items']) for r in b['rows']) + len(b.get('stacked_items', []))
                pack_items_into_bin([piece], b, max_40_wt, max_40_len)
                after  = sum(len(r['items']) for r in b['rows']) + len(b.get('stacked_items', []))
                if after > before:
                    remaining.remove(piece)
        return remaining

    # ① FR 화물 배치 → 남는 공간에 HC, Dry 순으로 채우기
    fr_bins, c_no  = _pack_group(fr_pieces, 1)
    remaining_hc   = _fill_gaps(fr_bins, hc_pieces)
    remaining_dry  = _fill_gaps(fr_bins, dry_pieces)

    # ② HC 화물 배치 → 남는 공간에 Dry 채우기 → 20ft Dry 활용 극대화
    hc_bins, c_no  = _pack_group(remaining_hc, c_no)
    remaining_dry  = _fill_gaps(hc_bins, remaining_dry)

    # ③ 나머지 Dry 화물 배치 (L·중량 기준 20ft Dry 자동 판별)
    dry_bins, c_no = _pack_group(remaining_dry, c_no)

    bins = fr_bins + hc_bins + dry_bins
    for idx, b in enumerate(bins):
        b['id'] = idx + 1
    return apply_labels(bins, max_20_len, max_20_wt, max_40_len - 430, max_dry_h, max_hc_h)

# --- 3. 메인 화면 로직 ---
def reset_data():
    if 'bins' in st.session_state: del st.session_state['bins']
    if 'manual_mode' in st.session_state: del st.session_state['manual_mode']

st.markdown("### 📤 패킹리스트 업로드 (드래그 앤 드롭)")
file = st.file_uploader("이곳에 파일을 끌어다 놓으세요.", type=['csv', 'xlsx'], on_change=reset_data)

if file is not None:
    try:
        # --- 엑셀/CSV 원본 읽기 ---
        # 엑셀의 첫 번째 시트가 비어 있는 경우가 있어, 데이터가 있는 시트를 자동으로 선택합니다.
        selected_sheet = None

        if file.name.endswith('.xlsx'):
            file.seek(0)
            excel_sheets = pd.read_excel(file, sheet_name=None, header=None)

            raw_full = None
            for sheet_name, sheet_df in excel_sheets.items():
                if not sheet_df.dropna(how='all').empty:
                    raw_full = sheet_df
                    selected_sheet = sheet_name
                    break

            if raw_full is None:
                st.error("엑셀에서 데이터가 있는 시트를 찾지 못했습니다.")
                st.stop()

            st.info(f"읽은 시트: {selected_sheet}")

        else:
            file.seek(0)
            raw_full = pd.read_csv(file, header=None)

        # ── 헤더 행 자동 감지 ──────────────────────────────────────────
        # B(1), J(9), L(11), N(13) 열이 모두 숫자인 첫 번째 행을 데이터 시작점으로 결정
        # → 화주마다 헤더 위치가 달라도 자동 대응
        data_start_row = None
        for scan_idx in range(len(raw_full)):
            b_val = raw_full.iloc[scan_idx, 1]
            j_val = raw_full.iloc[scan_idx, 9]
            l_val = raw_full.iloc[scan_idx, 11]
            n_val = raw_full.iloc[scan_idx, 13]
            try:
                if (pd.notna(b_val) and pd.notna(j_val) and pd.notna(l_val) and pd.notna(n_val)
                        and float(str(j_val).replace(',', '')) > 0
                        and float(str(l_val).replace(',', '')) > 0
                        and float(str(n_val).replace(',', '')) > 0):
                    data_start_row = scan_idx
                    break
            except (ValueError, TypeError):
                continue

        if data_start_row is None:
            st.error("⚠️ 데이터 행을 찾지 못했습니다. B열(PKG), J열(L), L열(W), N열(H)에 숫자값이 있는지 확인하세요.")
            st.stop()

        st.caption(f"📍 데이터 시작 행: 엑셀 {data_start_row + 1}행 (시트: {selected_sheet})")

        raw_process = (
            raw_full
            .iloc[data_start_row:]
            .reset_index(drop=False)
            .rename(columns={'index': 'orig_idx'})
        )

        p_data = []
        for i in range(len(raw_process)):
            row = raw_process.iloc[i]

            # 열 고정 매핑: B=1, D=3, E=4, H=7(NetWt), J=9(L), L=11(W), N=13(H)
            pkg_v = str(row[1]).replace('00:00:00', '').replace('.0', '').strip() if pd.notna(row[1]) else None
            l_v = clean_num(row[9])       # J열: LENGTH
            w_v = clean_num(row[11])      # L열: WIDTH
            h_v = clean_num(row[13])      # N열: HEIGHT
            weight_v = clean_num(row[8])  # I열: GROSS WEIGHT

            # PKG NO와 치수(LWH)만 필수값으로 체크
            if not pkg_v or pkg_v.lower() in ['nan', 'none', '.', ''] or l_v == 0 or w_v == 0 or h_v == 0:
                continue

            p_data.append({
                'PKG NO': pkg_v,
                'GROUP': str(row[4]) if pd.notna(row[4]) else "-",   # E열: Description (그룹핑 기준)
                'ITEM': str(row[3]) if pd.notna(row[3]) else "-",    # D열: ITEM
                'DESC': str(row[4]) if pd.notna(row[4]) else "-",    # E열: Description
                'L': l_v,
                'W': w_v,
                'H': h_v,
                'WEIGHT': weight_v,
                'STACK_OK': False,
                # 원본 엑셀 행 위치 보존 (다운로드 시 결과 기입용)
                'row_idx': int(row['orig_idx'])
            })

        df = pd.DataFrame(p_data)
        if df.empty:
            st.warning("⚠️ 지정된 열(B, J, L, N)에서 필수 데이터를 찾을 수 없습니다. 양식을 확인하세요.")
        else:
            if 'bins' not in st.session_state or not st.session_state.get('manual_mode', False):
                st.session_state.bins = calculate_expert_packing(df, max_40_wt, max_40_len, max_20_wt, max_20_len, max_dry_h, max_hc_h, use_balancing)
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
                            pack_items_into_bin([item], new_bins_dict[t_id], max_40_wt, max_40_len)
                        st.session_state.bins = apply_labels(sorted(list(new_bins_dict.values()), key=lambda x: x['id']), max_20_len, max_20_wt, max_40_len-430, max_dry_h, max_hc_h)
                        st.rerun()

                with st.expander("👁️ 적재 단면도 및 제원 확인", expanded=False):
                    cur_max_l = max_20_len if "20ft" in b['c_label'] else max_40_len
                    cur_max_w = max_20_wt if "20ft" in b['c_label'] else max_40_wt
                    cur_max_h = max_hc_h if "HC" in b['c_label'] else max_dry_h
                    used_width = max([r['used_W'] for r in b['rows']] + [0])
                    max_stacked_h = max([s['H'] for s in b.get('stacked_items', [])] + [0])
                    used_height = b['max_H'] + max_stacked_h if b.get('stacked_items') else b['max_H']
                    
                    c1, c2, c3, c4 = st.columns(4)
                    c1.markdown(f"**📏 길이:** {b['used_L']:,}/{cur_max_l:,}mm"); c1.progress(min(1.0, b['used_L']/cur_max_l))
                    c2.markdown(f"**⚖️ 중량:** {b['total_W']:,}/{cur_max_w:,}kg"); c2.progress(min(1.0, b['total_W']/cur_max_w))
                    
                    if used_width > 2350: c3.markdown(f"**↔️ 폭:** {used_width:,}/2350mm <span style='color:{ALERT_COLOR};font-weight:bold;'>[OW +{used_width-2350:,}]</span>", unsafe_allow_html=True)
                    else: c3.markdown(f"**↔️ 폭:** {used_width:,}/2350mm"); c3.progress(min(1.0, used_width/2350))
                    
                    if used_height > cur_max_h: c4.markdown(f"**↕️ 높이:** {used_height:,}/{cur_max_h:,}mm <span style='color:{ALERT_COLOR};font-weight:bold;'>[OH +{used_height-cur_max_h:,}]</span>", unsafe_allow_html=True)
                    else: c4.markdown(f"**↕️ 높이:** {used_height:,}/{cur_max_h:,}mm"); c4.progress(min(1.0, used_height/cur_max_h))
                    
                    # 범례 표시
                    st.markdown(
                        f"<span style='background:#e67e22;padding:2px 8px;border-radius:4px;color:white;font-size:12px;'>■ HC 필요 (H>{max_dry_h}mm)</span>"
                        f"&nbsp;&nbsp;"
                        f"<span style='background:{ACCENT_COLOR};padding:2px 8px;border-radius:4px;color:white;font-size:12px;'>■ DRY 가능 (H≤{max_dry_h}mm)</span>",
                        unsafe_allow_html=True
                    )
                    fig = go.Figure()
                    fig.add_shape(type="rect", x0=b['used_L'], y0=0, x1=cur_max_l, y1=2350, fillcolor="#e1e4e8", opacity=0.4, line_width=0)
                    fig.add_shape(type="rect", x0=0, y0=0, x1=cur_max_l, y1=2350, line=dict(color=MAIN_COLOR, width=2))
                    cx = 0
                    for r in b['rows']:
                        cy = (2350 - r['used_W']) / 2
                        for item in r['items']:
                            # CLP 참고: HC 필요 화물(H > DRY 기준) → 주황, Dry 가능 → 파란색
                            item_color = "#e67e22" if item['H'] > max_dry_h else ACCENT_COLOR
                            fig.add_shape(type="rect", x0=cx, y0=cy, x1=cx+item['L'], y1=cy+item['W'], fillcolor=item_color, opacity=0.85, line=dict(color="white", width=1))
                            fig.add_annotation(x=cx+item['L']/2, y=cy+item['W']/2,
                                text=f"{item['PKG NO']}<br><sub>H{item['H']}</sub>",
                                showarrow=False, font=dict(color="white", size=9))
                            cy += item['W']
                        cx += r['max_L']
                    fig.add_shape(type="line", x0=b['used_L'], y0=-200, x1=b['used_L'], y1=2800, line=dict(color=ALERT_COLOR, width=2, dash="dash"))
                    if b['used_L'] > 100: fig.add_annotation(x=b['used_L']/2, y=2650, text=f"적재: {b['used_L']:,}mm", showarrow=False, font=dict(color=MAIN_COLOR, size=13, weight="bold"))
                    if cur_max_l - b['used_L'] > 100: fig.add_annotation(x=b['used_L'] + (cur_max_l - b['used_L'])/2, y=2650, text=f"잔여: {cur_max_l - b['used_L']:,}mm", showarrow=False, font=dict(color=ALERT_COLOR, size=13, weight="bold"))
                    fig.update_layout(xaxis=dict(visible=False, range=[-200, max_40_len+400]), yaxis=dict(visible=False, range=[-300, 3100]), height=280, margin=dict(l=10, r=10, t=30, b=10), paper_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig, use_container_width=True, key=f"plot_{b['id']}")
                st.markdown('</div>', unsafe_allow_html=True)

            # --- 최종 결과 다운로드: 원본 엑셀 양식 유지 ---
            st.markdown("---")

            mapping = {
                item['row_idx']: bx['c_label']
                for bx in bins
                for item in (
                    [i for r in bx['rows'] for i in r['items']]
                    + bx.get('stacked_items', [])
                )
            }

            if file.name.endswith('.xlsx'):
                file.seek(0)
                wb = load_workbook(file)

                # 위에서 자동 선택한 데이터 시트에 결과만 추가
                ws = wb[selected_sheet]

                target_col = ws.max_column + 1
                target_letter = ws.cell(row=1, column=target_col).column_letter

                # 제목 입력: 원본 양식 기준 4행에 표시
                ws.cell(row=4, column=target_col).value = "배정 컨테이너"

                # 제목 셀 스타일 복사
                if target_col > 1:
                    src_cell = ws.cell(row=4, column=target_col - 1)
                    dst_cell = ws.cell(row=4, column=target_col)
                    dst_cell.font = copy(src_cell.font)
                    dst_cell.fill = copy(src_cell.fill)
                    dst_cell.border = copy(src_cell.border)
                    dst_cell.alignment = copy(src_cell.alignment)
                    dst_cell.number_format = src_cell.number_format

                # 배정 결과 입력
                # row_idx는 pandas 기준 0부터 시작하므로 엑셀 행 번호는 +1
                for r_idx, label in mapping.items():
                    excel_row = int(r_idx) + 1
                    ws.cell(row=excel_row, column=target_col).value = label

                    # 왼쪽 셀의 양식 복사
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

    except Exception as e: st.error(f"데이터 처리 중 오류 발생: {e}")
