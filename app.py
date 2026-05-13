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
    with st.expander("📄 엑셀 업로드 표준 규격", expanded=True):
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
            <tr><td>REMARK</td><td>O 열</td><td>선택</td></tr>
            <tr><td>LOAD</td><td>P 열</td><td>선택</td></tr>
        </table>
        <div style='font-size:10px; color:#555; margin-top:8px; line-height:1.6;'>
            <b>REMARK 키워드</b><br>
            · <b>BOX</b> : Q'ty = 실제 박스 수<br>
            · <b>2단 / 3단</b> : 다단 적재 허용<br>
            · 복합 예시 : <code>BOX,2단</code><br><br>
            <b>LOAD 키워드</b><br>
            · <b>FORK_W</b> : 긴쪽→W (LSE, 포크 구멍 긴쪽)<br>
            · <b>FORK_L</b> : 긴쪽→L (FR 빈공간 채우기 등)<br>
            · <b>4WAY</b> : 4방향 자유 (포크 구멍 4면)<br>
            · 비워두면 사이드바 모드 설정 따라감
        </div>
        <p style='font-size:11px; color:gray; margin-top:10px;'>* 데이터 시작 행은 자동으로 감지됩니다.</p>
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
            "REMARK": [""],
            "LOAD": [""]
        }
        df_template = pd.DataFrame(template_data)
        tow = io.BytesIO()
        with pd.ExcelWriter(tow, engine='openpyxl') as writer:
            df_template.to_excel(writer, index=False, sheet_name='PackingList')
            # 사용설명서 시트
            guide_rows = [
                ["■ REMARK 키워드 (O열)", ""],
                ["BOX",    "Q'ty = 실제 박스 수 (동일 화물을 Q'ty개로 복제)"],
                ["2단",    "해당 화물 2단 적재 허용 (치수 무관)"],
                ["3단",    "해당 화물 3단 적재 허용 (치수 무관)"],
                ["BOX,2단","복합 적용 예시 (쉼표로 구분)"],
                ["", ""],
                ["■ LOAD 키워드 (P열)", ""],
                ["FORK_W", "긴쪽 → W방향 고정 (LSE 화물 기본, 포크 구멍이 긴쪽에만 있는 경우)"],
                ["FORK_L", "긴쪽 → L방향 고정 (FR 빈공간 채우기 등 특수 케이스)"],
                ["4WAY",   "4방향 자유 배치 (포크 구멍 4면, FR 대형화물)"],
                ["(비워둠)","사이드바 적재 방식 모드 설정을 따라감"],
                ["", ""],
                ["■ 적재 방식 (사이드바 선택)", ""],
                ["일반",   "L 소비 최소 자동 최적화"],
                ["LSE",    "전체 화물에 FORK_W 적용 (개별 LOAD 키워드 있으면 개별 우선)"],
            ]
            df_guide = pd.DataFrame(guide_rows, columns=["키워드", "설명"])
            df_guide.to_excel(writer, index=False, sheet_name='사용설명서')
            ws_g = writer.sheets['사용설명서']
            ws_g.column_dimensions['A'].width = 18
            ws_g.column_dimensions['B'].width = 60
        st.download_button(
            label="📥 신규 화주용 양식 다운로드",
            data=tow.getvalue(),
            file_name="CALT_CLP_TEMPLATE.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.header("⚙️ 배정 옵션 설정")
    load_mode = st.radio(
        "📐 적재 방식",
        ["일반 (효율 최적)", "LSE (포크방향 고정)"],
        index=0, key="load_mode",
        help="LSE: 긴쪽→W 고정 (포크 구멍이 긴쪽에만 있는 화물) / 일반: L 소비 최소 자동 최적화"
    )
    lse_mode = (load_mode == "LSE (포크방향 고정)")
    st.markdown("---")
    with st.expander("⚖️ 컨테이너 제원", expanded=True):
        st.markdown("**🟦 20ft DRY**")
        max_20_wt  = st.number_input("최대 중량 (kg)",  15000, 40000, 28250, key="i_20wt")
        max_20_len = st.number_input("최대 길이 (mm)",  5500,  6500,  5900,  key="i_20len")
        st.markdown("---")
        st.markdown("**🟫 40ft DRY / HC**")
        max_40_wt  = st.number_input("최대 중량 (kg)",  20000, 40000, 29500, key="i_40wt")
        max_40_len = st.number_input("최대 길이 (mm)",  11000, 13000, 11900, key="i_40len")
        max_dry_h  = st.number_input("DRY 내부 높이 (mm)", 2000, 3000, 2370, key="i_dryh")
        max_hc_h   = st.number_input("HC 내부 높이 (mm)",  2000, 3500, 2670, key="i_hch")
        st.markdown("---")
        st.markdown("**🔴 20ft FR (Flat Rack)**")
        max_fr20_wt  = st.number_input("최대 중량 (kg)",  10000, 40000, 25000, key="i_fr20wt")
        max_fr20_len = st.number_input("최대 길이 (mm)",  4000,  7000,  5600,  key="i_fr20len")
        max_fr_h     = st.number_input("내부 높이 (mm)",  1000,  3000,  2260,  key="i_frh")
        max_fr_w     = st.number_input("기둥간 폭 (mm)",  1000,  3000,  2080,  key="i_frw")
        st.markdown("---")
        st.markdown("**🔴 40ft FR (Flat Rack)**")
        max_fr40_wt  = st.number_input("최대 중량 (kg)",  10000, 45000, 30000, key="i_fr40wt")
        max_fr40_len = st.number_input("최대 길이 (mm)",  8000,  14000, 11600, key="i_fr40len")

    if st.button("🔄 AI 재계산 실행"):
        st.session_state['manual_mode'] = False

# --- 짐 배치 로직 (이전과 동일하지만 입력 정렬 및 혼적 유지) ---
def pack_items_into_bin(pieces, b, max_40_wt, max_40_len, max_h=None, allow_free_stack=False):
    """base(rows) 배치 전용. 스태킹은 post_stack_bins() 후처리에서 처리."""
    for piece in pieces:
        placed = False
        row_found = False
        for r in b['rows']:
            temp_max_L = max(r['max_L'], piece['L'])
            if r['used_W'] + piece['W'] <= 2340 and b['used_L'] + (temp_max_L - r['max_L']) <= max_40_len and b['total_W'] + piece['WEIGHT'] <= max_40_wt:
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

def apply_labels(bins, max_20_len, max_20_wt, max_dry_h, max_hc_h, max_fr20_len, max_fr20_wt, max_fr40_len, max_fr40_wt):
    for b in bins:
        # 수동 재계산으로 타입이 강제 지정된 bin은 레이블 유지
        if b.get('forced_label'):
            continue
        is_20ft_size  = b['used_L'] <= max_20_len  and b['total_W'] <= max_20_wt
        is_20fr_size  = b['used_L'] <= max_fr20_len and b['total_W'] <= max_fr20_wt
        is_ow = b['max_W'] > 2340
        # OH 판단: 치수별 정확한 누적 높이 체크
        _base_items = [i for r in b['rows'] for i in r['items']]
        _stk_items  = b.get('stacked_items', [])
        if _stk_items:
            from collections import Counter as _C
            _base_cnt = _C(i['H'] for i in _base_items)
            _stk_cnt  = _C(s['H'] for s in _stk_items)
            _effective_h = 0
            for h_val, stk_n in _stk_cnt.items():
                base_n = _base_cnt.get(h_val, max(1, len(_base_items)))
                layers = 1 + math.ceil(stk_n / max(1, base_n))
                _effective_h = max(_effective_h, h_val * layers)
            _effective_h = max(_effective_h, b['max_H'])
        else:
            _effective_h = b['max_H']
        is_oh = _effective_h > max_hc_h
        tags = []
        if is_oh: tags.append("OH")
        if is_ow: tags.append("OW")
        if tags:
            base = "20ft Flat Rack" if is_20fr_size else "40ft Flat Rack"
            b['c_label'] = f"{base} [{' + '.join(tags)}] #{b['id']}"
        else:
            # 단적 포함 실제 높이 계산 → HC 필요 여부 정확 판단
            base_items = [i for r in b['rows'] for i in r['items']]
            stk_items  = b.get('stacked_items', [])
            if stk_items and base_items:
                import math as _math
                stk_layers   = _math.ceil(len(stk_items) / max(1, len(base_items)))
                max_base_h   = max((i['H'] for i in base_items), default=0)
                max_stk_h    = max((s['H'] for s in stk_items), default=0)
                effective_h  = max_base_h + stk_layers * max_stk_h
            else:
                effective_h = b['max_H']
            # effective_h > DRY 한도면 HC 필요, 그렇지 않으면 DRY
            if effective_h > max_dry_h: base = "40ft HC"
            else: base = "20ft Dry" if is_20ft_size else "40ft Dry"
            b['c_label'] = f"{base} #{b['id']}"
    return bins

def calculate_expert_packing(df, max_40_wt, max_40_len, max_20_wt, max_20_len, max_dry_h, max_hc_h, lse_mode=False):
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
        can_a = lg <= 2340   # 방향 A: el=s(짧은쪽→L), ew=lg(긴쪽→W)
        can_b = s  <= 2340   # 방향 B: el=lg(긴쪽→L), ew=s(짧은쪽→W)

        for p in items:
            fork_dir = p.get('FORK_DIR')  # 개별 LOAD 키워드 우선
            is_fr    = (p['W'] > 2340 or p['H'] > max_hc_h)  # FR 대형화물 여부

            # ── 방향 결정 우선순위 ──────────────────────────────
            # 1순위: 개별 FORK_DIR 키워드
            # 2순위: FR 대형화물 → 4WAY (효율 최적)
            # 3순위: FR bin 채우기용 DRY/HC 화물 → FORK_L (긴쪽→L)
            # 4순위: LSE 모드 → FORK_W (긴쪽→W)
            # 5순위: 일반 → L 소비 최소 자동

            if fork_dir == 'FORK_W' or (fork_dir is None and lse_mode and not is_fr):
                # 긴쪽 → W (LSE 기본, 포크가 긴쪽 구멍으로 진입)
                use_el, use_ew = (s, lg) if can_a else (lg, s)

            elif fork_dir == 'FORK_L':
                # 긴쪽 → L
                use_el, use_ew = (lg, s) if can_b else (s, lg)

            elif fork_dir == '4WAY' or is_fr:
                # 4방향 자유 → L 소비 최소 (일반 최적화)
                if can_a and can_b:
                    slots_a = max(1, int(2340 // lg))
                    slots_b = max(1, int(2340 // s))
                    L_a = math.ceil(n / slots_a) * s
                    L_b = math.ceil(n / slots_b) * lg
                    use_el, use_ew = (lg, s) if (L_b < L_a and slots_b >= 2) else (s, lg)
                elif can_b: use_el, use_ew = lg, s
                else:       use_el, use_ew = s, lg

            else:
                # 일반 모드 기본: L 소비 최소
                if can_a and can_b:
                    slots_a = max(1, int(2340 // lg))
                    slots_b = max(1, int(2340 // s))
                    L_a = math.ceil(n / slots_a) * s
                    L_b = math.ceil(n / slots_b) * lg
                    use_el, use_ew = (lg, s) if (L_b < L_a and slots_b >= 2) else (s, lg)
                elif can_b: use_el, use_ew = lg, s
                else:       use_el, use_ew = s, lg

            all_pieces.append({**p, 'L': use_el, 'W': use_ew})
    # ── Step 3: 화물 등급 분류 (비용 절감 배치 원칙)
    # 우선순위: FR(OW/OH) → HC → Dry
    # 각 상위 등급 bin의 남는 공간에 하위 등급 화물을 채워 컨테이너 수 최소화
    sk = lambda x: (-x['W'], -x['H'], -x['L'], x['GROUP'])
    fr_pieces  = sorted([p for p in all_pieces if p['W'] > 2340 or p['H'] > max_hc_h],  key=sk)
    hc_pieces  = sorted([p for p in all_pieces if p['W'] <= 2340 and max_dry_h < p['H'] <= max_hc_h], key=sk)
    dry_pieces = sorted([p for p in all_pieces if p['W'] <= 2340 and p['H'] <= max_dry_h], key=sk)

    def _post_stack_bins(group_bins, allow_free):
        """base 배치 완료 후 REMARK 단수 / 일반 자유 스태킹 후처리."""
        for b in group_bins:
            from collections import defaultdict
            all_base = [item for r in b['rows'] for item in r['items']]
            if not all_base:
                continue

            if allow_free:
                # ── 일반 모드 자유 스태킹: 높이 한도 내 최대 단수 계산
                groups = defaultdict(list)
                for item in all_base:
                    groups[(item['L'], item['W'], item['H'])].append(item)
                for (l, w, h), items in groups.items():
                    max_layers = int(max_hc_h / h) if h > 0 else 1
                    if max_layers <= 1:
                        continue
                    n = len(items)
                    base_keep = math.ceil(n / max_layers)
                    to_stack = items[base_keep:]
                    for item in to_stack:
                        for r in b['rows']:
                            if item in r['items']:
                                r['items'].remove(item)
                                r['used_W'] -= item['W']
                                break
                        b['stacked_items'].append(item)
            else:
                # ── REMARK 단수 지정 스태킹
                groups = defaultdict(list)
                for item in all_base:
                    ms = item.get('MAX_STK', 1)
                    if ms > 1:
                        groups[(item['L'], item['W'], item['H'], ms)].append(item)
                for (l, w, h, max_stk), items in groups.items():
                    # 높이 한도 적용
                    max_layers_h = int(max_hc_h / h) if h > 0 else max_stk
                    effective_stk = min(max_stk, max_layers_h)
                    if effective_stk <= 1:
                        continue
                    n = len(items)
                    base_keep = math.ceil(n / effective_stk)
                    to_stack = items[base_keep:]
                    for item in to_stack:
                        for r in b['rows']:
                            if item in r['items']:
                                r['items'].remove(item)
                                r['used_W'] -= item['W']
                                break
                        b['stacked_items'].append(item)

            # 빈 row 제거 + 통계 재계산
            b['rows'] = [r for r in b['rows'] if r['items']]
            for r in b['rows']:
                r['max_L']  = max(i['L'] for i in r['items'])
                r['used_W'] = sum(i['W'] for i in r['items'])
            b['used_L'] = sum(r['max_L'] for r in b['rows'])
            b['max_H']  = max((i['H'] for r in b['rows'] for i in r['items']), default=0)
            b['max_W']  = max((i['W'] for r in b['rows'] for i in r['items']), default=0)
            b['total_W'] = sum(i['WEIGHT'] for r in b['rows'] for i in r['items'])                          + sum(s['WEIGHT'] for s in b.get('stacked_items', []))

    def _pack_group(pieces, c_no_start, p_max_wt=None, p_max_len=None):
        _wt  = p_max_wt  or max_40_wt
        _len = p_max_len or max_40_len
        if not pieces:
            return [], c_no_start
        c_no = c_no_start
        group_bins = []
        for piece in pieces:
            placed = False
            for b in group_bins:
                pack_items_into_bin([piece], b, _wt, _len)
                if any(piece in r['items'] for r in b['rows']):
                    placed = True; break
            if not placed:
                new_bin = {'id': c_no, 'rows': [], 'used_L': 0, 'total_W': 0,
                           'max_W': 0, 'max_H': 0, 'stacked_items': [], 'groups': set()}
                pack_items_into_bin([piece], new_bin, _wt, _len)
                group_bins.append(new_bin); c_no += 1
        group_bins = [b for b in group_bins if b['used_L'] > 0]
        # 후처리: base 배치 완료 후 스태킹 적용
        _post_stack_bins(group_bins, not lse_mode)
        return group_bins, c_no

    def _fill_gaps(bins, candidates, p_max_wt=None, p_max_len=None):
        """bin의 남는 L 공간에 candidates를 채워 넣고, 배치된 화물을 제거한 나머지 반환"""
        _wt  = p_max_wt  or max_40_wt
        _len = p_max_len or max_40_len
        remaining = list(candidates)
        for b in sorted(bins, key=lambda x: x['used_L']):
            for piece in list(remaining):
                before = sum(len(r['items']) for r in b['rows'])
                pack_items_into_bin([piece], b, _wt, _len)
                after  = sum(len(r['items']) for r in b['rows'])
                if after > before:
                    remaining.remove(piece)
        return remaining

    # ① FR 화물 배치 (FR 전용 제원 사용) → 남는 공간에 HC, Dry 순으로 채우기
    fr_bins, c_no  = _pack_group(fr_pieces, 1, p_max_wt=max_fr40_wt, p_max_len=max_fr40_len)
    remaining_hc   = _fill_gaps(fr_bins, hc_pieces,  p_max_wt=max_fr40_wt, p_max_len=max_fr40_len)
    remaining_dry  = _fill_gaps(fr_bins, dry_pieces, p_max_wt=max_fr40_wt, p_max_len=max_fr40_len)

    # ② HC 화물 배치 → 남는 공간에 Dry 채우기 → 20ft Dry 활용 극대화
    hc_bins, c_no  = _pack_group(remaining_hc, c_no)
    remaining_dry  = _fill_gaps(hc_bins, remaining_dry)

    # ③ 나머지 Dry 화물 배치 (L·중량 기준 20ft Dry 자동 판별)
    dry_bins, c_no = _pack_group(remaining_dry, c_no)

    bins = fr_bins + hc_bins + dry_bins
    for idx, b in enumerate(bins):
        b['id'] = idx + 1
    return apply_labels(bins, max_20_len, max_20_wt, max_dry_h, max_hc_h, max_fr20_len, max_fr20_wt, max_fr40_len, max_fr40_wt)

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

            # 열 고정 매핑: B=1, D=3, E=4, F=5(Q'ty), I=8(GrossWt), J=9(L), L=11(W), N=13(H), O=14(REMARK)
            pkg_v    = str(row[1]).replace('00:00:00', '').replace('.0', '').strip() if pd.notna(row[1]) else None
            qty_v    = int(clean_num(row[5])) if clean_num(row[5]) > 0 else 1  # F열: Q'ty
            l_v      = clean_num(row[9])   # J열: LENGTH
            w_v      = clean_num(row[11])  # L열: WIDTH
            h_v      = clean_num(row[13])  # N열: HEIGHT
            weight_v = clean_num(row[8])   # I열: GROSS WEIGHT

            # O열: REMARK, P열: LOAD (없으면 빈 문자열)
            remark_raw = str(row[14]).strip().upper() if (len(row) > 14 and pd.notna(row[14])) else ""
            remark_keys = [k.strip() for k in remark_raw.replace('，', ',').split(',')]
            load_raw    = str(row[15]).strip().upper() if (len(row) > 15 and pd.notna(row[15])) else ""

            is_box   = 'BOX' in remark_keys
            max_stk  = 3 if '3단' in remark_keys else (2 if '2단' in remark_keys else 1)
            stack_ok = max_stk > 1
            # LOAD 키워드: FORK_L / FORK_W / 4WAY (없으면 None → 토글 따라감)
            fork_dir = load_raw if load_raw in ('FORK_L', 'FORK_W', '4WAY') else None

            # PKG NO와 치수(LWH)만 필수값으로 체크
            if not pkg_v or pkg_v.lower() in ['nan', 'none', '.', ''] or l_v == 0 or w_v == 0 or h_v == 0:
                continue

            repeat = qty_v if is_box else 1
            for seq in range(repeat):
                suffix = f"-{seq+1:03d}" if repeat > 1 else ""
                p_data.append({
                    'PKG NO'  : f"{pkg_v}{suffix}",
                    'GROUP'   : str(row[4]) if pd.notna(row[4]) else "-",
                    'ITEM'    : str(row[3]) if pd.notna(row[3]) else "-",
                    'DESC'    : str(row[4]) if pd.notna(row[4]) else "-",
                    'L'       : l_v,
                    'W'       : w_v,
                    'H'       : h_v,
                    'WEIGHT'  : weight_v,
                    'STACK_OK': stack_ok,
                    'MAX_STK' : max_stk,
                    'FORK_DIR': fork_dir,   # FORK_L / FORK_W / 4WAY / None
                    'row_idx' : int(row['orig_idx'])
                })

        df = pd.DataFrame(p_data)
        if df.empty:
            st.warning("⚠️ 지정된 열(B, J, L, N)에서 필수 데이터를 찾을 수 없습니다. 양식을 확인하세요.")
        else:
            if 'bins' not in st.session_state or not st.session_state.get('manual_mode', False):
                st.session_state.bins = calculate_expert_packing(df, max_40_wt, max_40_len, max_20_wt, max_20_len, max_dry_h, max_hc_h, lse_mode)
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

            # 타입별 제원 매핑
            TYPE_SPECS = {
                '20ft Dry': {'max_len': max_20_len,   'max_wt': max_20_wt,   'max_h': max_dry_h, 'label_base': '20ft Dry'},
                '40ft Dry': {'max_len': max_40_len,   'max_wt': max_40_wt,   'max_h': max_dry_h, 'label_base': '40ft Dry'},
                '40ft HC' : {'max_len': max_40_len,   'max_wt': max_40_wt,   'max_h': max_hc_h,  'label_base': '40ft HC'},
                '20ft FR' : {'max_len': max_fr20_len, 'max_wt': max_fr20_wt, 'max_h': max_fr_h,  'label_base': '20ft Flat Rack'},
                '40ft FR' : {'max_len': max_fr40_len, 'max_wt': max_fr40_wt, 'max_h': max_fr_h,  'label_base': '40ft Flat Rack'},
            }
            def _get_cur_type(label):
                if '20ft Dry'       in label: return '20ft Dry'
                if '40ft HC'        in label: return '40ft HC'
                if '40ft Dry'       in label: return '40ft Dry'
                if '20ft Flat Rack' in label: return '20ft FR'
                if '40ft Flat Rack' in label: return '40ft FR'
                return '40ft Dry'

            for b in bins:
                if b['used_L'] == 0 and not b.get('stacked_items'): continue
                st.markdown(f'<div class="container-box">', unsafe_allow_html=True)

                # ── 컨테이너 헤더 + 타입 재계산 UI ──────────────────────
                hcol1, hcol2, hcol3 = st.columns([4, 2, 1])
                hcol1.markdown(f"### 📦 {b['c_label']}")
                cur_type = _get_cur_type(b['c_label'])
                new_type = hcol2.selectbox(
                    "타입 변경", list(TYPE_SPECS.keys()),
                    index=list(TYPE_SPECS.keys()).index(cur_type),
                    key=f"type_sel_{b['id']}", label_visibility="collapsed"
                )
                if hcol3.button("🔄 재계산", key=f"retype_{b['id']}"):
                    spec = TYPE_SPECS[new_type]
                    all_items = [i for r in b['rows'] for i in r['items']] + b.get('stacked_items', [])
                    # 새 제원으로 재배치 시도
                    new_bin = {'id': b['id'], 'rows': [], 'used_L': 0, 'total_W': 0,
                               'max_W': 0, 'max_H': 0, 'stacked_items': [], 'groups': set()}
                    overflow = []
                    for item in all_items:
                        before = sum(len(r['items']) for r in new_bin['rows']) + len(new_bin.get('stacked_items',[]))
                        pack_items_into_bin([item], new_bin, spec['max_wt'], spec['max_len'])
                        after  = sum(len(r['items']) for r in new_bin['rows']) + len(new_bin.get('stacked_items',[]))
                        if after == before:
                            overflow.append(item)
                    # 높이/폭 초과 화물 체크
                    h_exceed = [i for i in (new_bin['stacked_items'] + [x for r in new_bin['rows'] for x in r['items']]) if i['H'] > spec['max_h']]
                    # 레이블 강제 지정 (apply_labels가 덮어쓰지 않도록 forced_label=True)
                    new_bin['c_label']     = f"{spec['label_base']} #{b['id']}"
                    new_bin['forced_label'] = True
                    # 기존 bins에서 해당 bin 교체
                    updated = []
                    for bx in st.session_state.bins:
                        if bx['id'] == b['id']: updated.append(new_bin)
                        else: updated.append(bx)
                    # 초과 화물 → 자동 배치로 새 bin 생성 (apply_labels로 타입 자동 결정)
                    if overflow:
                        max_id = max(bx['id'] for bx in updated)
                        ov_bin = {'id': max_id+1, 'rows': [], 'used_L': 0, 'total_W': 0,
                                  'max_W': 0, 'max_H': 0, 'stacked_items': [], 'groups': set()}
                        for item in overflow:
                            pack_items_into_bin([item], ov_bin, max_40_wt, max_40_len)
                        updated.append(ov_bin)
                        st.warning(f"⚠️ {len(overflow)}개 화물이 {new_type} 제원 초과 → 새 컨테이너 #{max_id+1}로 분리됐습니다.")
                    if h_exceed:
                        st.error(f"🚨 {len(h_exceed)}개 화물(H 초과)이 포함됩니다. CLP 작성 시 확인하세요.")
                    st.session_state.bins = apply_labels(
                        sorted(updated, key=lambda x: x['id']),
                        max_20_len, max_20_wt, max_dry_h, max_hc_h,
                        max_fr20_len, max_fr20_wt, max_fr40_len, max_fr40_wt
                    )
                    st.rerun()
                # 바닥 화물 최소 L×W (하중 위험 판단 기준)
                base_items = [item for r in b['rows'] for item in r['items']]
                base_min_L = min((i['L'] for i in base_items), default=0)
                base_min_W = min((i['W'] for i in base_items), default=0)

                t_data = []
                for r in b['rows']:
                    for item in r['items']:
                        t_data.append({**item, '위치': '바닥', '이동': f"{b['id']}번", '⚠️': ''})
                for s in b.get('stacked_items', []):
                    # 바닥 최솟값보다 L 또는 W가 크면 하중 위험
                    danger = '🚨' if (s['L'] > base_min_L or s['W'] > base_min_W) else ''
                    t_data.append({**s, '위치': '단적', '이동': f"{b['id']}번", '⚠️': danger})

                df_edit = pd.DataFrame(t_data)[['⚠️', '위치', 'PKG NO', 'ITEM', 'L', 'W', 'H', 'WEIGHT', '이동']]
                edited_df = st.data_editor(df_edit, hide_index=True, use_container_width=True, key=f"ed_{b['id']}",
                                        column_config={"이동": st.column_config.SelectboxColumn("🚚 이동", options=target_options)},
                                        disabled=['⚠️', '위치', 'PKG NO', 'ITEM', 'L', 'W', 'H', 'WEIGHT'])
                
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
                        st.session_state.bins = apply_labels(sorted(list(new_bins_dict.values()), key=lambda x: x['id']), max_20_len, max_20_wt, max_dry_h, max_hc_h, max_fr20_len, max_fr20_wt, max_fr40_len, max_fr40_wt)
                        st.rerun()

                with st.expander("👁️ 적재 단면도 및 제원 확인", expanded=False):
                    # c_label 기반 정확한 제원 결정
                    _lbl = b['c_label']
                    if   '20ft Dry'       in _lbl: cur_max_l, cur_max_w, cur_max_h = max_20_len,   max_20_wt,   max_dry_h
                    elif '40ft HC'        in _lbl: cur_max_l, cur_max_w, cur_max_h = max_40_len,   max_40_wt,   max_hc_h
                    elif '40ft Dry'       in _lbl: cur_max_l, cur_max_w, cur_max_h = max_40_len,   max_40_wt,   max_dry_h
                    elif '20ft Flat Rack' in _lbl: cur_max_l, cur_max_w, cur_max_h = max_fr20_len, max_fr20_wt, max_fr_h
                    elif '40ft Flat Rack' in _lbl: cur_max_l, cur_max_w, cur_max_h = max_fr40_len, max_fr40_wt, max_fr_h
                    else:                          cur_max_l, cur_max_w, cur_max_h = max_40_len,   max_40_wt,   max_dry_h
                    used_width = max([r['used_W'] for r in b['rows']] + [0])
                    # 다단 높이 계산: 동일 치수 기준 최대 적재 단수 × 높이
                    if b.get('stacked_items'):
                        from collections import Counter
                        base_cnt = Counter((i['L'],i['W'],i['H']) for r in b['rows'] for i in r['items'])
                        stk_cnt  = Counter((s['L'],s['W'],s['H']) for s in b['stacked_items'])
                        max_stk_h = max(
                            (h * (1 + stk_cnt.get((l,w,h),0) // max(1, base_cnt.get((l,w,h),1))))
                            for (l,w,h) in base_cnt if (l,w,h) in stk_cnt
                        ) if stk_cnt else b['max_H']
                        used_height = max(b['max_H'], max_stk_h)
                    else:
                        used_height = b['max_H']
                    
                    c1, c2, c3, c4 = st.columns(4)
                    c1.markdown(f"**📏 길이:** {b['used_L']:,}/{cur_max_l:,}mm"); c1.progress(min(1.0, b['used_L']/cur_max_l))
                    c2.markdown(f"**⚖️ 중량:** {b['total_W']:,}/{cur_max_w:,}kg"); c2.progress(min(1.0, b['total_W']/cur_max_w))
                    
                    if used_width > 2350: c3.markdown(f"**↔️ 폭:** {used_width:,}/2340mm <span style='color:{ALERT_COLOR};font-weight:bold;'>[OW +{used_width-2350:,}]</span>", unsafe_allow_html=True)
                    else: c3.markdown(f"**↔️ 폭:** {used_width:,}/2340mm"); c3.progress(min(1.0, used_width/2340))
                    
                    if used_height > cur_max_h: c4.markdown(f"**↕️ 높이:** {used_height:,}/{cur_max_h:,}mm <span style='color:{ALERT_COLOR};font-weight:bold;'>[OH +{used_height-cur_max_h:,}]</span>", unsafe_allow_html=True)
                    else: c4.markdown(f"**↕️ 높이:** {used_height:,}/{cur_max_h:,}mm"); c4.progress(min(1.0, used_height/cur_max_h))
                    
                    # 범례 표시
                    st.markdown(
                        f"<span style='background:{ALERT_COLOR};padding:2px 8px;border-radius:4px;color:white;font-size:12px;'>■ FR (OW/OH)</span>"
                        f"&nbsp;&nbsp;"
                        f"<span style='background:#e67e22;padding:2px 8px;border-radius:4px;color:white;font-size:12px;'>■ HC (H>{max_dry_h}mm)</span>"
                        f"&nbsp;&nbsp;"
                        f"<span style='background:{ACCENT_COLOR};padding:2px 8px;border-radius:4px;color:white;font-size:12px;'>■ DRY (H≤{max_dry_h}mm)</span>",
                        unsafe_allow_html=True
                    )
                    from collections import Counter
                    stk_cnt  = Counter((s['L'],s['W'],s['H']) for s in b.get('stacked_items',[]))
                    base_cnt = Counter((i['L'],i['W'],i['H']) for r in b['rows'] for i in r['items'])

                    # ── 평면도 (L × W, 위에서 본 뷰) ──────────────────────
                    fig = go.Figure()
                    fig.add_shape(type="rect", x0=b['used_L'], y0=0, x1=cur_max_l, y1=2340, fillcolor="#e1e4e8", opacity=0.4, line_width=0)
                    fig.add_shape(type="rect", x0=0, y0=0, x1=cur_max_l, y1=2340, line=dict(color=MAIN_COLOR, width=2))
                    cx = 0
                    for r in b['rows']:
                        cy = (2340 - r['used_W']) / 2
                        for item in r['items']:
                            dim = (item['L'], item['W'], item['H'])
                            layers = 1 + math.ceil(stk_cnt.get(dim, 0) / max(1, base_cnt.get(dim, 1)))
                            if item['W'] > 2340 or item['H'] > max_hc_h: item_color = ALERT_COLOR
                            elif item['H'] > max_dry_h: item_color = "#e67e22"
                            else: item_color = ACCENT_COLOR
                            # 다단 화물: 테두리 강조
                            border = dict(color="#FFD700", width=3) if layers > 1 else dict(color="white", width=1)
                            fig.add_shape(type="rect", x0=cx, y0=cy, x1=cx+item['L'], y1=cy+item['W'], fillcolor=item_color, opacity=0.85, line=border)
                            # 다단 화물: 바닥 PKG NO + 단수 표시 (stacked는 테이블에서 확인)
                            if layers > 1:
                                label = f"<b>×{layers}단</b><br>{item['PKG NO']}<br><sub>H{item['H']}</sub>"
                            else:
                                label = f"{item['PKG NO']}<br><sub>H{item['H']}</sub>"
                            fig.add_annotation(x=cx+item['L']/2, y=cy+item['W']/2,
                                text=label, showarrow=False, font=dict(color="white", size=9))
                            cy += item['W']
                        cx += r['max_L']
                    fig.add_shape(type="line", x0=b['used_L'], y0=-200, x1=b['used_L'], y1=2800, line=dict(color=ALERT_COLOR, width=2, dash="dash"))
                    if b['used_L'] > 100: fig.add_annotation(x=b['used_L']/2, y=2650, text=f"적재: {b['used_L']:,}mm", showarrow=False, font=dict(color=MAIN_COLOR, size=13, weight="bold"))
                    if cur_max_l - b['used_L'] > 100: fig.add_annotation(x=b['used_L'] + (cur_max_l - b['used_L'])/2, y=2650, text=f"잔여: {cur_max_l - b['used_L']:,}mm", showarrow=False, font=dict(color=ALERT_COLOR, size=13, weight="bold"))
                    fig.update_layout(xaxis=dict(visible=False, range=[-200, max_40_len+400]), yaxis=dict(visible=False, range=[-300, 3100]), height=260, margin=dict(l=10, r=10, t=10, b=10), paper_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig, use_container_width=True, key=f"plot_{b['id']}")

                st.markdown('</div>', unsafe_allow_html=True)

            # --- 최종 결과 다운로드: 원본 엑셀 양식 유지 ---
            st.markdown("---")

            # ── mapping 생성: BOX 화물은 컨테이너별 요약, 일반은 1:1
            from collections import defaultdict, Counter as _Counter
            row_container_cnt = defaultdict(_Counter)  # row_idx → {c_label: count}
            row_detail_list   = defaultdict(list)       # row_idx → [(pkg_no, c_label)]
            for bx in bins:
                for item in ([i for r in bx['rows'] for i in r['items']] + bx.get('stacked_items', [])):
                    row_container_cnt[item['row_idx']][bx['c_label']] += 1
                    row_detail_list[item['row_idx']].append((item['PKG NO'], bx['c_label'], item['L'], item['W'], item['H'], item['WEIGHT']))

            mapping = {}
            for row_idx, counter in row_container_cnt.items():
                total = sum(counter.values())
                if total > 1:  # BOX 모드 → 컨테이너별 요약
                    mapping[row_idx] = " / ".join(
                        f"{lbl} ({cnt}개)" for lbl, cnt in counter.items()
                    )
                else:
                    mapping[row_idx] = list(counter.keys())[0]

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

                # 배정 결과 입력 (row_idx는 pandas 기준 → 엑셀 행 번호 +1)
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

                ws.column_dimensions[target_letter].width = 30

                # ── BOX 화물이 있으면 "배정상세" 시트 추가
                box_details = [(pkg, c_label, item_l, item_w, item_h, item_wt)
                               for row_idx, details in row_detail_list.items()
                               if len(details) > 1
                               for pkg, c_label, item_l, item_w, item_h, item_wt in details]
                if box_details:
                    if "배정상세" in wb.sheetnames:
                        del wb["배정상세"]
                    ws_d = wb.create_sheet("배정상세")
                    ws_d.append(["PKG NO", "L (mm)", "W (mm)", "H (mm)", "WEIGHT (kg)", "배정 컨테이너"])
                    ws_d.column_dimensions["A"].width = 20
                    ws_d.column_dimensions["B"].width = 10
                    ws_d.column_dimensions["C"].width = 10
                    ws_d.column_dimensions["D"].width = 10
                    ws_d.column_dimensions["E"].width = 14
                    ws_d.column_dimensions["F"].width = 30
                    for pkg, c_label, l, w, h, wt in box_details:
                        ws_d.append([pkg, l, w, h, wt, c_label])

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
