import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import math
import os
import io

# --- 1. 화면 스타일 ---
st.set_page_config(page_title="CALTALOGIS CLP 시스템", layout="wide")
st.markdown("""
    <style>
    .navy-title { color: #001f3f; font-size: 32px; font-weight: bold; margin-bottom: 10px; }
    div[data-testid="stTable"] { font-size: 11px !important; }
    th { background-color: #f0f2f6 !important; font-size: 11px !important; font-weight: bold; color: #001f3f; }
    td { font-size: 11px !important; }
    .stExpander { border: 1px solid #e6e9ef; border-radius: 5px; margin-bottom: 10px; }
    .stButton>button { width: 100%; font-weight: bold; color: white; background-color: #001f3f; border-radius: 5px; }
    .stButton>button:hover { background-color: #003366; color: white; border-color: #001f3f; }
    </style>
    """, unsafe_allow_html=True)

logo_path = "칼트로지스로고.png"
if os.path.exists(logo_path):
    st.image(logo_path, width=280)

st.markdown('<p class="navy-title">칼트로지스 CLP 시스템</p>', unsafe_allow_html=True)


def clean_num(val):
    try:
        if pd.isna(val) or str(val).strip() in ['', '.', 'X', 'x']:
            return None
        return float(str(val).replace(',', '').strip())
    except:
        return None


# --- 2. 짐 배치 핵심 로직 ---
def pack_items_into_bin(pieces, b, max_40_wt, max_40_len):
    for piece in pieces:
        if piece['STACK_OK'] and piece['WEIGHT'] <= 1000 and b['total_W'] + piece['WEIGHT'] <= max_40_wt:
            if 'stacked_items' not in b:
                b['stacked_items'] = []
            b['stacked_items'].append(piece)
            b['total_W'] += piece['WEIGHT']
            b['groups'].add(piece['GROUP'])
            continue

        row_found = False

        for r in b['rows']:
            temp_max_L = max(r['max_L'], piece['L'])

            if (
                r['used_W'] + piece['W'] <= 2350
                and b['used_L'] + (temp_max_L - r['max_L']) <= max_40_len
                and b['total_W'] + piece['WEIGHT'] <= max_40_wt
            ):
                r['items'].append(piece)
                r['used_W'] += piece['W']
                b['used_L'] += (temp_max_L - r['max_L'])
                r['max_L'] = temp_max_L
                b['total_W'] += piece['WEIGHT']
                b['max_W'] = max(b['max_W'], piece['W'])
                b['max_H'] = max(b['max_H'], piece['H'])
                b['groups'].add(piece['GROUP'])
                row_found = True
                break

        if row_found:
            continue

        if b['used_L'] + piece['L'] <= max_40_len and b['total_W'] + piece['WEIGHT'] <= max_40_wt:
            b['rows'].append({
                'items': [piece],
                'used_W': piece['W'],
                'max_L': piece['L']
            })
            b['used_L'] += piece['L']
            b['total_W'] += piece['WEIGHT']
            b['max_W'] = max(b['max_W'], piece['W'])
            b['max_H'] = max(b['max_H'], piece['H'])
            b['groups'].add(piece['GROUP'])


def apply_labels(bins, max_20_len, max_20_wt, fr_max_len, max_dry_h):
    for b in bins:
        is_20ft_size = b['used_L'] <= max_20_len and b['total_W'] <= max_20_wt
        is_ow = b['max_W'] > 2300
        is_oh = b['max_H'] > 1900
        is_ol = b['used_L'] > fr_max_len
        is_fv = is_ol or (is_ow and is_oh)

        tags = []

        if is_fv:
            tags.append("FV")
        else:
            if is_oh:
                tags.append("OH")
            if is_ow:
                tags.append("OW")
            if is_ol:
                tags.append("OL")

        if b['type'] == "FR":
            base = "20ft Flat Rack" if is_20ft_size else "40ft Flat Rack"
            if is_fv:
                base += " (FV)"
            b['c_label'] = f"{base} [{' + '.join(tags)}] #{b['id']}" if tags else f"{base} #{b['id']}"
        else:
            if b['max_H'] > max_dry_h:
                base = "40ft HC (High Cube)"
            else:
                base = "20ft Dry" if is_20ft_size else "40ft Dry"
            b['c_label'] = f"{base} #{b['id']}"

    return bins


def calculate_expert_packing(
    df,
    max_40_wt,
    max_40_len,
    max_20_wt,
    max_20_len,
    max_dry_h,
    max_hc_h,
    use_balancing
):
    all_pieces = []
    fr_max_len = max_40_len - 430

    for _, row in df.iterrows():
        l = int(row['L'] + 0.5)
        w = int(row['W'] + 0.5)
        h = int(row['H'] + 0.5)
        weight = int(row['WEIGHT'] + 0.5)

        if max(l, w) <= 2350:
            el = min(l, w)
            ew = max(l, w)
        else:
            el = max(l, w)
            ew = min(l, w)

        req = "FR" if ew > 2350 or el > fr_max_len or h > max_hc_h else "DRY"

        all_pieces.append({
            **row,
            'L': el,
            'W': ew,
            'H': h,
            'WEIGHT': weight,
            'REQ_TYPE': req
        })

    all_pieces.sort(key=lambda x: (x['REQ_TYPE'], x['GROUP'], x['PKG NO']))

    bins = []
    c_no = 1
    req_types = sorted(list(set(p['REQ_TYPE'] for p in all_pieces)))

    for r_type in req_types:
        r_pieces = [p for p in all_pieces if p['REQ_TYPE'] == r_type]
        type_bins = []

        if use_balancing:
            total_l = sum(p['L'] for p in r_pieces)
            total_w = sum(p['WEIGHT'] for p in r_pieces)
            est_bins = max(
                1,
                math.ceil(total_l / max_40_len),
                math.ceil(total_w / max_40_wt)
            )

            for _ in range(est_bins):
                type_bins.append({
                    'id': c_no,
                    'type': r_type,
                    'rows': [],
                    'used_L': 0,
                    'total_W': 0,
                    'max_W': 0,
                    'max_H': 0,
                    'stacked_items': [],
                    'groups': set()
                })
                c_no += 1

        for piece in r_pieces:
            placed = False

            if use_balancing:
                type_bins.sort(
                    key=lambda b: (
                        0 if piece['GROUP'] in b['groups'] else 1,
                        b['used_L']
                    )
                )

            for b in (type_bins if use_balancing else bins):
                if not use_balancing and b['type'] != piece['REQ_TYPE']:
                    continue

                pack_items_into_bin([piece], b, max_40_wt, max_40_len)

                if piece in b.get('stacked_items', []) or any(piece in r['items'] for r in b['rows']):
                    placed = True
                    break

            if not placed:
                new_bin = {
                    'id': c_no,
                    'type': piece['REQ_TYPE'],
                    'rows': [],
                    'used_L': 0,
                    'total_W': 0,
                    'max_W': 0,
                    'max_H': 0,
                    'stacked_items': [],
                    'groups': set()
                }

                pack_items_into_bin([piece], new_bin, max_40_wt, max_40_len)

                if use_balancing:
                    type_bins.append(new_bin)
                else:
                    bins.append(new_bin)

                c_no += 1

        if use_balancing:
            bins.extend([b for b in type_bins if b['used_L'] > 0])

    bins.sort(key=lambda x: x['id'])

    return apply_labels(bins, max_20_len, max_20_wt, fr_max_len, max_dry_h)


# --- 3. UI 설정 ---
with st.expander("⚙️ 컨테이너 제원 및 배정 옵션", expanded=False):
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        max_40_wt = st.number_input("⚖️ 40ft 중량 (kg)", 20000, 40000, 29500)
        max_20_wt = st.number_input("⚖️ 20ft 중량 (kg)", 15000, 40000, 28250)

    with c2:
        max_40_len = st.number_input("📏 40ft 길이 (mm)", 11000, 13000, 12034)
        max_20_len = st.number_input("📏 20ft 길이 (mm)", 5500, 6500, 5899)

    with c3:
        max_dry_h = st.number_input("📏 DRY 높이 (mm)", 2000, 3000, 2390)
        max_hc_h = st.number_input("📏 HC 높이 (mm)", 2000, 3500, 2695)

    with c4:
        st.markdown("<br>", unsafe_allow_html=True)
        use_balancing = st.checkbox("⚖️ 균분적재 (해제 시 몰아적재)", value=True)
        allow_stacking = st.checkbox("🏢 다단적재 (단적) 허용", value=True)

    if st.button("🔄 설정 적용 및 재계산"):
        st.session_state['manual_mode'] = False


file = st.file_uploader("패킹리스트 업로드", type=['csv', 'xlsx'])


if file is not None:
    try:
        # --- 4. 파일 읽기 ---
        # 엑셀의 첫 번째 시트가 비어 있는 경우가 있어, 데이터가 있는 첫 번째 시트를 자동 선택합니다.
        if file.name.endswith('.xlsx'):
            excel_sheets = pd.read_excel(file, sheet_name=None, header=None)

            raw_full = None
            selected_sheet = None

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
            raw_full = pd.read_csv(file, header=None)

        # --- 5. 원본 엑셀 컬럼 매핑 ---
        # pandas는 0부터 시작합니다.
        # A열=0, B열=1, C열=2, D열=3 ...
        #
        # 현재 설정:
        # No.of PKG   : B열
        # ITEM         : D열
        # DESCRIPTION  : E열
        # NET WEIGHT   : I열
        # LENGTH       : J열
        # WIDTH        : L열
        # HEIGHT       : N열

        COL_PKG = 1       # B열: No.of PKG
        COL_ITEM = 3      # D열: ITEM
        COL_DESC = 4      # E열: DESCRIPTION
        COL_WEIGHT = 7    # I열: NET WEIGHT 또는 Gross Weight 기준
        COL_L = 9         # J열: LENGTH
        COL_W = 11        # L열: WIDTH
        COL_H = 13        # N열: HEIGHT

        # 업로드 예시 파일 기준:
        # 엑셀 6행이 제목, 7행부터 실제 데이터입니다.
        # pandas 0-base 기준으로 6은 엑셀 7행입니다.
        data_start_row = 6

        raw_process = (
            raw_full
            .iloc[data_start_row:]
            .reset_index(drop=False)
            .rename(columns={'index': 'orig_idx'})
        )

        p_data = []

        for i in range(len(raw_process)):
            row = raw_process.iloc[i]

            s_ok = row.astype(str).str.contains('단적허용').any() if allow_stacking else False

            pkg_v = str(row[COL_PKG]).replace('00:00:00', '').replace('.0', '').strip()

            if pkg_v.lower() in ['nan', 'none', '.', '']:
                continue

            p_data.append({
                'PKG NO': pkg_v,
                'GROUP': str(row[COL_DESC]),
                'ITEM': str(row[COL_ITEM]),
                'L': clean_num(row[COL_L]),
                'W': clean_num(row[COL_W]),
                'H': clean_num(row[COL_H]),
                'WEIGHT': clean_num(row[COL_WEIGHT]) or 0,
                'STACK_OK': s_ok,
                'row_idx': int(row['orig_idx'])
            })

        df = pd.DataFrame(p_data)

        if df.empty:
            st.error("불러온 데이터가 없습니다. 시트 또는 시작행을 확인해 주세요.")
            st.stop()

        # 필수값: PKG, L, W, H
        # WEIGHT는 없으면 0으로 처리합니다.
        df = df.dropna(subset=['L', 'W', 'H'])
        df = df[~df['ITEM'].str.upper().str.contains('TOTAL', na=False)]

        if df.empty:
            st.error("필수값을 읽지 못했습니다. B/J/L/N 열에 값이 있는지 확인해 주세요.")
            st.stop()

        # 확인용 미리보기
        with st.expander("📋 불러온 원본 데이터 확인", expanded=False):
            st.dataframe(
                df[['PKG NO', 'GROUP', 'ITEM', 'L', 'W', 'H', 'WEIGHT']],
                use_container_width=True
            )

        # --- 6. 자동 배정 계산 ---
        if 'bins' not in st.session_state or not st.session_state.get('manual_mode', False):
            st.session_state.bins = calculate_expert_packing(
                df,
                max_40_wt,
                max_40_len,
                max_20_wt,
                max_20_len,
                max_dry_h,
                max_hc_h,
                use_balancing
            )
            st.session_state.manual_mode = True

        bins = st.session_state.bins
        target_options = [f"{b_opt['id']}번" for b_opt in bins] + ["✨ 새 컨테이너"]

        # --- 7. 결과 출력 ---
        st.markdown("---")
        st.subheader("🔍 배정 무결성 검증 (Cross-Check)")

        packed_qty = (
            sum([len(r['items']) for b in bins for r in b['rows']])
            + sum([len(b.get('stacked_items', [])) for b in bins])
        )

        st.metric("총 적재 수량", f"{packed_qty} PKG / {len(df)} PKG")
        st.markdown("---")

        for b in bins:
            if b['used_L'] == 0 and not b.get('stacked_items'):
                continue

            st.markdown(f"#### 📦 {b['c_label']}")

            t_data = []

            for r in b['rows']:
                for item in r['items']:
                    t_data.append({
                        **item,
                        '상태': '바닥',
                        '이동 목적지': f"{b['id']}번"
                    })

            for s in b.get('stacked_items', []):
                t_data.append({
                    **s,
                    '상태': '단적',
                    '이동 목적지': f"{b['id']}번"
                })

            df_edit = pd.DataFrame(t_data)[
                ['상태', 'PKG NO', 'GROUP', 'ITEM', 'L', 'W', 'H', 'WEIGHT', '이동 목적지']
            ]

            edited_df = st.data_editor(
                df_edit,
                hide_index=True,
                use_container_width=True,
                key=f"ed_{b['id']}",
                column_config={
                    "이동 목적지": st.column_config.SelectboxColumn(
                        "🚚 이동 (클릭)",
                        options=target_options
                    )
                },
                disabled=['상태', 'PKG NO', 'GROUP', 'ITEM', 'L', 'W', 'H', 'WEIGHT']
            )

            if st.button(f"🚀 {b['id']}번 컨테이너 변경 적용", key=f"apply_{b['id']}"):
                moves = [
                    (r['PKG NO'], r['이동 목적지'])
                    for _, r in edited_df.iterrows()
                    if r['이동 목적지'] != f"{b['id']}번"
                ]

                if moves:
                    new_alloc = []
                    max_id = max([bx['id'] for bx in st.session_state.bins])

                    for bx in st.session_state.bins:
                        for item in (
                            [i for r in bx['rows'] for i in r['items']]
                            + bx.get('stacked_items', [])
                        ):
                            target = bx['id']

                            for m_pkg, m_tgt in moves:
                                if str(item['PKG NO']) == str(m_pkg):
                                    target = max_id + 1 if "새" in m_tgt else int(m_tgt.replace("번", ""))

                            new_alloc.append((item, target))

                    new_bins_dict = {}

                    for item, t_id in new_alloc:
                        if t_id not in new_bins_dict:
                            new_bins_dict[t_id] = {
                                'id': t_id,
                                'type': item['REQ_TYPE'],
                                'rows': [],
                                'used_L': 0,
                                'total_W': 0,
                                'max_W': 0,
                                'max_H': 0,
                                'stacked_items': [],
                                'groups': set()
                            }

                        pack_items_into_bin(
                            [item],
                            new_bins_dict[t_id],
                            max_40_wt,
                            max_40_len
                        )

                    st.session_state.bins = apply_labels(
                        sorted(list(new_bins_dict.values()), key=lambda x: x['id']),
                        max_20_len,
                        max_20_wt,
                        max_40_len - 430,
                        max_dry_h
                    )

                    st.rerun()

            with st.expander(f"📊 요약: {b['used_L']:,}mm / {b['total_W']:,}kg (도면)"):
                d_len = max_20_len if "20ft" in b['c_label'] else max_40_len

                fig = go.Figure()

                fig.add_shape(
                    type="rect",
                    x0=0,
                    y0=0,
                    x1=d_len,
                    y1=2350,
                    line=dict(color="black", width=2)
                )

                cx = 0

                for r in b['rows']:
                    cy = (2350 - r['used_W']) / 2

                    for item in r['items']:
                        fig.add_shape(
                            type="rect",
                            x0=cx,
                            y0=cy,
                            x1=cx + item['L'],
                            y1=cy + item['W'],
                            fillcolor="royalblue",
                            opacity=0.7
                        )

                        fig.add_annotation(
                            x=cx + item['L'] / 2,
                            y=cy + item['W'] / 2,
                            text=item['PKG NO'],
                            showarrow=False,
                            font=dict(color="white")
                        )

                        cy += item['W']

                    cx += r['max_L'] + 50

                fig.update_layout(
                    xaxis=dict(range=[-200, d_len + 500]),
                    yaxis=dict(range=[-500, 3000]),
                    height=300,
                    margin=dict(l=0, r=0, t=0, b=0)
                )

                st.plotly_chart(fig, use_container_width=True)

        # --- 8. 엑셀 다운로드: 원본 양식 유지 ---
        mapping = {}

        for bx in bins:
            for item in (
                [i for r in bx['rows'] for i in r['items']]
                + bx.get('stacked_items', [])
            ):
                mapping[item['row_idx']] = bx['c_label']

        export_df = raw_full.copy()

        target_col = export_df.shape[1]
        export_df[target_col] = ""

        # 기존 양식 상단 쪽에 제목 입력
        export_df.iloc[3, target_col] = "배정 컨테이너"

        for r_idx, label in mapping.items():
            export_df.iloc[r_idx, target_col] = label

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            export_df.to_excel(
                writer,
                index=False,
                header=False,
                sheet_name='CLP_결과'
            )

        st.markdown("---")
        st.subheader("💾 최종 배정 결과 다운로드")

        st.download_button(
            label="📥 원본 양식에 컨테이너 번호 달아서 저장하기",
            data=output.getvalue(),
            file_name="CLP_배정결과_원본양식.xlsx",
            use_container_width=True
        )

    except Exception as e:
        st.error(f"오류 발생: {e}")
