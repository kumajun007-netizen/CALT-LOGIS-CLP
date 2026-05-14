import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import math
import os
import io
import re
from openpyxl import load_workbook
from copy import copy

# --- 3D Bin Packing 알고리즘 (캐시 충돌 방지형) ---

class Item:
    def __init__(self, name, length, width, height, weight, max_stk=1, group='', p_seq=999999, orig_idx=0):
        self.name = name
        self.length = length
        self.width = width
        self.height = height
        self.weight = weight
        self.max_stk = max_stk
        self.group = group
        self.p_seq = p_seq
        self.orig_idx = orig_idx
        self.rotation_type = 0 # 0: 원본, 1: 90도 회전 (L, W 스왑)
        self.position = [0, 0, 0] # x, y, z

# 스트림릿 캐시 충돌을 막기 위한 독립 헬퍼 함수
def get_item_dim(item):
    rot = getattr(item, 'rotation_type', 0)
    l = getattr(item, 'length', 0)
    w = getattr(item, 'width', 0)
    h = getattr(item, 'height', 0)
    return [l, w, h] if rot == 0 else [w, l, h]

class Bin:
    def __init__(self, name, length, width, height, max_weight):
        self.name = name
        self.length = length
        self.width = width
        self.height = height
        self.max_weight = max_weight
        self.items = []

    def get_total_weight(self):
        return sum([getattr(item, 'weight', 0) for item in self.items])

    def intersect(self, item1, item2):
        dim1 = get_item_dim(item1)
        dim2 = get_item_dim(item2)
        cx = item1.position[0] < item2.position[0] + dim2[0] and item1.position[0] + dim1[0] > item2.position[0]
        cy = item1.position[1] < item2.position[1] + dim2[1] and item1.position[1] + dim1[1] > item2.position[1]
        cz = item1.position[2] < item2.position[2] + dim2[2] and item1.position[2] + dim1[2] > item2.position[2]
        return cx and cy and cz

    def put_item(self, item, pivot, support_item=None):
        valid_item_position = item.position
        item.position = pivot
        fit = False

        for i in range(2): # 높이(H)는 유지한 채 바닥면 90도 회전 시도
            item.rotation_type = i
            dim = get_item_dim(item)
            
            # 1. 컨테이너 내부 이탈 방지
            if (item.position[0] + dim[0] <= self.length and
                item.position[1] + dim[1] <= self.width and
                item.position[2] + dim[2] <= self.height):
                
                # 2. 총 중량 한도 체크
                if self.get_total_weight() + item.weight <= self.max_weight:
                    
                    # 3. ★ 실무형 혼합 적재(Mixed Stacking) 룰 검증 ★
                    # 위에 쌓는 경우, 아래 화물의 면적이 내 면적을 완전히 받쳐줄 수 있는지 검사
                    if support_item is not None:
                        sup_dim = get_item_dim(support_item)
                        # 아래 화물보다 길이가 길거나 폭이 넓으면 무너질 위험이 있으므로 허용 안함
                        if dim[0] > sup_dim[0] or dim[1] > sup_dim[1]:
                            continue 

                    # 4. 다른 화물과 겹치는지 검사
                    intersect = False
                    for b_item in self.items:
                        if self.intersect(b_item, item):
                            intersect = True
                            break
                    
                    if not intersect:
                        fit = True
                        break
        
        if not fit:
            item.position = valid_item_position
            item.rotation_type = 0
        return fit

class Packer:
    def __init__(self):
        self.bins = []
        self.items = []

    def add_bin(self, bin): self.bins.append(bin)
    def add_item(self, item): self.items.append(item)

    def pack_to_bin(self, bin, item):
        if not bin.items:
            if bin.put_item(item, [0, 0, 0], None):
                bin.items.append(item)
                return True
            return False

        for ib in bin.items:
            ib_dim = get_item_dim(ib)
            
            # X축 (옆에 붙이기)
            if bin.put_item(item, [ib.position[0] + ib_dim[0], ib.position[1], ib.position[2]], None):
                bin.items.append(item); return True
                
            # Y축 (앞에 붙이기)
            if bin.put_item(item, [ib.position[0], ib.position[1] + ib_dim[1], ib.position[2]], None):
                bin.items.append(item); return True
                
            # Z축 (★ 위에 쌓기 - support_item 전달)
            if bin.put_item(item, [ib.position[0], ib.position[1], ib.position[2] + ib_dim[2]], support_item=ib):
                bin.items.append(item); return True

        return False

    def pack(self):
        # 최적화를 위해 부피가 큰 화물부터, 동일 부피면 PKG 번호순으로 먼저 넣습니다.
        self.items.sort(key=lambda item: (get_item_dim(item)[0] * get_item_dim(item)[1] * get_item_dim(item)[2], -item.p_seq), reverse=True)
        
        for item in self.items:
            for bin in self.bins:
                if self.pack_to_bin(bin, item):
                    break

# --- UI 연동 브릿지 ---
def apply_labels_3d(bins_list, max_dry_h, max_hc_h):
    formatted_bins = []
    for idx, b in enumerate(bins_list):
        used_L = max([i.position[0] + get_item_dim(i)[0] for i in b.items]) if b.items else 0
        max_H = max([i.position[2] + get_item_dim(i)[2] for i in b.items]) if b.items else 0
        total_W = sum([getattr(i, 'weight', 0) for i in b.items])
        
        lbl = "40ft HC" if max_H > max_dry_h else "40ft Dry"
        gui_bin = {
            'id': idx + 1,
            'c_label': f"{lbl} #{idx+1}",
            'used_L': used_L,
            'total_W': total_W,
            'max_H': max_H,
            'items': b.items
        }
        formatted_bins.append(gui_bin)
    return formatted_bins

def run_3d_packing(df, max_40_len, max_40_wt, max_hc_h):
    packer = Packer()
    for _, row in df.iterrows():
        l, w, h, wt = int(row['L']+.5), int(row['W']+.5), int(row['H']+.5), int(row['WEIGHT']+.5)
        nums = re.findall(r'\d+', str(row['PKG NO']))
        p_seq = int("".join(nums)) if nums else 999999
        
        # 3D 환경이므로 시작할 때 L과 W를 어떻게 넣어도 내부에서 회전하며 맞춤
        item = Item(row['PKG NO'], l, w, h, wt, p_seq=p_seq)
        packer.add_item(item)

    # 컨테이너 여유있게 생성
    for i in range(15):
        packer.add_bin(Bin(f"40HC-{i}", max_40_len, 2340, max_hc_h, max_40_wt))

    packer.pack()
    used_bins = [b for b in packer.bins if len(b.items) > 0]
    return apply_labels_3d(used_bins, 2370, max_hc_h)

# --- 메인 렌더링 영역 ---
st.set_page_config(page_title="CALT-LOGIS 3D CLP", layout="wide")
MAIN_COLOR="#001f3f"; SUB_COLOR="#f0f2f6"; ACCENT_COLOR="#007bff"; ALERT_COLOR="#e74c3c"

st.markdown(f"""<style>
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
* {{ font-family:'Pretendard',sans-serif; }}
.main {{ background-color:#f8f9fa; }}
.header-container {{ background-color:{MAIN_COLOR};padding:25px;border-radius:10px;color:white;margin-bottom:25px;text-align:center;box-shadow:0 4px 6px rgba(0,0,0,0.1); }}
[data-testid="stFileUploadDropzone"] {{ border:2px dashed {MAIN_COLOR};background-color:#eaf1fb;padding:40px;border-radius:12px; margin-top: 10px; }}
.stButton>button {{ width:100%;font-weight:600;color:white;background-color:{MAIN_COLOR};border-radius:8px;border:none;padding:0.5rem 1rem; }}
.container-box {{ background-color:white;padding:20px;border-radius:10px;border:1px solid #e1e4e8;margin-bottom:20px; }}
.kpi-card {{ background-color:white;padding:15px;border-radius:10px;border-left:5px solid {MAIN_COLOR};box-shadow:0 2px 4px rgba(0,0,0,0.05); height:100%; }}
.kpi-title {{ font-size:14px; color:#555; margin-bottom:5px; font-weight:600; }}
.kpi-value {{ font-size:28px; font-weight:800; color:#111; }}
</style>""", unsafe_allow_html=True)

st.markdown(f'<div class="header-container"><div style="font-size:30px;font-weight:800;">CALT-LOGIS CLP SYSTEM</div><div style="font-size:15px;opacity:0.8;margin-top:5px;">Busan New Port Center | 3D Algorithm Engine (Mixed Stacking)</div></div>', unsafe_allow_html=True)

def clean_num(val):
    try:
        if pd.isna(val) or str(val).strip() in ['','.',  'X','x','NaN','nan']: return 0.0
        return float(str(val).replace(',','').strip())
    except: return 0.0

def reset_data():
    for k in ['bins','manual_mode']:
        if k in st.session_state: del st.session_state[k]

with st.sidebar:
    st.header("⚙️ 배정 옵션 설정")
    with st.expander("⚖️ 40ft 제원", expanded=True):
        max_40_wt  = st.number_input("최대 중량 (kg)", 20000,40000,29500)
        max_40_len = st.number_input("최대 길이 (mm)", 11000,13000,11900)
        max_hc_h   = st.number_input("HC 내부 높이 (mm)",  2000,3500,2670)

    if st.button("🔄 AI 3D 재계산 실행 (캐시 초기화)"):
        reset_data()
        st.success("캐시가 초기화되었습니다. 파일을 다시 드래그 앤 드롭 해주세요.")

st.markdown("### 📤 패킹리스트 업로드")
file = st.file_uploader("파일을 업로드하세요. (3D 테트리스 자동 연산)", type=['csv','xlsx'], on_change=reset_data)

if file is not None:
    try:
        COL_PKG=1;COL_ITEM=3;COL_DESC=4;COL_QTY=5;COL_WEIGHT=8;COL_L=9;COL_W=11;COL_H=13;COL_REMARK=14;COL_LOAD=15
        
        if file.name.lower().endswith('.xlsx'):
            file.seek(0); sheets=pd.read_excel(file,sheet_name=None,header=None)
            best=0
            for sn,sd in sheets.items():
                if sn=='사용설명서': continue
                cnt=sum(1 for _,r in sd.iterrows() if str(r.iloc[COL_PKG]).strip().lower() not in ['','nan'] and clean_num(r.iloc[COL_L])>0)
                if cnt>best: best=cnt; raw_full=sd
        else:
            file.seek(0); raw_full=pd.read_csv(file,header=None)

        p_data=[]
        for oi in range(len(raw_full)):
            row=raw_full.iloc[oi]
            if len(row) <= COL_H: continue 
            try:
                pkg=str(row.iloc[COL_PKG]).replace('00:00:00','').replace('.0','').strip()
                lv, wv, hv = clean_num(row.iloc[COL_L]), clean_num(row.iloc[COL_W]), clean_num(row.iloc[COL_H])
                if pkg.lower() in ['','nan'] or lv==0: continue
                wtv=clean_num(row.iloc[COL_WEIGHT])
                qty=int(clean_num(row.iloc[COL_QTY])) if clean_num(row.iloc[COL_QTY])>0 else 1
                rem = str(row.iloc[COL_REMARK]).upper()
                
                repeat = qty if 'BOX' in rem else 1
                for seq in range(repeat):
                    sfx = f"-{seq+1:03d}" if repeat>1 else ""
                    p_data.append({'PKG NO':f"{pkg}{sfx}", 'L':lv, 'W':wv, 'H':hv, 'WEIGHT':wtv/repeat})
            except: continue

        df=pd.DataFrame(p_data)
        if df.empty: st.warning("데이터를 찾을 수 없습니다.")
        else:
            if 'bins' not in st.session_state:
                st.session_state.bins = run_3d_packing(df, max_40_len, max_40_wt, max_hc_h)

            bins=st.session_state.bins

            # KPI 섹션
            st.subheader("📊 3D 적재 요약 (통합 배정)")
            c1, c2, c3 = st.columns(3)
            
            packed = sum(len(b['items']) for b in bins)
            total_w = sum(b['total_W'] for b in bins)
            
            c1.markdown(f'<div class="kpi-card"><div class="kpi-title">배정 완료 / 전체 화물</div><div class="kpi-value"><span style="color:{ACCENT_COLOR};">{packed}</span> / {len(df)} <span style="font-size:16px;">PKG</span></div></div>', unsafe_allow_html=True)
            c2.markdown(f'<div class="kpi-card"><div class="kpi-title">사용된 컨테이너</div><div class="kpi-value">{len(bins)} <span style="font-size:16px;">UNIT</span></div></div>', unsafe_allow_html=True)
            c3.markdown(f'<div class="kpi-card"><div class="kpi-title">총 배정 중량</div><div class="kpi-value">{total_w:,.0f} <span style="font-size:16px;">kg</span></div></div>', unsafe_allow_html=True)

            for b in bins:
                st.markdown('<div class="container-box">', unsafe_allow_html=True)
                st.markdown(f"### 📦 {b['c_label']} (소요 길이: {b['used_L']:,} / {max_40_len:,} mm)")
                
                # 3D 배치 결과를 확인하기 쉬운 표로 정리
                t_data = []
                for item in b['items']:
                    dim = get_item_dim(item)
                    z_pos = item.position[2]
                    # 높이가 0보다 크면 무조건 무언가(다른 박스)의 위에 쌓여있는 상태(혼적/다단)
                    stack_status = "바닥" if z_pos == 0 else f"혼합/다단 (H:{z_pos:,}부터)" 
                    
                    t_data.append({
                        'PKG NO': item.name,
                        '적재 상태': stack_status,
                        'X (시작점~끝점)': f"{item.position[0]:,} ~ {item.position[0] + dim[0]:,}",
                        'Y (시작점~끝점)': f"{item.position[1]:,} ~ {item.position[1] + dim[1]:,}",
                        '적재된 길이(L)': dim[0], 
                        '적재된 폭(W)': dim[1], 
                        '적재된 높이(H)': dim[2]
                    })
                
                # 표 출력 시 적재된 위치(X, Y, Z)에 따라 정렬하여 직관적으로 보여줌
                df_res = pd.DataFrame(t_data).sort_values(by=['X (시작점~끝점)', 'Y (시작점~끝점)', '적재 상태'])
                st.dataframe(df_res, use_container_width=True, hide_index=True)
                st.markdown('</div>',unsafe_allow_html=True)

    except Exception as e: 
        st.error(f"오류가 발생했습니다. 좌측의 [AI 3D 재계산 실행] 버튼을 눌러주세요. 상세: {e}")
