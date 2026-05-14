import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import math
import os
import io
import re
from openpyxl import load_workbook
from copy import copy

# --- 3D Bin Packing 알고리즘 (py3dbp 경량화 자체 구현) ---
# (외부 라이브러리 설치 부담을 줄이기 위해 핵심 로직을 내장합니다)
from decimal import Decimal

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
        # 회전 상태 (0: 원래, 1: W와 L 회전) - 높이는 회전하지 않음(파손 방지)
        self.rotation_type = 0 
        self.position = [0, 0, 0] # x, y, z 좌표

    def get_dimension(self):
        if self.rotation_type == 0: return [self.length, self.width, self.height]
        elif self.rotation_type == 1: return [self.width, self.length, self.height]
        return [self.length, self.width, self.height]

class Bin:
    def __init__(self, name, length, width, height, max_weight):
        self.name = name
        self.length = length
        self.width = width
        self.height = height
        self.max_weight = max_weight
        self.items = []
        self.unfitted_items = []

    def get_total_weight(self):
        return sum([item.weight for item in self.items])

    def get_total_volume(self):
        return sum([item.length * item.width * item.height for item in self.items])

    def put_item(self, item, pivot):
        fit = False
        valid_item_position = item.position
        item.position = pivot

        for i in range(0, 2): # 높이 회전 금지 (0: 기본, 1: L-W 스왑)
            item.rotation_type = i
            dimension = item.get_dimension()
            
            # 컨테이너 밖으로 나가는지 체크
            if (
                item.position[0] + dimension[0] <= self.length and
                item.position[1] + dimension[1] <= self.width and
                item.position[2] + dimension[2] <= self.height
            ):
                # 중량 초과 체크
                if self.get_total_weight() + item.weight <= self.max_weight:
                    # 겹침 체크
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
        return fit

    def intersect(self, item1, item2):
        rect1 = item1.position + item1.get_dimension()
        rect2 = item2.position + item2.get_dimension()

        # x축 겹침
        cx = rect1[0] < rect2[0] + rect2[3] and rect1[0] + rect1[3] > rect2[0]
        # y축 겹침
        cy = rect1[1] < rect2[1] + rect2[4] and rect1[1] + rect1[4] > rect2[1]
        # z축 겹침
        cz = rect1[2] < rect2[2] + rect2[5] and rect1[2] + rect1[5] > rect2[2]

        return cx and cy and cz

class Packer:
    def __init__(self):
        self.bins = []
        self.items = []
        self.unfit_items = []
        self.total_items = 0

    def add_bin(self, bin): return self.bins.append(bin)
    def add_item(self, item): 
        self.total_items += 1
        return self.items.append(item)

    def pack_to_bin(self, bin, item):
        fitted = False
        if not bin.items:
            response = bin.put_item(item, [0, 0, 0])
            if response:
                bin.items.append(item)
                fitted = True
            return fitted

        for axis in range(0, 3):
            for ib in bin.items:
                pivot = [0, 0, 0]
                w, h, d = ib.get_dimension()
                if axis == 0: pivot = [ib.position[0] + w, ib.position[1], ib.position[2]]
                elif axis == 1: pivot = [ib.position[0], ib.position[1] + h, ib.position[2]]
                elif axis == 2: pivot = [ib.position[0], ib.position[1], ib.position[2] + d]

                if bin.put_item(item, pivot):
                    bin.items.append(item)
                    fitted = True
                    break
            if fitted: break
        return fitted

    def pack(self):
        # 크기가 큰 순서, 그 다음 번호 순서로 정렬하여 테트리스 시작
        self.items.sort(key=lambda item: (item.length * item.width * item.height, -item.p_seq), reverse=True)
        
        for bin in self.bins:
            unpacked = []
            for item in self.items:
                if not self.pack_to_bin(bin, item):
                    unpacked.append(item)
            self.items = unpacked

# --- UI 및 기존 시스템 로직 ---

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
.kpi-card {{ background-color:white;padding:15px;border-radius:10px;border-left:5px solid {MAIN_COLOR};box-shadow:0 2px 4px rgba(0,0,0,0.05); height:100%; }}
.kpi-title {{ font-size:14px; color:#555; margin-bottom:5px; font-weight:600; }}
.kpi-value {{ font-size:28px; font-weight:800; color:#111; }}
</style>""", unsafe_allow_html=True)

st.markdown(f'<div class="header-container"><div style="font-size:30px;font-weight:800;">CALT-LOGIS CLP SYSTEM</div><div style="font-size:15px;opacity:0.8;margin-top:5px;">Busan New Port Center | 3D Algorithm Engine</div></div>', unsafe_allow_html=True)

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
        st.markdown("""<table class="guide-table">
<tr><th>항목</th><th>엑셀 열</th></tr>
<tr class="essential"><td>PKG NO</td><td>B</td></tr>
<tr class="essential"><td>L / W / H</td><td>J / L / N</td></tr>
<tr><td>WEIGHT</td><td>I</td></tr>
</table>""", unsafe_allow_html=True)
        st.download_button("📥 양식 다운로드", io.BytesIO(), "CALT_CLP_TEMPLATE.xlsx")

    st.header("⚙️ 배정 옵션 설정")
    with st.expander("⚖️ 컨테이너 제원", expanded=False):
        max_20_wt  = st.number_input("최대 중량 (20ft)", 15000,40000,28250)
        max_20_len = st.number_input("최대 길이 (20ft)", 5500,6500,5900)
        max_40_wt  = st.number_input("최대 중량 (40ft)", 20000,40000,29500)
        max_40_len = st.number_input("최대 길이 (40ft)", 11000,13000,11900)
        max_dry_h  = st.number_input("DRY 내부 높이", 2000,3000,2370)
        max_hc_h   = st.number_input("HC 내부 높이",  2000,3500,2670)

    if st.button("🔄 AI 3D 재계산 실행"): reset_data()

# --- 3D Packing Bridge (DataFrame -> 3D Engine -> GUI Data) ---
def apply_labels_3d(bins_list, max_dry_h, max_hc_h):
    formatted_bins = []
    for idx, b in enumerate(bins_list):
        used_L = max([i.position[0] + i.get_dimension()[0] for i in b.items]) if b.items else 0
        max_H = max([i.position[2] + i.get_dimension()[2] for i in b.items]) if b.items else 0
        total_W = sum([i.weight for i in b.items])
        
        # 라벨링 판별
        lbl = "40ft HC" if max_H > max_dry_h else "40ft Dry"
        
        gui_bin = {
            'id': idx + 1,
            'c_label': f"{lbl} #{idx+1}",
            'used_L': used_L,
            'total_W': total_W,
            'max_H': max_H,
            'items': b.items # 3D 계산된 아이템 객체들
        }
        formatted_bins.append(gui_bin)
    return formatted_bins

def run_3d_packing(df, max_40_len, max_40_wt, max_hc_h):
    packer = Packer()
    
    # 1. 엑셀 데이터를 3D Item 객체로 변환
    for _, row in df.iterrows():
        l, w, h, wt = int(row['L']+.5), int(row['W']+.5), int(row['H']+.5), int(row['WEIGHT']+.5)
        nums = re.findall(r'\d+', str(row['PKG NO']))
        p_seq = int("".join(nums)) if nums else 999999
        
        # 3D 회전을 위해 FORK_L은 (W, L, H)로 초기 세팅
        if row.get('LOAD_KW') == 'FORK_L':
            item = Item(row['PKG NO'], w, l, h, wt, p_seq=p_seq, orig_idx=row.get('row_idx', 0))
        else:
            item = Item(row['PKG NO'], l, w, h, wt, p_seq=p_seq, orig_idx=row.get('row_idx', 0))
        
        # 메타데이터 주입
        item.item_desc = row.get('ITEM', '-')
        item.group_desc = row.get('GROUP', '-')
        packer.add_item(item)

    # 2. 충분한 컨테이너(Bin) 준비 (넉넉하게 생성해두고 안 쓰면 삭제)
    for i in range(10):
        packer.add_bin(Bin(f"40HC-{i}", max_40_len, 2340, max_hc_h, max_40_wt))

    # 3. 3D 패킹 실행
    packer.pack()

    # 4. 쓰인 컨테이너만 필터링하여 GUI 포맷으로 변환
    used_bins = [b for b in packer.bins if len(b.items) > 0]
    
    return apply_labels_3d(used_bins, max_dry_h, max_hc_h)

# --- 메인 실행 ---
st.markdown("### 📤 패킹리스트 업로드")
load_mode_ui = st.radio("👉 **지게차(포크) 기본 진입 방향 설정**", ["FORK_L", "FORK_W", "4WAY"], horizontal=True, on_change=reset_data)
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
                cnt=sum(1 for _,r in sd.iterrows() if str(r.iloc[COL_PKG]).strip().lower() not in ['','nan'] and clean_num(r.iloc[COL_L])>0)
                if cnt>best: best=cnt; raw_full=sd; selected_sheet=sn
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
                if pkg.lower() in ['','nan'] or lv==0: continue
                wtv=clean_num(row.iloc[COL_WEIGHT])
                qty=int(clean_num(row.iloc[COL_QTY])) if clean_num(row.iloc[COL_QTY])>0 else 1
                rem = str(row.iloc[COL_REMARK]).upper()
                load_val = str(row.iloc[COL_LOAD]).upper()
                
                load_kw = "FORK_W" if "FORK_W" in load_val else ("FORK_L" if "FORK_L" in load_val else load_mode_ui)
                
                repeat = qty if 'BOX' in rem else 1
                for seq in range(repeat):
                    sfx=f"-{seq+1:03d}" if repeat>1 else ""
                    p_data.append({'PKG NO':f"{pkg}{sfx}", 'L':lv, 'W':wv, 'H':hv, 'WEIGHT':wtv/repeat, 'LOAD_KW':load_kw, 'row_idx':int(oi)})
            except: continue

        df=pd.DataFrame(p_data)
        if df.empty: st.warning("필수 데이터를 찾을 수 없습니다.")
        else:
            if 'bins' not in st.session_state or not st.session_state.get('manual_mode',False):
                # ★★★ 3D 엔진으로 교체 호출 ★★★
                st.session_state.bins = run_3d_packing(df, max_40_len, max_40_wt, max_hc_h)
                st.session_state.manual_mode=True

            bins=st.session_state.bins

            # KPI 섹션
            st.subheader("📊 3D 테트리스 실시간 적재 요약")
            c1, c2, c3, c4 = st.columns(4)
            
            packed = sum(len(b['items']) for b in bins)
            total_w = sum(b['total_W'] for b in bins)
            
            c1.markdown(f'<div class="kpi-card"><div class="kpi-title">배정 화물 / 전체 화물</div><div class="kpi-value"><span style="color:{ACCENT_COLOR};">{packed}</span> / {len(df)}</div></div>', unsafe_allow_html=True)
            c2.markdown(f'<div class="kpi-card"><div class="kpi-title">컨테이너 수</div><div class="kpi-value">{len(bins)} UNIT</div></div>', unsafe_allow_html=True)
            c4.markdown(f'<div class="kpi-card"><div class="kpi-title">배정 중량</div><div class="kpi-value">{total_w:,.0f} kg</div></div>', unsafe_allow_html=True)

            for b in bins:
                st.markdown('<div class="container-box">', unsafe_allow_html=True)
                st.markdown(f"### 📦 {b['c_label']} (사용 길이: {b['used_L']:,}mm)")
                
                # 3D 결과 테이블 출력 (Z축 추가)
                t_data = []
                for item in b['items']:
                    dim = item.get_dimension()
                    t_data.append({
                        'PKG NO': item.name,
                        'X (길이축)': f"{item.position[0]:,} ~ {item.position[0] + dim[0]:,}",
                        'Y (폭축)': f"{item.position[1]:,} ~ {item.position[1] + dim[1]:,}",
                        'Z (높이축)': f"{item.position[2]:,} ~ {item.position[2] + dim[2]:,} (단적 여부 확인)",
                        'L': dim[0], 'W': dim[1], 'H': dim[2], 'WEIGHT': item.weight
                    })
                st.dataframe(pd.DataFrame(t_data), use_container_width=True, hide_index=True)
                st.markdown('</div>',unsafe_allow_html=True)

    except Exception as e: st.error(f"데이터 처리 중 오류 발생: {e}")
