import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os

# 페이지 설정
st.set_page_config(page_title="호흡기내과 임상연구 배정", layout="wide", page_icon="🏥")

# 엑셀 파일 이름
STATUS_EXCEL = "status.xlsx"
CRITERIA_FILE = "criteria.xlsx"

# -----------------------------------------------------------------------------
# 1. 엑셀 데이터 로드 함수 (줄바꿈 문제 완벽 해결 버전)
# -----------------------------------------------------------------------------
@st.cache_data(ttl=600)
def load_status_from_excel():
    data = {}
    # 기본 메시지 정의
    default_msg = {
        "copd_sit_severe": "데이터 없음 (엑셀 확인 필요)",
        "copd_sit_maint": "데이터 없음 (엑셀 확인 필요)",
        "copd_sit_be": "데이터 없음 (엑셀 확인 필요)",
        "asthma_eos": "Areteia 등 (데이터 없음)",
        "asthma_rhinitis": "대원제약 등 (데이터 없음)",
        "asthma_bio": "Sanofi 등 (데이터 없음)",
        "etc_be": "데이터 없음",
        "etc_cough": "데이터 없음",
        "etc_acute": "데이터 없음",
        "etc_ipf": "데이터 없음"
    }
    
    if not os.path.exists(STATUS_EXCEL):
        return default_msg

    try:
        wb = load_workbook(STATUS_EXCEL, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=1, values_only=True):
            if row[0] and len(row) > 1:
                key = str(row[0]).strip()
                val = str(row[1]) if row[1] else ""
                
                # [핵심 수정] 엑셀의 줄바꿈을 화면에 강제로 표시하기 위한 처리
                # 1. 윈도우식 줄바꿈(\r\n)을 일반 줄바꿈(\n)으로 통일
                val = val.replace('\r\n', '\n')
                # 2. 줄바꿈(\n)을 "공백 2칸 + 줄바꿈"으로 변경해야 Streamlit이 인식함
                val = val.replace('\n', '  \n') 
                
                data[key] = val
        wb.close()
    except Exception as e:
        st.error(f"엑셀 읽기 오류: {e}")
        return default_msg
        
    # 데이터 병합
    for k, v in default_msg.items():
        if k not in data:
            data[k] = v
    return data

# 데이터 로드
status_data = load_status_from_excel()

# -----------------------------------------------------------------------------
# 2. 웹 화면 구성 (Streamlit)
# -----------------------------------------------------------------------------

st.title("🏥 건국대병원 호흡기내과 임상연구 배정 도우미")
st.markdown(f"Status Data: `{STATUS_EXCEL}` (2025.12 Ver)")
st.divider()

# 탭 생성
tab1, tab2, tab3 = st.tabs(["🫁 COPD", "🌿 천식 (Asthma)", "🦠 기타 (BE/기침/감기)"])

# [탭 1] COPD
with tab1:
    st.header("COPD 환자 배정")
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("1단계: 레지스트리")
        is_new_copd = st.checkbox("기관지확장제 반응 검사 후 FEV1/FVC < 0.7 (신규 진단)")
        
        if is_new_copd:
            st.success("✅ **[필수] KOCOSS 레지스트리 등록 (담당: 함경은)**\n\n* 신규 환자 필수 등록\n* 대상자 중 '노쇠/근감소증 연구' 동시 등록 가능")
            st.info("👉 유형 분류: TB / BE / Asthma / PRISM / Smoker 중 선택")
        else:
            st.write("기존 등록 환자입니다.")

    with col2:
        st.subheader("2단계: 특수 조건 (박초아 담당)")
        home_o2 = st.checkbox("가정 산소 요법 사용 중")
        cough_copd = st.checkbox("만성 기침 (8주 이상, 원인미상)")
        vaccine = st.checkbox("RSV 백신 접종 고려 (50세 이상)")
        
        if home_o2: st.warning("👉 [가정산소] IIT. 마이숨 (MyBreath)")
        if cough_copd: st.warning("👉 [만성기침] IIT. 만성기침 레지스트리")
        if vaccine: st.warning("👉 [백신] GSK. Arexvy PMS")
    
    st.divider()

    st.subheader("3단계: 임상시험(SIT) 추가 배정")
    copd_sit = st.radio("환자의 임상 상태를 선택하세요", 
                        ["선택 안함", "빈번한 급성 악화 (중증/생물학적제제)", "안정적 유지 치료 필요", "기관지확장증 주증상"])
    
    if copd_sit == "빈번한 급성 악화 (중증/생물학적제제)":
        st.error(status_data["copd_sit_severe"])
    elif copd_sit == "안정적 유지 치료 필요":
        st.info(status_data["copd_sit_maint"])
    elif copd_sit == "기관지확장증 주증상":
        st.success(status_data["copd_sit_be"])

# [탭 2] 천식
with tab2:
    st.header("천식 (Asthma) 환자 배정")
    
    st.info("✅ **[기본] TiGER / PRISM / KOSAR (담당: 함경은)**\n\n* 모든 중증/치료불응성 천식 환자 등록")
    
    st.markdown("### 환자 정보 입력")
    col_a, col_b = st.columns([1, 2])
    with col_a:
        eos_input = st.number_input("혈중 호산구(Eosinophil)", min_value=0, step=10)
    with col_b:
        has_rhinitis = st.checkbox("알레르기 비염 동반")
        has_cough_asthma = st.checkbox("만성 기침 (8주 이상) 동반")
        is_uncontrolled = st.checkbox("기존 치료로 조절 안됨 (Uncontrolled)")
        
    st.markdown("### 배정 결과")
    results = []
    
    # 1순위
    if eos_input >= 300:
        st.success(status_data["asthma_eos"])
        results.append(True)
    
    # 2순위
    if has_rhinitis:
        st.warning(status_data["asthma_rhinitis"])
        results.append(True)
    if has_cough_asthma:
        st.warning(status_data["etc_cough"])
        results.append(True)
        
    # 3순위
    if is_uncontrolled:
        st.error(status_data["asthma_bio"])
        results.append(True)
        
    if not results:
        st.info("👉 특별한 SIT 대상이 아닙니다. 1단계 레지스트리 등록을 우선 진행하세요.")

# [탭 3] 기타
with tab3:
    st.header("기타 (BE / 기침 / 급성기관지염 / IPF)")
    
    diagnosis = st.radio("주 진단명을 선택하세요", 
                         ["기관지확장증 (Bronchiectasis)", "만성 기침 (Chronic Cough)", "급성 기관지염 (Acute Bronchitis)", "IPF (특발성 폐섬유증)"])
    
    st.markdown("### 배정 가이드")
    if diagnosis == "기관지확장증 (Bronchiectasis)":
        st.success(status_data["etc_be"])
    elif diagnosis == "만성 기침 (Chronic Cough)":
        st.warning(status_data["etc_cough"])
    elif diagnosis == "급성 기관지염 (Acute Bronchitis)":
        st.info(status_data["etc_acute"])
    elif diagnosis == "IPF (특발성 폐섬유증)":
        st.error(status_data["etc_ipf"])

# ==========================================
# [통합 기능] 하단에 상세 엑셀 파일 표시 (5개 탭 통합 + 검색)
# ==========================================
st.divider()
st.header("📑 연구별 상세 선정/제외 기준 (Reference)")

if os.path.exists(CRITERIA_FILE):
    try:
        # 1. 읽어올 시트 이름 목록 정의
        target_sheets = ["천식", "COPD", "BE기침기관지염", "기타(IPF, 암)", "예정"]
        
        all_dfs = [] # 데이터프레임을 모을 리스트
        
        # 2. 각 시트를 순서대로 읽어서 리스트에 추가
        for sheet in target_sheets:
            try:
                # 시트별로 데이터 읽기 (모두 문자로 변환)
                temp_df = pd.read_excel(CRITERIA_FILE, sheet_name=sheet).astype(str)
                
                # 어떤 탭에서 왔는지 구분을 위해 '분류' 컬럼 추가 (맨 앞에 삽입)
                temp_df.insert(0, "분류", sheet)
                
                all_dfs.append(temp_df)
            except ValueError:
                # 해당 시트가 없으면 건너뜀 (에러 방지)
                continue
        
        # 3. 모든 시트 데이터를 하나로 합치기
        if all_dfs:
            df = pd.concat(all_dfs, ignore_index=True)
            
            # 'nan' 문자열을 빈칸으로 정리
            df = df.replace("nan", "")

            # ---------------------------------------------------------
            # 4. 검색 기능
            # ---------------------------------------------------------
            col_search, col_view = st.columns([3, 1])
            
            with col_search:
                search_query = st.text_input("🔍 키워드 검색 (전체 탭 통합 검색)", placeholder="예: 천식, COPD, 호산구, 녹농균")
            
            if search_query:
                query = search_query.strip()
                # 모든 컬럼에서 검색 (대소문자 무시)
                mask = df.apply(lambda row: row.astype(str).str.contains(query, case=False).any(), axis=1)
                df_display = df[mask]
            else:
                df_display = df

            st.caption(f"총 **{len(df_display)}**건의 연구 기준이 표시됩니다. (검색 대상: {', '.join(target_sheets)})")

            # ---------------------------------------------------------
            # 5. 화면 표시
            # ---------------------------------------------------------
            with col_view:
                view_mode = st.radio("보기 모드", ["요약 보기", "전체 펼쳐보기"], index=0)

            if view_mode == "요약 보기":
                st.dataframe(
                    df_display, 
                    use_container_width=True, 
                    hide_index=True,
                    height=500
                )
            else:
                st.markdown("##### 👇 전체 내용 보기")
                st.table(df_display)
        else:
            st.warning("⚠️ 엑셀 파일은 있지만, 지정된 시트(탭)를 하나도 찾을 수 없습니다.")

    except Exception as e:
        st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
else:
    st.info("ℹ️ 상세 기준 파일(criteria.xlsx)이 업로드되지 않았습니다.")