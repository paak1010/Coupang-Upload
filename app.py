import streamlit as st
import pandas as pd
import io

# 1. 페이지 기본 설정
st.set_page_config(page_title="쿠팡 수주업로드 자동 입력 시스템", page_icon="📦", layout="wide")

# ==========================================
# 🎨 사이드바 (Sidebar) 디자인
# ==========================================
with st.sidebar:
    st.image("https://static.wikia.nocookie.net/mycompanies/images/d/de/Fe328a0f-a347-42a0-bd70-254853f35374.jpg/revision/latest?cb=20191117172510", use_container_width=True)
    
    st.markdown("---")
    st.header("⚙️ 작업 설정")
    uploaded_file = st.file_uploader("쿠팡 발주 엑셀 업로드", type=['xlsx'])
    
    st.markdown("---")
    st.caption("💡 1발주 = 1LOT 원칙 (부분/분할 할당 불가)")
    st.caption("💡 쿠팡 유효기한(커트라인) 완벽 필터링 적용")
    st.caption("💡 금일재고 / 잔량 / 입고가능여부 자동 산출")
    st.caption("Developed by Jay (Ace Mode 🚀)")

# ==========================================
# 메인 화면 디자인
# ==========================================
st.title("쿠팡 수주업로드 자동 입력 대시보드 (찐퉁 모드)")
st.markdown("Mentholatum : Moving The Heart")

if uploaded_file:
    try:
        # 데이터 불러오기
        df_order = pd.read_excel(uploaded_file, sheet_name='서식(수주업로드)', header=0) 
        df_inv = pd.read_excel(uploaded_file, sheet_name='재고(마스크팩x10)', header=2)

        # 불필요한 열 날리기
        if '잔여일수' in df_order.columns:
            start_idx = list(df_order.columns).index('잔여일수')
            cols_to_drop = df_order.columns[start_idx:]
            df_order = df_order.drop(columns=cols_to_drop)

        # 찐퉁 엑셀 포맷을 위한 필수 열 세팅
        for col in ['LOT', '유효일자', '할당상태', '유효일자 입고가능', '금일재고', '잔량']:
            if col not in df_order.columns:
                df_order[col] = ''

        # [핵심] 쿠팡 유효기한 날짜 연산용 임시 보존 열 생성
        if '쿠팡 유효기한' in df_order.columns:
            df_order['쿠팡 유효기한_보존'] = pd.to_datetime(df_order['쿠팡 유효기한'], errors='coerce')
        else:
            df_order['쿠팡 유효기한_보존'] = pd.NaT

        # 데이터 정제
        df_order['MECODE'] = df_order['MECODE'].astype(str).str.replace(r'[^a-zA-Z0-9]', '', regex=True).str.upper()
        df_inv['상품'] = df_inv['상품'].astype(str).str.replace(r'[^a-zA-Z0-9]', '', regex=True).str.upper()

        df_order['수량'] = pd.to_numeric(df_order['수량'].astype(str).str.replace(r'[^\d.]', '', regex=True), errors='coerce').fillna(0)
        df_inv['환산'] = pd.to_numeric(df_inv['환산'].astype(str).str.replace(r'[^\d.]', '', regex=True), errors='coerce').fillna(0)
        
        df_inv['유효일자'] = pd.to_datetime(df_inv['유효일자'], errors='coerce').dt.normalize()
        df_inv['유효일자_보존'] = df_inv['유효일자'].fillna(pd.Timestamp('2099-12-31'))

        # OC2, PMM 조건 등 특정 상품 제한 제거 -> 모든 유효 재고 사용
        df_inv_valid = df_inv.copy()

        # 재고 그룹핑 (유효기간 촉박순 정렬 - FEFO)
        if not df_inv_valid.empty:
            df_inv_sorted = df_inv_valid.sort_values(by=['상품', '유효일자_보존', '환산'], ascending=[True, True, False])
            inv_grouped = df_inv_sorted.groupby(['상품', '유효일자_보존', '화주LOT']).agg({'환산': 'sum', '유효일자': 'first'}).reset_index()
        else:
            inv_grouped = pd.DataFrame(columns=['상품', '유효일자_보존', '화주LOT', '환산', '유효일자'])

        # 찐퉁 모드 할당 로직
        with st.spinner('쿠팡 납품 커트라인 검사 및 최적 LOT 매핑 중... 🚀'):
            for i, row in df_order.iterrows():
                mecode = row['MECODE']
                order_qty = row['수량']
                limit_date = row['쿠팡 유효기한_보존']  # 쿠팡 요구 최소 유통기한
                
                if pd.isna(mecode) or str(mecode) == 'NAN' or order_qty <= 0:
                    df_order.at[i, '할당상태'] = "제외"
                    continue
                
                available_inv_idx = inv_grouped[(inv_grouped['상품'] == mecode) & (inv_grouped['환산'] > 0)].index
                
                if len(available_inv_idx) == 0:
                    df_order.at[i, 'LOT'] = '재고없음'
                    df_order.at[i, '유효일자'] = '재고없음'
                    df_order.at[i, '할당상태'] = '재고없음'
                    df_order.at[i, '유효일자 입고가능'] = False
                    continue

                is_allocated = False

                for idx in available_inv_idx:
                    inv_qty = inv_grouped.at[idx, '환산']
                    inv_date = inv_grouped.at[idx, '유효일자']
                    
                    # [조건 1] 1발주 1LOT: 수량을 모두 커버할 수 있어야 함
                    if inv_qty < order_qty:
                        continue
                        
                    # [조건 2] 쿠팡 유효기한 필터링: 재고 유효기간이 쿠팡 커트라인보다 길거나 같아야 함
                    if pd.notna(limit_date) and pd.notna(inv_date):
                        if inv_date < limit_date:
                            continue # 쿠팡 기준 미달! 할당하지 않고 다음으로 임박한 재고로 넘어감
                            
                    # 모든 조건을 통과한 최적의 LOT 발견!
                    lot_str = inv_grouped.at[idx, '화주LOT']
                    date_str = inv_date.strftime('%Y-%m-%d') if pd.notna(inv_date) else '일자없음'
                    
                    before_qty = inv_qty
                    after_qty = inv_qty - order_qty
                    
                    # 마스터 재고 차감
                    inv_grouped.at[idx, '환산'] = after_qty
                    
                    # 찐퉁 폼에 맞춰 데이터프레임 기록
                    df_order.at[i, 'LOT'] = lot_str
                    df_order.at[i, '유효일자'] = date_str
                    df_order.at[i, '할당상태'] = "정상할당"
                    df_order.at[i, '유효일자 입고가능'] = True  # 찐퉁처럼 True 값 반환
                    df_order.at[i, '금일재고'] = int(before_qty)  # 차감 전 재고 기록
                    df_order.at[i, '잔량'] = int(after_qty)      # 차감 후 재고 기록
                    
                    is_allocated = True
                    break

                if not is_allocated:
                    # 수량이 안 맞거나, 남은 재고가 전부 쿠팡 유효기한 미달일 경우
                    df_order.at[i, 'LOT'] = '조건충족재고없음'
                    df_order.at[i, '유효일자'] = '조건충족재고없음'
                    df_order.at[i, '할당상태'] = '수량부족 or 기한미달'
                    df_order.at[i, '유효일자 입고가능'] = False
                    df_order.at[i, '금일재고'] = 0
                    df_order.at[i, '잔량'] = 0

        # 내부 연산용 열 삭제
        df_order.drop(columns=['쿠팡 유효기한_보존'], inplace=True, errors='ignore')

        if '발주원가' in df_order.columns:
            df_order['발주원가'] = pd.to_numeric(df_order['발주원가'], errors='coerce').fillna(0)
            df_order['합계'] = df_order['수량'] * df_order['발주원가']

        st.success("✅ [쿠팡 찐퉁 모드] 특정 상품 조건 해제 및 1LOT 원칙 매핑 완벽 적용!")

        # 결과 확인 미리보기에 찐퉁 양식 컬럼 추가
        st.subheader("📊 작업 결과 미리보기")
        preview_cols = ['MECODE', '수량', '쿠팡 유효기한', 'LOT', '유효일자', '유효일자 입고가능', '금일재고', '잔량', '할당상태']
        st.dataframe(df_order[[c for c in preview_cols if c in df_order.columns]].head(20), use_container_width=True)

        # 엑셀 다운로드 
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
            
            df_inv_left = inv_grouped[inv_grouped['환산'] > 0].copy()
            df_inv_left['유효일자'] = df_inv_left['유효일자'].dt.strftime('%Y-%m-%d')
            df_inv_left.drop(columns=['유효일자_보존'], inplace=True, errors='ignore')
            df_inv_left.rename(columns={'환산': '최종_남은재고'}, inplace=True)
            df_inv_left.to_excel(writer, index=False, sheet_name='작업후_잔여재고현황')
            
        st.download_button(
            label="💾 쿠팡 최종 완성 엑셀 파일 다운로드 (찐퉁 버전)",
            data=buffer.getvalue(),
            file_name="쿠팡_수주업로드_자동할당완료(찐퉁).xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다: {e}")
