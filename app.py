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
    st.caption("💡 임박일자 최우선 & 실시간 재고 차감")
    st.caption("Developed by Jay (Ace Mode 🚀)")

# ==========================================
# 메인 화면 디자인
# ==========================================
st.title("쿠팡 수주업로드 자동 입력 대시보드")
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

        # 상태 기록용 열 추가
        if 'LOT' not in df_order.columns: df_order['LOT'] = ''
        if '유효일자' not in df_order.columns: df_order['유효일자'] = ''
        df_order['할당상태'] = ''
        df_order['할당후_LOT잔여재고'] = None

        # 데이터 정제
        df_order['MECODE'] = df_order['MECODE'].astype(str).str.replace(r'[^a-zA-Z0-9]', '', regex=True).str.upper()
        df_inv['상품'] = df_inv['상품'].astype(str).str.replace(r'[^a-zA-Z0-9]', '', regex=True).str.upper()

        df_order['수량'] = pd.to_numeric(df_order['수량'].astype(str).str.replace(r'[^\d.]', '', regex=True), errors='coerce').fillna(0)
        df_inv['환산'] = pd.to_numeric(df_inv['환산'].astype(str).str.replace(r'[^\d.]', '', regex=True), errors='coerce').fillna(0)
        
        df_inv['유효일자'] = pd.to_datetime(df_inv['유효일자'], errors='coerce').dt.normalize()
        df_inv['유효일자_보존'] = df_inv['유효일자'].fillna(pd.Timestamp('2099-12-31'))

        # 불량 재고 걸러내기
        idx_pmm = (df_inv['상품'] == 'ME00621PMM') & (df_inv['유효일자'].dt.year != 2028)
        idx_oc2 = (df_inv['상품'] == 'ME90621OC2') & (~df_inv['화주LOT'].fillna('').astype(str).str.contains('분리배출'))
        df_inv_valid = df_inv[~(idx_pmm | idx_oc2)].copy()

        # 재고 그룹핑 (유효기간 촉박순 정렬 - FEFO)
        if not df_inv_valid.empty:
            df_inv_sorted = df_inv_valid.sort_values(by=['상품', '유효일자_보존', '환산'], ascending=[True, True, False])
            inv_grouped = df_inv_sorted.groupby(['상품', '유효일자_보존', '화주LOT']).agg({'환산': 'sum', '유효일자': 'first'}).reset_index()
        else:
            inv_grouped = pd.DataFrame(columns=['상품', '유효일자_보존', '화주LOT', '환산', '유효일자'])

        # 단일 LOT 할당 로직
        with st.spinner('실시간 재고 차감 및 단일 LOT 최적 할당 중... 🚀'):
            for i, row in df_order.iterrows():
                mecode = row['MECODE']
                order_qty = row['수량']
                
                # 수량이 0이거나 잘못된 데이터 패스
                if pd.isna(mecode) or str(mecode) == 'NAN' or order_qty <= 0:
                    df_order.at[i, '할당상태'] = "제외"
                    continue
                
                # 해당 상품의 남은 재고 인덱스 추출
                available_inv_idx = inv_grouped[(inv_grouped['상품'] == mecode) & (inv_grouped['환산'] > 0)].index
                
                if len(available_inv_idx) == 0:
                    df_order.at[i, 'LOT'] = '재고없음'
                    df_order.at[i, '유효일자'] = '재고없음'
                    df_order.at[i, '할당상태'] = '재고없음'
                    continue

                is_allocated = False

                # 임박순으로 정렬된 재고를 돌면서 "주문 수량을 100% 감당할 수 있는 LOT" 탐색
                for idx in available_inv_idx:
                    inv_qty = inv_grouped.at[idx, '환산']
                    
                    # 🔥 [핵심] 하나의 LOT가 발주 수량 이상을 가지고 있어야만 진행
                    if inv_qty >= order_qty:
                        lot_str = inv_grouped.at[idx, '화주LOT']
                        date_str = inv_grouped.at[idx, '유효일자'].strftime('%Y-%m-%d') if pd.notna(inv_grouped.at[idx, '유효일자']) else '일자없음'
                        
                        # 마스터 재고 차감
                        inv_grouped.at[idx, '환산'] -= order_qty
                        
                        # 데이터프레임 기록
                        df_order.at[i, 'LOT'] = lot_str
                        df_order.at[i, '유효일자'] = date_str
                        df_order.at[i, '할당상태'] = "정상할당"
                        df_order.at[i, '할당후_LOT잔여재고'] = int(inv_grouped.at[idx, '환산'])
                        
                        is_allocated = True
                        break # 성공적으로 할당했으므로 다음 주문으로 넘어감 (루프 탈출)

                # 모든 LOT를 다 돌았는데도 이 주문을 한 번에 감당할 LOT가 없는 경우
                if not is_allocated:
                    df_order.at[i, 'LOT'] = '단일LOT수량부족'
                    df_order.at[i, '유효일자'] = '단일LOT수량부족'
                    df_order.at[i, '할당상태'] = '수량부족(부분할당불가)'

        # 금액 재계산
        if '발주원가' in df_order.columns:
            df_order['발주원가'] = pd.to_numeric(df_order['발주원가'], errors='coerce').fillna(0)
            df_order['합계'] = df_order['수량'] * df_order['발주원가']

        st.success("✅ [쿠팡 모드] 1발주 1LOT 원칙으로 매핑 및 임박일자 최적 할당 완료!")

        # 결과 확인
        st.subheader("📊 작업 결과 미리보기")
        preview_cols = ['MECODE', '상품명', '수량', 'LOT', '유효일자', '할당상태', '할당후_LOT잔여재고']
        st.dataframe(df_order[[c for c in preview_cols if c in df_order.columns]].head(20), use_container_width=True)

        # 엑셀 다운로드 
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            # 1. 수주업로드 시트
            df_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
            
            # 2. 작업이 모두 끝난 후 남은 재고 현황 시트
            df_inv_left = inv_grouped[inv_grouped['환산'] > 0].copy()
            df_inv_left['유효일자'] = df_inv_left['유효일자'].dt.strftime('%Y-%m-%d')
            df_inv_left.drop(columns=['유효일자_보존'], inplace=True, errors='ignore')
            df_inv_left.rename(columns={'환산': '최종_남은재고'}, inplace=True)
            df_inv_left.to_excel(writer, index=False, sheet_name='작업후_잔여재고현황')
            
        st.download_button(
            label="💾 쿠팡 최종 완성 엑셀 파일 다운로드",
            data=buffer.getvalue(),
            file_name="쿠팡_수주업로드_자동할당완료(단일LOT).xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다: {e}")
