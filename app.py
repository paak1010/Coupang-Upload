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
    st.caption("💡 다중 LOT 보유 시 블라인드(공백) 및 최소유효일자 통일")
    st.caption("Developed by Jay (Ace Mode 🚀)")

# ==========================================
# 메인 화면 디자인
# ==========================================
st.title("쿠팡 수주업로드 자동 입력 대시보드 (찐퉁 모드 최종)")
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
                df_order[col] = None

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

        # 재고 그룹핑 (유효기간 촉박순 정렬 - FEFO)
        if not df_inv.empty:
            df_inv_sorted = df_inv.sort_values(by=['상품', '유효일자_보존', '환산'], ascending=[True, True, False])
            inv_grouped = df_inv_sorted.groupby(['상품', '유효일자_보존', '화주LOT']).agg({'환산': 'sum', '유효일자': 'first'}).reset_index()
        else:
            inv_grouped = pd.DataFrame(columns=['상품', '유효일자_보존', '화주LOT', '환산', '유효일자'])

        # 찐퉁 모드 할당 로직
        with st.spinner('쿠팡 납품 커트라인 검사 및 다중 LOT 블라인드 처리 중... 🚀'):
            for i, row in df_order.iterrows():
                mecode = row['MECODE']
                order_qty = row['수량']
                limit_date = row['쿠팡 유효기한_보존']
                
                if pd.isna(mecode) or str(mecode) == 'NAN' or order_qty <= 0:
                    df_order.at[i, '할당상태'] = "제외"
                    continue
                
                # 할당 시점 기준, 해당 상품의 '재고가 남아있는(환산>0)' 모든 LOT 조회
                available_inv_df = inv_grouped[(inv_grouped['상품'] == mecode) & (inv_grouped['환산'] > 0)]
                available_inv_idx = available_inv_df.index
                
                if len(available_inv_idx) == 0:
                    df_order.at[i, 'LOT'] = '재고없음'
                    df_order.at[i, '유효일자'] = '재고없음'
                    df_order.at[i, '할당상태'] = '재고없음'
                    df_order.at[i, '유효일자 입고가능'] = False
                    continue

                # 🔥 [추가된 로직] 환산 재고가 남아있는 유니크 LOT 개수와 가장 짧은 유효일자 파악
                unique_lots_count = available_inv_df['화주LOT'].nunique()
                shortest_valid_date = available_inv_df['유효일자'].min()
                
                is_allocated = False

                for idx in available_inv_idx:
                    inv_qty = inv_grouped.at[idx, '환산']
                    inv_date = inv_grouped.at[idx, '유효일자']
                    
                    # [조건 1] 1발주 1LOT: 수량을 모두 커버할 수 있어야 함
                    if inv_qty < order_qty:
                        continue
                        
                    # [조건 2] 쿠팡 유효기한 커트라인 통과해야 함
                    if pd.notna(limit_date) and pd.notna(inv_date):
                        if inv_date < limit_date:
                            continue 
                            
                    # 할당 확정 (실제 차감 처리)
                    actual_lot_str = inv_grouped.at[idx, '화주LOT']
                    actual_date_str = inv_date.strftime('%Y-%m-%d') if pd.notna(inv_date) else '일자없음'
                    
                    before_qty = inv_qty
                    after_qty = inv_qty - order_qty
                    
                    # 마스터 재고 실시간 차감
                    inv_grouped.at[idx, '환산'] = after_qty
                    
                    # 🔥 [핵심 기록 로직] 다중 LOT가 남아있다면 공백 처리 후 최단 날짜 적용
                    if unique_lots_count > 1:
                        display_lot = None  # 엑셀 기록 시 완전 빈칸(NaN)으로 남도록 None 처리
                        display_date = shortest_valid_date.strftime('%Y-%m-%d') if pd.notna(shortest_valid_date) else '일자없음'
                    else:
                        display_lot = actual_lot_str
                        display_date = actual_date_str
                    
                    # 데이터프레임 기록
                    df_order.at[i, 'LOT'] = display_lot
                    df_order.at[i, '유효일자'] = display_date
                    df_order.at[i, '할당상태'] = "정상할당"
                    df_order.at[i, '유효일자 입고가능'] = True
                    df_order.at[i, '금일재고'] = int(before_qty)
                    df_order.at[i, '잔량'] = int(after_qty)
                    
                    is_allocated = True
                    break

                if not is_allocated:
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

        st.success("✅ [쿠팡 찐퉁 모드 최종판] 다중 LOT 블라인드 처리 및 선입선출 최단일자 표기 완벽 적용!")

        # 결과 확인
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
            label="💾 쿠팡 최종 완성 엑셀 파일 다운로드 (찐퉁 마스터 버전)",
            data=buffer.getvalue(),
            file_name="쿠팡_수주업로드_자동할당완료_Master.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다: {e}")
