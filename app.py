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
    st.caption("💡 낱개 단위 할당 및 수량 부족 시 LOT 행 분할 적용 (임박순)")
    st.caption("💡 실시간 재고 차감 및 잔여 재고 산출")
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

        if 'LOT' not in df_order.columns: df_order['LOT'] = ''
        if '유효일자' not in df_order.columns: df_order['유효일자'] = ''

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

        # 재고 그룹핑 (유효기간 촉박순 정렬)
        if not df_inv_valid.empty:
            df_inv_sorted = df_inv_valid.sort_values(by=['상품', '유효일자_보존', '환산'], ascending=[True, True, False])
            inv_grouped = df_inv_sorted.groupby(['상품', '유효일자_보존', '화주LOT']).agg({'환산': 'sum', '유효일자': 'first'}).reset_index()
        else:
            inv_grouped = pd.DataFrame(columns=['상품', '유효일자_보존', '화주LOT', '환산', '유효일자'])

        final_order_rows = []

        # 할당 및 실시간 재고 차감 로직
        with st.spinner('실시간 재고 차감 및 최적 LOT 분할 중... 🚀'):
            for i, row in df_order.iterrows():
                mecode = row['MECODE']
                order_qty = row['수량']
                
                if pd.isna(mecode) or str(mecode) == 'NAN' or order_qty <= 0:
                    row['할당상태'] = "제외"
                    final_order_rows.append(row)
                    continue
                
                available_inv_idx = inv_grouped[(inv_grouped['상품'] == mecode) & (inv_grouped['환산'] > 0)].index
                
                if len(available_inv_idx) == 0:
                    row['LOT'] = '재고없음'
                    row['유효일자'] = '재고없음'
                    row['할당상태'] = '재고없음'
                    row['할당후_LOT잔여재고'] = 0
                    final_order_rows.append(row)
                    continue

                remaining_qty = order_qty

                for idx in available_inv_idx:
                    if remaining_qty <= 0:
                        break 
                        
                    inv_qty = inv_grouped.at[idx, '환산']
                    lot_str = inv_grouped.at[idx, '화주LOT']
                    date_str = inv_grouped.at[idx, '유효일자'].strftime('%Y-%m-%d') if pd.notna(inv_grouped.at[idx, '유효일자']) else '일자없음'
                    
                    if inv_qty <= 0:
                        continue

                    take_qty = min(remaining_qty, inv_qty)
                    
                    # 🔥 [핵심] 여기서 마스터 재고가 실시간으로 차감됩니다!
                    inv_grouped.at[idx, '환산'] -= take_qty
                    remaining_qty -= take_qty
                    
                    # 현재 행 복사 및 기입
                    new_row = row.copy()
                    new_row['수량'] = int(take_qty)  # 깔끔하게 정수로
                    new_row['LOT'] = lot_str
                    new_row['유효일자'] = date_str
                    
                    # 방금 차감하고 남은 해당 LOT의 '찐' 잔여 재고 기록 (눈으로 확인 가능!)
                    new_row['할당후_LOT잔여재고'] = int(inv_grouped.at[idx, '환산'])
                    
                    if take_qty == order_qty:
                        new_row['할당상태'] = "정상할당"
                    else:
                        new_row['할당상태'] = "분할할당"

                    final_order_rows.append(new_row)

                if remaining_qty > 0:
                    short_row = row.copy()
                    short_row['수량'] = int(remaining_qty)
                    short_row['LOT'] = '재고부족'
                    short_row['유효일자'] = '재고부족'
                    short_row['할당상태'] = '재고부족'
                    short_row['할당후_LOT잔여재고'] = 0
                    final_order_rows.append(short_row)

        df_final_order = pd.DataFrame(final_order_rows)

        # 금액 재계산
        if '발주원가' in df_final_order.columns and '합계' in df_final_order.columns:
            df_final_order['발주원가'] = pd.to_numeric(df_final_order['발주원가'], errors='coerce').fillna(0)
            df_final_order['합계'] = df_final_order['수량'] * df_final_order['발주원가']

        st.success("✅ [쿠팡 모드] LOT 분할 및 임박일자 최적 할당 완료! (실시간 재고 차감 완벽 적용)")

        # 결과 확인
        st.subheader("📊 작업 결과 미리보기")
        preview_cols = ['MECODE', '상품명', '수량', 'LOT', '유효일자', '할당상태', '할당후_LOT잔여재고']
        st.dataframe(df_final_order[[c for c in preview_cols if c in df_final_order.columns]].head(20), use_container_width=True)

        # 엑셀 다운로드 (작업 후 남은 재고 시트까지 추가!)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            # 1. 수주업로드 시트
            df_final_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
            
            # 2. 작업이 모두 끝난 후 남은 재고 시트 (센스 추가)
            df_inv_left = inv_grouped[inv_grouped['환산'] > 0].copy()
            df_inv_left['유효일자'] = df_inv_left['유효일자'].dt.strftime('%Y-%m-%d')
            df_inv_left.drop(columns=['유효일자_보존'], inplace=True, errors='ignore')
            df_inv_left.rename(columns={'환산': '최종_남은재고'}, inplace=True)
            df_inv_left.to_excel(writer, index=False, sheet_name='작업후_잔여재고현황')
            
        st.download_button(
            label="💾 쿠팡 최종 완성 엑셀 파일 다운로드",
            data=buffer.getvalue(),
            file_name="쿠팡_수주업로드_자동할당완료.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다: {e}")
