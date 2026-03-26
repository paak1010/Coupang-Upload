import streamlit as st
import pandas as pd
import io

# 1. 페이지 기본 설정
st.set_page_config(page_title="쿠팡 수주업로드 자동 입력 시스템", page_icon="📦", layout="wide")

with st.sidebar:
    st.image("https://static.wikia.nocookie.net/mycompanies/images/d/de/Fe328a0f-a347-42a0-bd70-254853f35374.jpg/revision/latest?cb=20191117172510", use_container_width=True)
    st.markdown("---")
    st.header("⚙️ 작업 설정")
    uploaded_file = st.file_uploader("쿠팡 발주 엑셀 업로드", type=['xlsx'])
    st.markdown("---")
    st.caption("✅ 0개 주문 포함 모든 행 정직하게 처리")
    st.caption("✅ ME코드별 통합 분석 (단일 LOT 유지 / 다수 공백)")
    st.caption("Developed by Jay (Ace Mode 🚀)")

st.title("쿠팡 수주업로드 자동 입력 대시보드 (0개 조건 제거 버전)")
st.markdown("Mentholatum : Moving The Heart")

if uploaded_file:
    try:
        # 1. 데이터 불러오기
        df_order = pd.read_excel(uploaded_file, sheet_name='서식(수주업로드)', header=0) 
        df_inv = pd.read_excel(uploaded_file, sheet_name='재고(마스크팩x10)', header=2, dtype={'화주LOT': str})

        # 2. 데이터 타입 강제 정제 (오류 방지 핵심)
        # 수량/환산 컬럼에서 콤마(,) 등 문자를 제거하고 순수 숫자로 변환
        def clean_numeric(x):
            if pd.isna(x): return 0
            s = str(x).replace(',', '').strip()
            return pd.to_numeric(s, errors='coerce') or 0

        df_order['수량'] = df_order['수량'].apply(clean_numeric)
        df_inv['환산'] = df_inv['환산'].apply(clean_numeric)
        
        df_order['MECODE'] = df_order['MECODE'].astype(str).str.strip().str.upper()
        df_inv['상품'] = df_inv['상품'].astype(str).str.strip().str.upper()
        df_inv['화주LOT'] = df_inv['화주LOT'].astype(str).str.strip()
        
        # 날짜 정제
        df_inv['유효일자'] = pd.to_datetime(df_inv['유효일자'], errors='coerce').dt.normalize()
        if '쿠팡 유효기한' in df_order.columns:
            df_order['쿠팡 유효기한_보존'] = pd.to_datetime(df_order['쿠팡 유효기한'], errors='coerce')
        else:
            df_order['쿠팡 유효기한_보존'] = pd.NaT

        # 결과 열 초기화
        for col in ['LOT', '유효일자', '할당상태', '유효일자 입고가능', '금일재고', '잔량']:
            df_order[col] = ''

        # 재고 그룹화 (FEFO)
        inv_grouped = df_inv[df_inv['환산'] > 0].groupby(['상품', '유효일자', '화주LOT'])['환산'].sum().reset_index()
        inv_grouped = inv_grouped.sort_values(by=['상품', '유효일자', '환산'], ascending=[True, True, False])

        # 3. ME코드별 통합 분석 및 할당
        with st.spinner('전체 행 분석 중... 🚀'):
            # NAN 제외한 실제 존재하는 ME코드들 추출
            valid_mecodes = [m for m in df_order['MECODE'].unique() if m not in ['NAN', 'NONE', '']]
            
            for mecode in valid_mecodes:
                order_indices = df_order[df_order['MECODE'] == mecode].index
                total_needed = df_order.loc[order_indices, '수량'].sum()
                
                # 해당 ME코드 재고 필터 (쿠팡 기한 만족 기준)
                # 대표 기한은 첫 번째 행 기준 (보통 한 품목은 기한이 같음)
                rep_date = df_order.loc[order_indices[0], '쿠팡 유효기한_보존']
                
                stock_pool = inv_grouped[inv_grouped['상품'] == mecode].copy()
                if pd.notna(rep_date):
                    stock_pool = stock_pool[stock_pool['유효일자'] >= rep_date]

                # --- 판단 로직 ---
                if stock_pool.empty:
                    df_order.loc[order_indices, '할당상태'] = '재고없음/기한미달'
                    continue

                # 전체 수량을 채울 때 사용되는 로트 종류 파악
                cum_stock = stock_pool['환산'].cumsum()
                # 수량이 0이어도 최소한 첫 번째 로트는 체크되도록 설계
                needed_mask = (cum_stock.shift(1).fillna(0) < max(total_needed, 0.1))
                used_lots_info = stock_pool[needed_mask]
                
                lot_count = used_lots_info['화주LOT'].nunique()
                target_date = used_lots_info['유효일자'].min().strftime('%Y-%m-%d')
                target_lot = used_lots_info['화주LOT'].iloc[0] if lot_count == 1 else ''

                # 각 행에 일괄 기입 (수량 0인 행 포함)
                for idx in order_indices:
                    row_qty = df_order.at[idx, '수량']
                    
                    # 로트 기입 결정 (섞여있으면 공백)
                    df_order.at[idx, 'LOT'] = target_lot
                    df_order.at[idx, '유효일자'] = target_date
                    df_order.at[idx, '할당상태'] = "정상할당"
                    df_order.at[idx, '유효일자 입고가능'] = True

                    # 재고 차감 (실시간 계산용 - 수량이 0이면 차감 안됨)
                    if row_qty > 0:
                        temp_qty = row_qty
                        for s_idx in stock_pool.index:
                            if temp_qty <= 0: break
                            avail = inv_grouped.at[s_idx, '환산']
                            take = min(avail, temp_qty)
                            inv_grouped.at[s_idx, '환산'] -= take
                            temp_qty -= take

        # 후처리
        df_order.drop(columns=['쿠팡 유효기한_보존'], inplace=True, errors='ignore')
        st.success("✅ 모든 행(수량 0 포함)에 대해 로트 규칙 적용 완료!")
        st.dataframe(df_order.head(20), use_container_width=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
            
        st.download_button(
            label="💾 쿠팡 최종 엑셀 다운로드",
            data=buffer.getvalue(),
            file_name="쿠팡_수주업로드_전체행반영.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    except Exception as e:
        st.error(f"데이터 처리 중 오류 발생: {e}")
