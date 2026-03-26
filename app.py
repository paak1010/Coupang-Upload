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
    st.caption("✅ 동일 ME코드 내 로트 1종: LOT/유효일자 전체 입력")
    st.caption("✅ 동일 ME코드 내 로트 다수: LOT 공백/최단일자 통일")
    st.caption("Developed by Jay (Ace Mode 🚀)")

st.title("쿠팡 수주업로드 자동 입력 대시보드 (로트 매칭 찐최종)")
st.markdown("Mentholatum : Moving The Heart")

if uploaded_file:
    try:
        # 데이터 불러오기
        df_order = pd.read_excel(uploaded_file, sheet_name='서식(수주업로드)', header=0) 
        df_inv = pd.read_excel(uploaded_file, sheet_name='재고(마스크팩x10)', header=2, dtype={'화주LOT': str})

        # 필수 열 세팅
        for col in ['LOT', '유효일자', '할당상태', '유효일자 입고가능', '금일재고', '잔량']:
            df_order[col] = ''

        # 유효기한 데이터 정제
        if '쿠팡 유효기한' in df_order.columns:
            df_order['쿠팡 유효기한_보존'] = pd.to_datetime(df_order['쿠팡 유효기한'], errors='coerce')
        else:
            df_order['쿠팡 유효기한_보존'] = pd.NaT

        df_order['MECODE'] = df_order['MECODE'].astype(str).str.strip().str.upper()
        df_inv['상품'] = df_inv['상품'].astype(str).str.strip().str.upper()
        df_inv['유효일자'] = pd.to_datetime(df_inv['유효일자'], errors='coerce').dt.normalize()

        # [핵심] 재고 그룹핑 (미리 상품별 로트 종류 수를 파악)
        # 쿠팡 커트라인에 걸리지 않는 '살아있는 재고'만 대상으로 함
        inv_grouped = df_inv[df_inv['환산'] > 0].copy()

        # 찐퉁 모드 할당 로직 시작
        with st.spinner('조건 검사 중... (단일 로트 유지 vs 다중 로트 공백) 🚀'):
            for i, row in df_order.iterrows():
                mecode = row['MECODE']
                order_qty = row['수량']
                limit_date = row['쿠팡 유효기한_보존']
                
                if pd.isna(mecode) or mecode == 'NAN' or order_qty <= 0:
                    df_order.at[i, '할당상태'] = "제외"
                    continue
                
                # 1. 현재 이 상품의 가용한 재고 필터링 (유효기간 조건 포함)
                current_stock = inv_grouped[inv_grouped['상품'] == mecode].copy()
                if pd.notna(limit_date):
                    current_stock = current_stock[current_stock['유효일자'] >= limit_date]
                
                if current_stock.empty:
                    df_order.at[i, '할당상태'] = '재고없음/기한미달'
                    continue

                # 2. 🔥 [사용자 요청 조건] 로트 종류 수 체크
                unique_lots = current_stock['화주LOT'].unique()
                unique_lots_count = len(unique_lots)
                shortest_date = current_stock['유효일자'].min()

                # 3. 할당 처리 (FEFO 순서대로 첫 번째 재고에서 차감)
                current_stock = current_stock.sort_values(by=['유효일자', '환산'], ascending=[True, False])
                target_idx = current_stock.index[0]
                
                if inv_grouped.at[target_idx, '환산'] < order_qty:
                    df_order.at[i, '할당상태'] = '단일LOT수량부족'
                    continue

                # 실제 재고 차감
                inv_grouped.at[target_idx, '환산'] -= order_qty
                
                # 4. 🔥 결과 입력 로직
                if unique_lots_count > 1:
                    # 조건 A: 다른 LOT가 하나라도 섞여 있다면 -> LOT 지우고 일자만 통일
                    df_order.at[i, 'LOT'] = '' 
                    df_order.at[i, '유효일자'] = shortest_date.strftime('%Y-%m-%d')
                else:
                    # 조건 B: LOT와 유효일자가 모두 동일한 상품들(즉 로트가 1종) -> 그대로 입력
                    df_order.at[i, 'LOT'] = current_stock.at[target_idx, '화주LOT']
                    df_order.at[i, '유효일자'] = current_stock.at[target_idx, '유효일자'].strftime('%Y-%m-%d')

                df_order.at[i, '할당상태'] = "정상할당"
                df_order.at[i, '유효일자 입고가능'] = True
                df_order.at[i, '금일재고'] = int(current_stock.at[target_idx, '환산'])
                df_order.at[i, '잔량'] = int(inv_grouped.at[target_idx, '환산'])

        # 후처리 및 다운로드
        df_order.drop(columns=['쿠팡 유효기한_보존'], inplace=True, errors='ignore')
        st.success("✅ 조건 반영 완료: 로트 1종은 유지, 다수 로트는 공백 및 일자 통일!")
        st.dataframe(df_order.head(20), use_container_width=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
        
        st.download_button(
            label="💾 쿠팡 최종 완성 엑셀 다운로드",
            data=buffer.getvalue(),
            file_name="쿠팡_수주업로드_로트조건반영.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"오류 발생: {e}")
