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
    st.caption("✅ 규칙: 동일 ME코드 내 LOT 1종이면 유지, 다수면 LOT 공백")
    st.caption("✅ 유효일자: 다수 LOT일 경우 가장 빠른 날짜로 통일")
    st.caption("Developed by Jay (Ace Mode 🚀)")

st.title("쿠팡 수주업로드 자동 입력 대시보드 (찐퉁 규칙 완벽 반영)")
st.markdown("Mentholatum : Moving The Heart")

if uploaded_file:
    try:
        # 데이터 불러오기 (LOT 번호 문자열 보존)
        df_order = pd.read_excel(uploaded_file, sheet_name='서식(수주업로드)', header=0) 
        df_inv = pd.read_excel(uploaded_file, sheet_name='재고(마스크팩x10)', header=2, dtype={'화주LOT': str})

        # 필수 열 세팅
        for col in ['LOT', '유효일자', '할당상태', '유효일자 입고가능', '금일재고', '잔량']:
            df_order[col] = ''

        # 날짜 및 텍스트 데이터 정제
        if '쿠팡 유효기한' in df_order.columns:
            df_order['쿠팡 유효기한_보존'] = pd.to_datetime(df_order['쿠팡 유효기한'], errors='coerce')
        else:
            df_order['쿠팡 유효기한_보존'] = pd.NaT

        df_order['MECODE'] = df_order['MECODE'].astype(str).str.strip().str.upper()
        df_inv['상품'] = df_inv['상품'].astype(str).str.strip().str.upper()
        df_inv['화주LOT'] = df_inv['화주LOT'].astype(str).str.strip()
        df_inv['유효일자'] = pd.to_datetime(df_inv['유효일자'], errors='coerce').dt.normalize()

        # [핵심] 재고 데이터 준비 (FEFO: 유효기한 짧은 순 정렬)
        inv_grouped = df_inv[df_inv['환산'] > 0].copy()
        inv_grouped = inv_grouped.sort_values(by=['상품', '유효일자', '환산'], ascending=[True, True, False])

        # 할당 로직 시작
        with st.spinner('서식 시트 규칙 분석 및 자동 할당 중... 🚀'):
            for i, row in df_order.iterrows():
                mecode = row['MECODE']
                order_qty = row['수량']
                limit_date = row['쿠팡 유효기한_보존']
                
                if pd.isna(mecode) or mecode == 'NAN' or order_qty <= 0:
                    df_order.at[i, '할당상태'] = "제외"
                    continue
                
                # 1. 현재 ME코드에 대해 쿠팡 기한을 만족하는 재고 필터링
                mask = (inv_grouped['상품'] == mecode) & (inv_grouped['환산'] > 0)
                if pd.notna(limit_date):
                    mask &= (inv_grouped['유효일자'] >= limit_date)
                
                current_stock = inv_grouped[mask]
                
                if current_stock.empty:
                    df_order.at[i, '할당상태'] = '재고없음/기한미달'
                    continue

                # 2. 🔥 사용자 규칙: 현재 사용 가능한 LOT 종류가 몇 개인가?
                unique_lots = current_stock['화주LOT'].unique()
                unique_lots_count = len(unique_lots)
                
                # 가장 우선순위 높은 재고 타겟
                target_idx = current_stock.index[0]
                inv_qty = inv_grouped.at[target_idx, '환산']
                inv_date = inv_grouped.at[target_idx, '유효일자']
                actual_lot = inv_grouped.at[target_idx, '화주LOT']

                # 수량 체크 (1발주 1LOT 원칙)
                if inv_qty < order_qty:
                    df_order.at[i, '할당상태'] = '단일LOT수량부족'
                    continue

                # 재고 차감
                inv_grouped.at[target_idx, '환산'] -= order_qty
                
                # 3. 🔥 [서식 시트 규칙 적용]
                if unique_lots_count > 1:
                    # 규칙 A: LOT가 섞여 있으면 LOT 칸은 비우고, 날짜는 유효 재고 중 가장 빠른 날짜로 통일
                    df_order.at[i, 'LOT'] = ''
                    df_order.at[i, '유효일자'] = current_stock['유효일자'].min().strftime('%Y-%m-%d')
                else:
                    # 규칙 B: LOT가 하나뿐이면 번호와 날짜를 그대로 입력
                    df_order.at[i, 'LOT'] = actual_lot
                    df_order.at[i, '유효일자'] = inv_date.strftime('%Y-%m-%d')

                # 나머지 정보 입력
                df_order.at[i, '할당상태'] = "정상할당"
                df_order.at[i, '유효일자 입고가능'] = True
                df_order.at[i, '금일재고'] = int(inv_qty)
                df_order.at[i, '잔량'] = int(inv_grouped.at[target_idx, '환산'])

        # 연산용 열 삭제 및 합계 계산
        df_order.drop(columns=['쿠팡 유효기한_보존'], inplace=True, errors='ignore')
        if '발주원가' in df_order.columns:
            df_order['발주원가'] = pd.to_numeric(df_order['발주원가'], errors='coerce').fillna(0)
            df_order['합계'] = df_order['수량'] * df_order['발주원가']

        st.success("✅ 서식 시트의 LOT/유효일자 규칙이 완벽하게 적용되었습니다.")

        # 미리보기 및 다운로드
        st.subheader("📊 할당 결과 미리보기")
        st.dataframe(df_order.head(20), use_container_width=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
            
        st.download_button(
            label="💾 쿠팡 최종 엑셀 다운로드 (에이스 등극! 🚀)",
            data=buffer.getvalue(),
            file_name="쿠팡_수주업로드_자동할당_완료.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    except Exception as e:
        st.error(f"데이터 처리 중 오류 발생: {e}")
