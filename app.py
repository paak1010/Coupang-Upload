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
    st.caption("✅ 조건: 동일 ME코드 내 LOT 1종이면 입력, 다수면 공백")
    st.caption("✅ 1발주 = 1LOT 원칙 (부분/분할 할당 불가)")
    st.caption("Developed by Jay (Ace Mode 🚀)")

# ==========================================
# 메인 화면 디자인
# ==========================================
st.title("쿠팡 수주업로드 자동 입력 대시보드 (찐퉁 수정판)")
st.markdown("Mentholatum : Moving The Heart")

if uploaded_file:
    try:
        # 데이터 불러오기
        df_order = pd.read_excel(uploaded_file, sheet_name='서식(수주업로드)', header=0) 
        df_inv = pd.read_excel(uploaded_file, sheet_name='재고(마스크팩x10)', header=2, dtype={'화주LOT': str})

        # 필수 열 세팅
        for col in ['LOT', '유효일자', '할당상태', '유효일자 입고가능', '금일재고', '잔량']:
            if col not in df_order.columns:
                df_order[col] = ''

        # 쿠팡 유효기한 데이터 정제
        if '쿠팡 유효기한' in df_order.columns:
            df_order['쿠팡 유효기한_보존'] = pd.to_datetime(df_order['쿠팡 유효기한'], errors='coerce')
        else:
            df_order['쿠팡 유효기한_보존'] = pd.NaT

        # MECODE 및 상품명 정제
        df_order['MECODE'] = df_order['MECODE'].astype(str).str.strip().str.upper()
        df_inv['상품'] = df_inv['상품'].astype(str).str.strip().str.upper()
        df_inv['화주LOT'] = df_inv['화주LOT'].astype(str).str.strip()

        df_order['수량'] = pd.to_numeric(df_order['수량'], errors='coerce').fillna(0)
        df_inv['환산'] = pd.to_numeric(df_inv['환산'], errors='coerce').fillna(0)
        
        df_inv['유효일자'] = pd.to_datetime(df_inv['유효일자'], errors='coerce').dt.normalize()
        df_inv['유효일자_보존'] = df_inv['유효일자'].fillna(pd.Timestamp('2099-12-31'))

        # 재고 그룹핑 (FEFO 순 정렬)
        df_inv_sorted = df_inv.sort_values(by=['상품', '유효일자_보존', '환산'], ascending=[True, True, False])
        inv_grouped = df_inv_sorted.groupby(['상품', '유효일자_보존', '화주LOT']).agg({'환산': 'sum', '유효일자': 'first'}).reset_index()

        # 찐퉁 모드 할당 로직
        with st.spinner('쿠팡 납품 커트라인 검사 및 LOT 조건 매핑 중... 🚀'):
            for i, row in df_order.iterrows():
                mecode = row['MECODE']
                order_qty = row['수량']
                limit_date = row['쿠팡 유효기한_보존']
                
                if pd.isna(mecode) or mecode == 'NAN' or order_qty <= 0:
                    df_order.at[i, '할당상태'] = "제외"
                    continue
                
                # [1] 현재 이 상품(MECODE)에서 쿠팡 기한을 만족하는 모든 재고 확인
                mask = (inv_grouped['상품'] == mecode) & (inv_grouped['환산'] > 0)
                if pd.notna(limit_date):
                    mask &= (inv_grouped['유효일자'] >= limit_date)
                
                current_valid_inv = inv_grouped[mask]
                
                if current_valid_inv.empty:
                    df_order.at[i, '할당상태'] = '재고없음/기한미달'
                    continue

                # [2] 🔥 핵심 조건: 현재 남은 로트의 '종류'가 몇 개인가?
                # (유효일자가 같더라도 로트 번호가 다르면 다중 로트로 간주)
                unique_lots = current_valid_inv['화주LOT'].unique()
                unique_lots_count = len(unique_lots)
                
                # 할당 대상 (가장 유효기간 짧은 첫 번째 행)
                target_idx = current_valid_inv.index[0]
                inv_qty = inv_grouped.at[target_idx, '환산']
                inv_date = inv_grouped.at[target_idx, '유효일자']
                actual_lot = inv_grouped.at[target_idx, '화주LOT']

                # 수량 부족 체크
                if inv_qty < order_qty:
                    df_order.at[i, '할당상태'] = '단일LOT수량부족'
                    continue

                # 재고 차감
                inv_grouped.at[target_idx, '환산'] -= order_qty
                
                # [3] 🔥 결과 입력 (조건 반영)
                if unique_lots_count > 1:
                    # 섞여 있다면 -> LOT 비우고, 유효일자는 유효 재고 중 가장 짧은 것으로 통일
                    df_order.at[i, 'LOT'] = ''
                    df_order.at[i, '유효일자'] = current_valid_inv['유효일자'].min().strftime('%Y-%m-%d')
                else:
                    # 하나뿐이라면 -> 로트와 일자 그대로 입력
                    df_order.at[i, 'LOT'] = actual_lot
                    df_order.at[i, '유효일자'] = inv_date.strftime('%Y-%m-%d')

                df_order.at[i, '할당상태'] = "정상할당"
                df_order.at[i, '유효일자 입고가능'] = True
                df_order.at[i, '금일재고'] = int(inv_qty)
                df_order.at[i, '잔량'] = int(inv_grouped.at[target_idx, '환산'])

        # 결과 출력
        df_order.drop(columns=['쿠팡 유효기한_보존'], inplace=True, errors='ignore')
        st.success("✅ 로트 조건(단일 유지 / 다수 공백) 반영 완료!")
        st.dataframe(df_order.head(20), use_container_width=True)

        # 엑셀 다운로드
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
            
        st.download_button(
            label="💾 쿠팡 최종 완성 엑셀 다운로드",
            data=buffer.getvalue(),
            file_name="쿠팡_수주업로드_로트수정본.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    except Exception as e:
        st.error(f"오류 발생: {e}")
