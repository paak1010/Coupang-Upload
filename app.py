# 찐퉁 모드 할당 로직
        with st.spinner('쿠팡 납품 커트라인 검사 및 다중 LOT 조건 최적화 중... 🚀'):
            for i, row in df_order.iterrows():
                mecode = row['MECODE']
                order_qty = row['수량']
                limit_date = row['쿠팡 유효기한_보존']
                
                if pd.isna(mecode) or str(mecode) == 'NAN' or order_qty <= 0:
                    df_order.at[i, '할당상태'] = "제외"
                    continue
                
                # 1. 현재 이 상품(MECODE)에서 쿠팡 가이드라인을 통과하는 재고만 필터링
                mask = (inv_grouped['상품'] == mecode) & (inv_grouped['환산'] > 0)
                if pd.notna(limit_date):
                    mask &= (inv_grouped['유효일자'] >= limit_date)
                
                valid_stock = inv_grouped[mask]
                
                if valid_stock.empty:
                    df_order.at[i, 'LOT'] = '조건충족재고없음'
                    df_order.at[i, '유효일자'] = '조건충족재고없음'
                    df_order.at[i, '할당상태'] = '수량부족 or 기한미달'
                    continue

                # 2. 할당 가능한 첫 번째 재고 선택 (FEFO 순서대로 정렬되어 있음)
                idx = valid_stock.index[0]
                inv_qty = inv_grouped.at[idx, '환산']
                inv_date = inv_grouped.at[idx, '유효일자']
                actual_lot_str = inv_grouped.at[idx, '화주LOT']

                # 수량 부족 체크
                if inv_qty < order_qty:
                    df_order.at[i, '할당상태'] = '단일LOT수량부족'
                    continue

                # 3. 🔥 핵심 로직: 현재 '가용한 전체 LOT 종류'가 몇 개인가?
                # (이 시점에 남은 LOT가 딱 하나라면 그 번호를 그대로 써야 함)
                current_unique_lots = valid_stock['화주LOT'].nunique()

                # 재고 차감
                before_qty = inv_qty
                after_qty = inv_qty - order_qty
                inv_grouped.at[idx, '환산'] = after_qty
                
                # 4. 출력값 결정
                if current_unique_lots > 1:
                    # 여러 LOT가 경합 중이면 안전하게 공백 처리 (최단유효일자 기준)
                    display_lot = '' 
                    display_date = valid_stock['유효일자'].min().strftime('%Y-%m-%d')
                else:
                    # LOT가 하나뿐이거나, 남은 게 이거 하나라면 그대로 기입
                    display_lot = actual_lot_str
                    display_date = inv_date.strftime('%Y-%m-%d') if pd.notna(inv_date) else ''

                # 결과 입력
                df_order.at[i, 'LOT'] = display_lot
                df_order.at[i, '유효일자'] = display_date
                df_order.at[i, '할당상태'] = "정상할당"
                df_order.at[i, '유효일자 입고가능'] = True
                df_order.at[i, '금일재고'] = int(before_qty)
                df_order.at[i, '잔량'] = int(after_qty)
