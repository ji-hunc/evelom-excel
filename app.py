import streamlit as st
import io
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import PatternFill, Alignment
from collections import Counter
from collections import defaultdict
from openpyxl.styles import Border, Side



st.title("Excel 처리 앱")
st.write("""
1️⃣ 엑셀 파일을 업로드하면,  
2️⃣ 가공된 파일을 다운로드할 수 있어요!
""")

uploaded_file = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"], accept_multiple_files=False)

if uploaded_file:
    if st.button("✅ 처리 시작하기"):
        try:
            # 파일명에서 날짜 추출하여 결과 파일명 설정
            uploaded_filename = uploaded_file.name
            match = re.search(r"\d{8}", uploaded_filename)
            if match:
                extracted_date = match.group()
                output_filename = f"결과_{extracted_date}.xlsx"
            else:
                output_filename = "processed_excel.xlsx"

            progress_bar = st.progress(0)
            st.info("파일 처리 중... 잠시만 기다려주세요.")

            # 1. openpyxl로 파일 열기
            wb = load_workbook(uploaded_file)
            ws = wb.active
            progress_bar.progress(10)

            # 2. 전체 셀 배경색 제거 (1행은 제외)
            no_fill = PatternFill(fill_type=None)
            for row_idx, row in enumerate(ws.iter_rows(), start=1):
                if row_idx == 1:
                    continue  # 첫 번째 행은 건드리지 않음
                for cell in row:
                    cell.fill = no_fill
            progress_bar.progress(20)

            # 3. 열 너비 조정
            ws.column_dimensions['A'].width = 12
            ws.column_dimensions['C'].width = 73
            ws.column_dimensions['D'].width = 5
            ws.column_dimensions['F'].width = 25
            ws.column_dimensions['H'].width = 7
            ws.column_dimensions['I'].width = 25
            ws.column_dimensions['E'].width = 25
            ws.column_dimensions['K'].width = 80
            progress_bar.progress(30)

            # 4. B열(수취인명) 기준으로 빈 행 추가 및 회색으로 채우기 + 행 높이 조정
            light_gray_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
            max_col_to_fill = max(ws.max_column, 50)
            if ws.max_row > 1:
                for r in range(ws.max_row, 1, -1):
                    current_recipient = ws.cell(row=r, column=2).value
                    previous_recipient = ws.cell(row=r - 1, column=2).value

                    if current_recipient is not None and previous_recipient is not None and current_recipient != previous_recipient:
                        if (r - 1) > 1:
                            ws.insert_rows(r)

                            # 👉 새로 삽입된 행의 높이 설정 (예: 30)
                            # ws.row_dimensions[r + 1].height = 30

                            for col_idx in range(1, max_col_to_fill + 1):
                                ws.cell(row=r, column=col_idx).fill = light_gray_fill
            progress_bar.progress(60)

            # 5. D열과 H열 수량 확인 후 핑크색으로 칠하고, 가운데 정렬 적용 (1행, 빈 행 제외)
            pink_fill = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")
            center_align = Alignment(horizontal="center", vertical="center")

            for row_idx in range(2, ws.max_row + 1):
                if ws.cell(row=row_idx, column=2).value is None:
                    continue  # 수취인명 비어있으면 건너뜀

                indices_to_check = [4, 8]  # D열, H열
                for col_idx in indices_to_check:
                    cell = ws.cell(row=row_idx, column=col_idx)
                    value = cell.value
                    if value is None:
                        continue

                    # 👉 가운데 정렬 적용
                    cell.alignment = center_align

                    try:
                        numeric_value = float(str(value).strip())
                        if numeric_value >= 2:
                            cell.fill = pink_fill
                    except Exception:
                        continue

            progress_bar.progress(90)

            # 6. C열(상품명)과 D열(수량)을 기준으로 판매량 합산
            product_sales = defaultdict(float)

            for row_idx in range(2, ws.max_row + 1):
                product = ws.cell(row=row_idx, column=3).value  # C열: 상품명
                quantity = ws.cell(row=row_idx, column=4).value  # D열: 수량

                if product is None or str(product).strip() == "":
                    continue

                try:
                    quantity_num = float(str(quantity).strip()) if quantity is not None else 0
                except Exception:
                    quantity_num = 0

                product_sales[str(product).strip()] += quantity_num

            # 기존 데이터 마지막 행에서 한 칸 띄운 후 출력 시작
            summary_start_row = ws.max_row + 2
            # ws.cell(row=summary_start_row - 1, column=3).value = "상품명별 총 수량"  # 제목 행

            for i, (product, total_qty) in enumerate(product_sales.items()):
                ws.cell(row=summary_start_row + i, column=3).value = product
                ws.cell(row=summary_start_row + i, column=4).value = total_qty
            
            # 테두리 추가
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )

            for i in range(len(product_sales)):
                row = summary_start_row + i
                for col in [3, 4]:  # C열, D열
                    cell = ws.cell(row=row, column=col)
                    cell.border = thin_border
                    cell2 = ws.cell(row=row, column=col + 1)
                    cell2.alignment = center_align
                    ws.row_dimensions[row].height = 25

            # 6. 최종 파일 저장
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            progress_bar.progress(100)

            st.success("🎉 처리 완료! 아래 버튼으로 다운로드하세요.")
            st.download_button(
                label="📥 가공된 엑셀 다운로드",
                data=output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"⚠️ 처리 중 오류 발생: {e}")
