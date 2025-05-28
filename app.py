import streamlit as st
import pandas as pd
import io
import msoffcrypto
import time
import re
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill
from collections import Counter

st.title("🔐 Excel 처리 앱")
st.write("""
1️⃣ 엑셀 파일을 업로드하고,  
2️⃣ 비밀번호를 입력하면,  
3️⃣ 가공된 파일을 다운로드할 수 있어요!
""")

uploaded_file = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"], accept_multiple_files=False)
password = st.text_input("비밀번호를 입력하세요", type="password")
progress_bar = st.progress(0)

if uploaded_file and password:
    if st.button("✅ 처리 시작하기"):
        try:
            # ✅ 업로드된 파일명에서 날짜 추출
            uploaded_filename = uploaded_file.name
            match = re.search(r"\d{8}", uploaded_filename)
            if match:
                extracted_date = match.group()
                output_filename = f"양식_{extracted_date}.xlsx"
            else:
                output_filename = "result.xlsx"

            progress_bar.progress(10)

            # 1️⃣ 비밀번호로 복호화
            office_file = msoffcrypto.OfficeFile(uploaded_file)
            office_file.load_key(password=password)
            decrypted = io.BytesIO()
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            progress_bar.progress(30)

            # 2️⃣ pandas로 읽기
            df = pd.read_excel(decrypted, engine="openpyxl")
            progress_bar.progress(50)

            # 3️⃣ 첫 번째 행 삭제
            df = df.iloc[1:, :].reset_index(drop=True)
            progress_bar.progress(70)

            # 4️⃣ 필요한 7개의 열만 선택
            needed_columns_idx = [12, 19, 25, 47, 49, 53, 54]
            needed_columns = [df.columns[idx] for idx in needed_columns_idx]
            df = df[needed_columns]

            # 5️⃣ 열 이름 새로 지정
            df.columns = [
                "수취인명",
                "상품명",
                "수량",
                "수취인 전화번호",
                "수취인 주소",
                "수취인 우편번호",
                "배송 메세지"
            ]

            progress_bar.progress(90)

            # 6️⃣ pandas로 임시로 저장
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Sheet1")

            # 7️⃣ openpyxl로 열 폭, 정렬, 중복 색칠 지정
            output.seek(0)
            wb = load_workbook(filename=output)
            ws = wb.active

            # 열 폭 리스트
            column_widths = [10, 50, 5, 15, 40, 20, 30]
            for idx, width in enumerate(column_widths):
                col_letter = ws.cell(row=1, column=idx+1).column_letter
                ws.column_dimensions[col_letter].width = width

            # C열(수량)을 가운데 정렬
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=3, max_col=3):
                for cell in row:
                    cell.alignment = Alignment(horizontal="center")

            # 🎨 중복 수취인명만 같은 색으로 표시
            receivers = [ws.cell(row=row_idx, column=1).value for row_idx in range(2, ws.max_row + 1)]
            receiver_counts = Counter(receivers)

            # 조금 더 진하게 보이는 색상 리스트
            fill_colors = [
                "B0C4DE",  # LightSteelBlue
                "ADD8E6",  # LightBlue
                "87CEEB",  # SkyBlue
                "D3D3D3",  # LightGray
                "C0C0C0",  # Silver
            ]

            color_map = {}
            color_idx = 0

            for row_idx in range(2, ws.max_row + 1):
                receiver = ws.cell(row=row_idx, column=1).value
                if receiver_counts[receiver] > 1:
                    if receiver not in color_map:
                        color_map[receiver] = fill_colors[color_idx % len(fill_colors)]
                        color_idx += 1
                    fill = PatternFill(start_color=color_map[receiver], end_color=color_map[receiver], fill_type="solid")
                    for col_idx in range(1, 8):
                        ws.cell(row=row_idx, column=col_idx).fill = fill

            # 최종 저장
            final_output = io.BytesIO()
            wb.save(final_output)
            final_output.seek(0)

            progress_bar.progress(100)

            # 다운로드 버튼
            st.success("🎉 처리 완료! 아래 버튼으로 다운로드하세요.")
            st.download_button(
                label="📥 가공된 엑셀 다운로드",
                data=final_output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"⚠️ 처리 중 오류 발생: {e}")
