# --- 8. 엑셀 다운로드: 원본 양식 유지 ---
mapping = {}

for bx in bins:
    for item in (
        [i for r in bx['rows'] for i in r['items']]
        + bx.get('stacked_items', [])
    ):
        mapping[item['row_idx']] = bx['c_label']


st.markdown("---")
st.subheader("💾 최종 배정 결과 다운로드")

if file.name.endswith('.xlsx'):
    # 원본 엑셀 파일 포인터를 처음으로 되돌림
    file.seek(0)

    # 원본 엑셀 양식 그대로 열기
    wb = load_workbook(file)

    # 실제 데이터가 있는 시트 선택
    # selected_sheet는 위에서 자동으로 찾은 시트명
    ws = wb[selected_sheet]

    # 마지막 사용 열 다음 열에 결과 입력
    target_col = ws.max_column + 1

    # 제목 입력
    ws.cell(row=4, column=target_col).value = "배정 컨테이너"

    # 가능하면 왼쪽 셀 스타일을 복사해서 제목 셀에도 적용
    if target_col > 1:
        from copy import copy

        title_source = ws.cell(row=4, column=target_col - 1)
        title_target = ws.cell(row=4, column=target_col)

        title_target.font = copy(title_source.font)
        title_target.fill = copy(title_source.fill)
        title_target.border = copy(title_source.border)
        title_target.alignment = copy(title_source.alignment)
        title_target.number_format = title_source.number_format

    # 배정 결과 입력
    # row_idx는 pandas 기준 0부터 시작하므로, 엑셀 행 번호는 +1
    for r_idx, label in mapping.items():
        excel_row = int(r_idx) + 1
        ws.cell(row=excel_row, column=target_col).value = label

        # 왼쪽 셀 스타일 복사
        if target_col > 1:
            from copy import copy

            source_cell = ws.cell(row=excel_row, column=target_col - 1)
            target_cell = ws.cell(row=excel_row, column=target_col)

            target_cell.font = copy(source_cell.font)
            target_cell.fill = copy(source_cell.fill)
            target_cell.border = copy(source_cell.border)
            target_cell.alignment = copy(source_cell.alignment)
            target_cell.number_format = source_cell.number_format

    # 열 너비 설정
    ws.column_dimensions[ws.cell(row=1, column=target_col).column_letter].width = 25

    output = io.BytesIO()
    wb.save(output)

    st.download_button(
        label="📥 원본 양식 그대로 컨테이너 번호 달아서 저장하기",
        data=output.getvalue(),
        file_name="CLP_배정결과_원본양식.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

else:
    st.warning("CSV 파일은 원본 엑셀 양식 보존이 불가능합니다. 양식 보존은 xlsx 파일에서만 지원됩니다.")
