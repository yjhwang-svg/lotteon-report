# -*- coding: utf-8 -*-
"""롯데ON 내부데이터 변환 — 원본 2종 업로드 → 작업완료 xlsx 다운로드."""

import openpyxl
import streamlit as st

from lotteon_transform import dataframe_to_xlsx_bytes, transform_workbooks


def _load_rows(uploaded_file):
    """UploadedFile → list of row tuples."""
    wb = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)
    rows = list(wb.active.iter_rows(values_only=True))
    wb.close()
    return rows


def main():
    st.set_page_config(page_title="롯데ON 내부데이터 변환", layout="centered")
    st.title("롯데ON 내부데이터 변환")
    st.markdown(
        "아래 두 파일을 업로드한 뒤 **변환하기**를 누르면 `.xlsx` 파일을 받을 수 있습니다."
    )

    f_gmv = st.file_uploader(
        "1) GMV · UV · 구매자수 원본 (xlsx)",
        type=["xlsx"],
        key="gmv",
    )
    f_fr = st.file_uploader(
        "2) 첫구매 · 재구매 원본 (xlsx)",
        type=["xlsx"],
        key="fr",
    )

    if st.button("변환하기", type="primary"):
        if f_gmv is None or f_fr is None:
            st.error("두 개의 xlsx 파일을 모두 업로드해 주세요.")
            return

        with st.spinner("변환 중..."):
            try:
                gmv_rows = _load_rows(f_gmv)
                fr_rows = _load_rows(f_fr)
                df = transform_workbooks(gmv_rows, fr_rows)
                xlsx_bytes = dataframe_to_xlsx_bytes(df)
            except Exception as e:
                st.error(f"변환 중 오류가 발생했습니다: {e}")
                return

        st.success(f"변환 완료 — 총 {len(df):,}행")
        st.download_button(
            label="결과 파일 다운로드 (.xlsx)",
            data=xlsx_bytes,
            file_name="롯데온_외부광고_변환결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
