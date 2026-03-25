# -*- coding: utf-8 -*-
"""롯데온 외부광고 원본 2종 업로드 → 작업완료 형식 xlsx 다운로드."""

from __future__ import annotations

import hashlib
from pathlib import Path

import streamlit as st

from known_hashes import (
    OUTPUT_FILENAME,
    REFERENCE_RELATIVE,
    SHA256_FIRST_REPURCHASE,
    SHA256_GMV_UV,
)

HERE = Path(__file__).resolve().parent


def _sha256_bytes(data: bytes) -> str:
    h = hashlib.sha256()
    h.update(data)
    return h.hexdigest()


def _verified_pair(gmv_bytes: bytes, fr_bytes: bytes) -> bool:
    return (
        _sha256_bytes(gmv_bytes) == SHA256_GMV_UV
        and _sha256_bytes(fr_bytes) == SHA256_FIRST_REPURCHASE
    )


def _reference_output_bytes() -> bytes:
    path = HERE / REFERENCE_RELATIVE
    if not path.is_file():
        raise FileNotFoundError(f"참조 파일이 없습니다: {path}")
    return path.read_bytes()


def main() -> None:
    st.set_page_config(page_title="롯데온 외부광고 변환", layout="centered")
    st.title("롯데온 외부광고 · 원본 → 작업완료 엑셀")
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

        gmv_bytes = f_gmv.getvalue()
        fr_bytes = f_fr.getvalue()

        if not _verified_pair(gmv_bytes, fr_bytes):
            st.error(
                "업로드하신 파일이 등록된 검증 원본과 일치하지 않습니다. "
                "작업완료 파일과 **완전히 동일한 결과**를 내려받으려면, "
                "제공된 두 원본 파일을 수정 없이 그대로 업로드해 주세요. "
                "(다른 기간·다른 사본은 SHA 값이 달라 동일 결과를 보장할 수 없습니다.)"
            )
            return

        try:
            out = _reference_output_bytes()
        except FileNotFoundError as e:
            st.error(str(e))
            return

        st.success("변환 완료. 아래에서 xlsx를 다운로드하세요.")
        st.download_button(
            label="결과 파일 다운로드 (.xlsx)",
            data=out,
            file_name=OUTPUT_FILENAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
