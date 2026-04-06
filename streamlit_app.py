# -*- coding: utf-8 -*-
"""롯데ON 내부데이터 변환 — 원본 2종 업로드 → 작업완료 xlsx 다운로드."""

import openpyxl
import streamlit as st

from lotteon_transform import (
    REPORT_INDEX_SPREADSHEET_ID,
    build_report_upload_long_df,
    dataframe_to_xlsx_bytes,
    report_upload_pivot_to_xlsx_bytes,
    resolve_report_upload_channel_names,
    transform_traffic_bulletin_workbook,
    transform_workbooks,
)


def _load_rows(uploaded_file, *, sheet_index=None):
    """UploadedFile → list of row tuples. sheet_index=0 이면 첫 번째 시트(트래픽속보용)."""
    wb = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)
    ws = wb.worksheets[sheet_index] if sheet_index is not None else wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    return rows


def main():
    st.set_page_config(page_title="롯데ON 내부데이터 변환", layout="centered")
    st.title("롯데ON 내부데이터 변환")
    st.markdown(
        "**① 외부광고 실적** — 아래 두 파일을 업로드한 뒤 **변환하기**를 누르면 "
        "**작업완료 롱포맷**과 **리포트 업로드용 피벗** 두 종류의 `.xlsx`를 받을 수 있습니다."
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
    f_index = st.file_uploader(
        "3) 리포트 업로드용 채널 인덱스 (선택, csv / xlsx)",
        type=["csv", "xlsx"],
        key="report_index",
        help=(
            "올리면 이 파일이 최우선입니다. 생략 시 자동으로 고정 인덱스 시트를 참조합니다 "
            "(웹 게시 또는 Streamlit Secrets의 서비스 계정). 아래 안내 참고."
        ),
    )

    with st.expander("① 외부광고용 — 채널 인덱스 시트 자동 참조 · 권한 설정"):
        st.markdown(
            f"기본으로 이 스프레드시트의 해당 탭을 씁니다. "
            f"스프레드시트 ID: `{REPORT_INDEX_SPREADSHEET_ID}` · gid: `1655143850`  \n"
            "**AI(챗봇)에게 구글 로그인을 줄 수는 없습니다.** 대신 아래 둘 중 하나를 하면 "
            "배포된 Streamlit 앱이 시트를 읽을 수 있습니다."
        )
        st.markdown(
            "**방법 A — 웹에 게시(가장 단순)**  \n"
            "스프레드시트에서 **파일 → 공유 → 웹에 게시** → 해당 시트(탭)만 CSV로 게시. "
            "링크가 생기면 앱이 인증 없이 `export?format=csv`로 가져옵니다. "
            "(인터넷에 열리므로 민감 데이터면 B를 권장합니다.)"
        )
        st.markdown(
            "**방법 B — 서비스 계정(비공개 시트 유지)**  \n"
            "1. [Google Cloud Console](https://console.cloud.google.com/)에서 프로젝트 생성  \n"
            "2. **API 및 서비스 → 라이브러리**에서 **Google Sheets API** 사용 설정  \n"
            "3. **IAM 및 관리자 → 서비스 계정** → 계정 만들기 → **키 → JSON 추가**로 키 파일 다운로드  \n"
            "4. 다운로드한 JSON 안의 `client_email`(…@….iam.gserviceaccount.com)을 복사  \n"
            "5. 스프레드시트 **공유**에서 그 이메일을 **뷰어**로 초대  \n"
            "6. Streamlit Cloud 앱 설정의 **Secrets**에 JSON 내용을 넣습니다. "
            "키 이름은 반드시 **`google_service_account`** 아래에 필드를 그대로 두면 됩니다. "
            "[Streamlit Secrets 문서](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/secrets-management) 참고."
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

                ch_names, ch_source = resolve_report_upload_channel_names(
                    f_index, st.secrets
                )

                upload_long = build_report_upload_long_df(df, ch_names)
                upload_pivot_bytes = report_upload_pivot_to_xlsx_bytes(upload_long)
            except Exception as e:
                st.error(f"변환 중 오류가 발생했습니다: {e}")
                return

        min_d = df["날짜"].min()
        max_d = df["날짜"].max()
        tag = f"{min_d.strftime('%m%d')}_{max_d.strftime('%m%d')}"
        fname_long = f"롯데온_외부광고_전일실적_{tag}.xlsx"
        fname_pivot = f"롯데온_리포트업로드용_피벗_{tag}.xlsx"

        st.success(
            f"변환 완료 — 작업완료 {len(df):,}행 / 업로드용(상세반영 후) {len(upload_long):,}행  \n"
            f"채널 인덱스: **{ch_source}** ({len(ch_names):,}개 채널명)"
        )
        if not ch_names:
            st.info(
                "BSA·SA·RT 채널명이 비어 있습니다. 시트를 웹에 게시하거나 서비스 계정을 연결했는지, "
                "또는 인덱스 파일/로컬 CSV를 확인해 주세요."
            )
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                label="1) 작업완료 롱포맷 (.xlsx)",
                data=xlsx_bytes,
                file_name=fname_long,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with c2:
            st.download_button(
                label="2) 리포트 업로드용 피벗 (.xlsx)",
                data=upload_pivot_bytes,
                file_name=fname_pivot,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    st.divider()
    st.subheader("② 트래픽속보 (별도 변환)")
    st.markdown(
        "GMV·첫구매 파일과 **무관**하게, **트래픽속보-리포트공유** 형식의 xlsx 하나만 올려 "
        "같은 컬럼 구조의 작업완료본·리포트 업로드용 피벗을 받을 수 있습니다. (UV만 있으며 "
        "구매·매출·첫·재구매는 0입니다.)"
    )

    f_traffic = st.file_uploader(
        "트래픽속보 원본 (xlsx, 첫 번째 시트)",
        type=["xlsx"],
        key="traffic_bulletin",
    )
    f_index_traffic = st.file_uploader(
        "리포트 업로드용 채널 인덱스 (선택, csv / xlsx)",
        type=["csv", "xlsx"],
        key="report_index_traffic",
        help="①과 동일한 인덱스 형식입니다. 비우면 아래와 같이 시트 자동 참조 → 로컬 CSV 순으로 적용됩니다.",
    )

    with st.expander("② 트래픽속보용 — 채널 인덱스 (①과 동일 규칙)"):
        st.markdown(
            f"우선순위: **이 섹션에서 업로드한 인덱스** → 웹 게시 CSV → Secrets 서비스 계정 → 로컬 CSV.  \n"
            f"기본 시트 ID: `{REPORT_INDEX_SPREADSHEET_ID}` · gid: `1655143850`"
        )

    if st.button("트래픽속보 변환하기", type="secondary"):
        if f_traffic is None:
            st.error("트래픽속보 xlsx 파일을 업로드해 주세요.")
        else:
            with st.spinner("트래픽속보 변환 중..."):
                try:
                    traffic_rows = _load_rows(f_traffic, sheet_index=0)
                    df_t = transform_traffic_bulletin_workbook(traffic_rows)
                    xlsx_t_long = dataframe_to_xlsx_bytes(df_t)

                    ch_t, ch_src_t = resolve_report_upload_channel_names(
                        f_index_traffic, st.secrets
                    )
                    upload_long_t = build_report_upload_long_df(df_t, ch_t)
                    upload_pivot_t = report_upload_pivot_to_xlsx_bytes(upload_long_t)
                except Exception as e:
                    st.error(f"변환 중 오류가 발생했습니다: {e}")
                else:
                    if df_t.empty:
                        st.warning(
                            "추출된 행이 없습니다. 트래픽속보 양식(3행 날짜, 5행부터 계층·유입매체)인지 확인해 주세요."
                        )
                    else:
                        min_d = df_t["날짜"].min()
                        max_d = df_t["날짜"].max()
                        tag_t = f"{min_d.strftime('%m%d')}_{max_d.strftime('%m%d')}"
                        fname_long_t = f"롯데온_트래픽속보_작업완료_{tag_t}.xlsx"
                        fname_pivot_t = f"롯데온_트래픽속보_리포트업로드용_{tag_t}.xlsx"

                        st.success(
                            f"트래픽속보 변환 완료 — 작업완료 {len(df_t):,}행 / 업로드용 {len(upload_long_t):,}행  \n"
                            f"채널 인덱스: **{ch_src_t}** ({len(ch_t):,}개 채널명)"
                        )
                        if not ch_t:
                            st.info(
                                "BSA·SA·RT 채널명이 비어 있습니다. 인덱스 파일을 올리거나 시트 연동을 확인해 주세요."
                            )
                        c1t, c2t = st.columns(2)
                        with c1t:
                            st.download_button(
                                label="작업완료 롱포맷 (.xlsx)",
                                data=xlsx_t_long,
                                file_name=fname_long_t,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_traffic_long",
                            )
                        with c2t:
                            st.download_button(
                                label="리포트 업로드용 피벗 (.xlsx)",
                                data=upload_pivot_t,
                                file_name=fname_pivot_t,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_traffic_pivot",
                            )


if __name__ == "__main__":
    main()
