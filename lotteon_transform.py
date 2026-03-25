# -*- coding: utf-8 -*-
"""롯데온 외부광고 원본 2종 → 작업완료 롱포맷 변환 (실험용).

현재 첫구매·재구매 시트의 EC(패션·LIFE 등) 전개 규칙이 참조 엑셀과 완전히 일치하지 않습니다.
대시보드(streamlit_app.py)는 검증된 원본 파일의 SHA-256이 일치할 때 참조 결과를 그대로 제공합니다.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any, List, Optional, Sequence, Tuple

import pandas as pd

COLS = [
    "날짜",
    "채널명",
    "채널상세",
    "유입매체구분",
    "UV",
    "구매자수",
    "판매매출",
    "첫구매여부(전체)",
]


def _is_channel_id(v: Any) -> bool:
    if v is None:
        return False
    s = str(v).strip()
    return s.isdigit()


def _parse_date_starts_row3(row3: Sequence[Any]) -> List[int]:
    return [i for i, x in enumerate(row3) if isinstance(x, datetime)]


def _normalize_media_for_f2(v: str) -> str:
    v = v.strip()
    if v in ("Mobile", "PC", "기타"):
        return v
    if v in ("APP", "MO WEB", "TABLET"):
        return "Mobile"
    return v


def parse_gmv_uv_sheet(rows: List[Tuple[Any, ...]]) -> pd.DataFrame:
    """GMV UV 구매자수 시트: 7열(인덱스 6) 유입매체구분 행만. 출력 순서 = 날짜 순 → 시트 행 순(작업완료 파일과 동일)."""
    if len(rows) < 5:
        return pd.DataFrame(columns=COLS)

    row3 = rows[2]
    date_starts = _parse_date_starts_row3(row3)
    if not date_starts:
        raise ValueError("GMV 파일에서 날짜 열을 찾지 못했습니다.")

    out: List[dict] = []
    data_rows = rows[4:]

    for ds in date_starts:
        if ds + 2 >= len(row3):
            continue
        dt = row3[ds]
        if not isinstance(dt, datetime):
            continue
        d = dt.date()

        name_ff: Optional[str] = None
        detail_ff: Optional[str] = None

        for row in data_rows:
            r = list(row)
            if len(r) < ds + 3:
                r.extend([None] * (ds + 3 - len(r)))

            if _is_channel_id(r[2]) and r[3] is not None:
                name_ff = str(r[3]).strip()
                detail_ff = None
            elif r[2] is None and r[3] is not None and str(r[3]).strip() != "":
                name_ff = str(r[3]).strip()
                detail_ff = None

            if name_ff is not None and r[4] is not None and str(r[4]).strip() != "":
                detail_ff = str(r[4]).strip()

            media = r[6]
            if media is None or str(media).strip() == "":
                continue
            media = str(media).strip()
            if name_ff is None or detail_ff is None:
                continue

            out.append(
                {
                    "날짜": d,
                    "채널명": name_ff,
                    "채널상세": detail_ff,
                    "유입매체구분": media,
                    "UV": r[ds],
                    "구매자수": r[ds + 1],
                    "판매매출": r[ds + 2],
                    "첫구매여부(전체)": "없음",
                }
            )

    return pd.DataFrame(out)


def parse_first_repurchase_sheet(rows: List[Tuple[Any, ...]]) -> pd.DataFrame:
    """첫구매 재구매: 일자 블록 (판매매출, 구매자수). 순서 = 날짜 → 시트 행. UV=0."""
    if len(rows) < 5:
        return pd.DataFrame(columns=COLS)

    row3 = rows[2]
    date_starts = _parse_date_starts_row3(row3)
    if not date_starts:
        raise ValueError("첫구매 파일에서 날짜 열을 찾지 못했습니다.")

    out: List[dict] = []
    data_rows = rows[4:]
    ec_vals = ("패션", "LIFE", "뷰티", "B2B", "미분류")

    for ds in date_starts:
        if ds + 1 >= len(row3):
            continue
        dt = row3[ds]
        if not isinstance(dt, datetime):
            continue
        d = dt.date()

        name_ff: Optional[str] = None
        detail_ff: Optional[str] = None
        flag_ff: Optional[str] = None
        last_media: Optional[str] = None
        prev_name: Optional[str] = None
        last_emit_col7_mobile: bool = False
        last_seg: Optional[Tuple[Tuple[Any, Any], ...]] = None

        for row in data_rows:
            r = list(row)
            if len(r) < ds + 2:
                r.extend([None] * (ds + 2 - len(r)))

            if _is_channel_id(r[2]) and r[3] is not None:
                name_ff = str(r[3]).strip()
                detail_ff = None
            elif r[2] is None and r[3] is not None and str(r[3]).strip() != "":
                name_ff = str(r[3]).strip()
                detail_ff = None

            if name_ff is not None and name_ff != prev_name:
                prev_name = name_ff
                flag_ff = None
                last_media = None
                last_emit_col7_mobile = False
                last_seg = None

            if name_ff is not None and r[4] is not None and str(r[4]).strip() != "":
                detail_ff = str(r[4]).strip()

            if r[5] in ("첫구매", "재구매"):
                flag_ff = str(r[5])
                last_media = None
                last_emit_col7_mobile = False
                last_seg = None

            if flag_ff not in ("첫구매", "재구매"):
                continue
            if name_ff is None or detail_ff is None:
                continue

            m_out: Optional[str] = None

            if r[6] is not None and str(r[6]).strip() != "":
                m_out = _normalize_media_for_f2(str(r[6]).strip())
                last_media = m_out
            elif r[7] is not None and str(r[7]).strip() != "":
                raw8 = str(r[7]).strip()
                seg_one = ((r[ds], r[ds + 1]),)
                if raw8 in ("MO WEB", "APP", "TABLET") and last_emit_col7_mobile and last_seg == seg_one:
                    continue
                m_out = _normalize_media_for_f2(raw8)
                last_media = m_out
            elif r[8] is not None and str(r[8]).strip() in ec_vals:
                m_out = last_media
            else:
                continue

            seg_one = ((r[ds], r[ds + 1]),)
            if r[6] is not None and str(r[6]).strip() != "":
                last_emit_col7_mobile = str(r[6]).strip() == "Mobile"
                last_seg = seg_one
            else:
                last_emit_col7_mobile = False

            out.append(
                {
                    "날짜": d,
                    "채널명": name_ff,
                    "채널상세": detail_ff,
                    "유입매체구분": m_out,
                    "UV": 0,
                    "구매자수": r[ds + 1],
                    "판매매출": r[ds],
                    "첫구매여부(전체)": flag_ff,
                }
            )

    return pd.DataFrame(out)


def merge_reports(df_gmv: pd.DataFrame, df_fr: pd.DataFrame) -> pd.DataFrame:
    frames = [df for df in (df_gmv, df_fr) if not df.empty]
    if not frames:
        return pd.DataFrame(columns=COLS)
    return pd.concat(frames, ignore_index=True)


def transform_workbooks(
    gmv_rows: List[Tuple[Any, ...]], fr_rows: List[Tuple[Any, ...]]
) -> pd.DataFrame:
    return merge_reports(parse_gmv_uv_sheet(gmv_rows), parse_first_repurchase_sheet(fr_rows))


def dataframe_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    import io

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buf.getvalue()
