# -*- coding: utf-8 -*-
"""롯데온 외부광고 원본 2종 → 작업완료 롱포맷 변환."""

from datetime import datetime
from typing import Any, List, Optional, Sequence, Tuple

import pandas as pd

import re

COLS_INTERNAL = ["날짜", "채널명", "채널상세", "유입매체구분", "UV", "구매자수", "판매매출", "첫구매여부(전체)"]
COLS = ["날짜", "채널명", "채널상세", "유입매체구분", "UV", "구매자수", "판매매출", "첫구매", "재구매"]

EC_VALS = frozenset(("패션", "LIFE", "뷰티", "B2B", "미분류"))

_SURROGATE_RE = re.compile(r"_x([0-9A-Fa-f]{4})_")


def _decode_surrogates(s: str) -> str:
    """openpyxl이 이모지를 _xD83D__xDE00_ 식으로 읽는 것을 실제 유니코드로 복원."""
    parts = _SURROGATE_RE.split(s)
    result = []
    i = 0
    while i < len(parts):
        if i % 2 == 0:
            result.append(parts[i])
        else:
            code = int(parts[i], 16)
            if 0xD800 <= code <= 0xDBFF and i + 2 < len(parts) and parts[i + 1] == "":
                low = int(parts[i + 2], 16)
                if 0xDC00 <= low <= 0xDFFF:
                    char = chr(0x10000 + (code - 0xD800) * 0x400 + (low - 0xDC00))
                    result.append(char)
                    i += 2
                else:
                    result.append(chr(code))
            else:
                try:
                    result.append(chr(code))
                except (ValueError, OverflowError):
                    result.append(parts[i])
        i += 1
    return "".join(result)


def _date_starts(row3: Sequence[Any]) -> List[int]:
    return [i for i, x in enumerate(row3) if isinstance(x, datetime)]


# ---------------------------------------------------------------------------
# File 1: GMV UV 구매자수
# ---------------------------------------------------------------------------

def _parse_gmv_leaves(data_rows: List[Tuple[Any, ...]]):
    """col7(유입매체구분)이 채워진 행 = leaf. (name, detail, media, row) 목록 반환."""
    leaves = []
    name_ff = None
    detail_ff = None

    for row in data_rows:
        r = list(row)

        if r[3] is not None and str(r[3]).strip():
            name_ff = _decode_surrogates(str(r[3]))
            detail_ff = None
        if r[4] is not None and str(r[4]).strip():
            detail_ff = _decode_surrogates(str(r[4]))

        media = r[6]
        if media is None or not str(media).strip():
            continue
        media = str(media).strip()
        if name_ff is None or detail_ff is None:
            continue

        leaves.append((name_ff, detail_ff, media, r))

    return leaves


def parse_gmv_uv_sheet(rows: List[Tuple[Any, ...]]) -> pd.DataFrame:
    if len(rows) < 5:
        return pd.DataFrame(columns=COLS_INTERNAL)

    row3 = rows[2]
    ds_list = _date_starts(row3)
    if not ds_list:
        raise ValueError("GMV 파일에서 날짜 열을 찾지 못했습니다.")

    leaves = _parse_gmv_leaves(rows[4:])
    out = []

    for ds in ds_list:
        dt = row3[ds]
        if not isinstance(dt, datetime):
            continue
        d = dt.date()

        for name, detail, media, r in leaves:
            if ds + 2 >= len(r):
                continue
            out.append({
                "날짜": d,
                "채널명": name,
                "채널상세": detail,
                "유입매체구분": media,
                "UV": r[ds],
                "구매자수": r[ds + 1],
                "판매매출": r[ds + 2],
                "첫구매여부(전체)": "없음",
            })

    return pd.DataFrame(out)


# ---------------------------------------------------------------------------
# File 2: 첫구매 재구매
# ---------------------------------------------------------------------------

def _parse_fr_leaves(data_rows: List[Tuple[Any, ...]]):
    """
    계층 구조에서 leaf 행만 추출.
    level 1 = col7(유입매체구분), level 2 = col8(유입매체), level 3 = col9(EC영업실).
    leaf = 다음 후보 행이 자신보다 깊지 않은 행.
    """
    candidates = []
    name_ff = None
    detail_ff = None
    flag_ff = None
    media_ff = None

    for row in data_rows:
        r = list(row)

        if r[3] is not None and str(r[3]).strip():
            name_ff = _decode_surrogates(str(r[3]))
            detail_ff = None
            flag_ff = None
            media_ff = None
        if r[4] is not None and str(r[4]).strip():
            detail_ff = _decode_surrogates(str(r[4]))
        if r[5] in ("첫구매", "재구매"):
            flag_ff = str(r[5])
            media_ff = None

        if flag_ff not in ("첫구매", "재구매"):
            continue
        if name_ff is None or detail_ff is None:
            continue

        level = 0
        if r[6] is not None and str(r[6]).strip():
            level = 1
            media_ff = str(r[6]).strip()
        elif r[7] is not None and str(r[7]).strip():
            level = 2
        elif r[8] is not None and str(r[8]).strip() in EC_VALS:
            level = 3
        else:
            continue

        if media_ff is None:
            continue

        candidates.append((level, name_ff, detail_ff, flag_ff, media_ff, r))

    leaves = []
    for i, cand in enumerate(candidates):
        is_leaf = True
        if i + 1 < len(candidates):
            if candidates[i + 1][0] > cand[0]:
                is_leaf = False
        if not is_leaf:
            continue
        r = cand[5]
        total_sales = r[9] if len(r) > 9 else None
        total_buyers = r[10] if len(r) > 10 else None
        if (total_sales is None or total_sales == 0) and (total_buyers is None or total_buyers == 0):
            continue
        leaves.append(cand[1:])  # (name, detail, flag, media, row)

    return leaves


def parse_first_repurchase_sheet(rows: List[Tuple[Any, ...]]) -> pd.DataFrame:
    if len(rows) < 5:
        return pd.DataFrame(columns=COLS_INTERNAL)

    row3 = rows[2]
    ds_list = _date_starts(row3)
    if not ds_list:
        raise ValueError("첫구매 파일에서 날짜 열을 찾지 못했습니다.")

    leaves = _parse_fr_leaves(rows[4:])
    out = []

    for ds in ds_list:
        dt = row3[ds]
        if not isinstance(dt, datetime):
            continue
        d = dt.date()

        for name, detail, flag, media, r in leaves:
            if ds + 1 >= len(r):
                continue
            out.append({
                "날짜": d,
                "채널명": name,
                "채널상세": detail,
                "유입매체구분": media,
                "UV": 0,
                "구매자수": r[ds + 1],
                "판매매출": r[ds],
                "첫구매여부(전체)": flag,
            })

    return pd.DataFrame(out)


# ---------------------------------------------------------------------------
# 통합
# ---------------------------------------------------------------------------

def _postprocess(df: pd.DataFrame) -> pd.DataFrame:
    """후처리: 기타→PC, 첫구매여부 피벗, 합계 0 제거, 키 기준 합산."""
    df["유입매체구분"] = df["유입매체구분"].replace("기타", "PC")
    df["날짜"] = pd.to_datetime(df["날짜"])

    for c in ("UV", "구매자수", "판매매출"):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)

    key = ["날짜", "채널명", "채널상세", "유입매체구분"]

    base = df[df["첫구매여부(전체)"] == "없음"].groupby(key, sort=False, as_index=False).agg(
        UV=("UV", "sum"), 구매자수=("구매자수", "sum"), 판매매출=("판매매출", "sum"),
    )

    first = df[df["첫구매여부(전체)"] == "첫구매"].groupby(key, sort=False, as_index=False).agg(
        첫구매=("구매자수", "sum"),
    )

    repurch = df[df["첫구매여부(전체)"] == "재구매"].groupby(key, sort=False, as_index=False).agg(
        재구매=("구매자수", "sum"),
    )

    result = base.merge(first, on=key, how="left").merge(repurch, on=key, how="left")
    result["첫구매"] = result["첫구매"].fillna(0).astype(int)
    result["재구매"] = result["재구매"].fillna(0).astype(int)

    all_zero = (
        (result["UV"] == 0) & (result["구매자수"] == 0) & (result["판매매출"] == 0)
        & (result["첫구매"] == 0) & (result["재구매"] == 0)
    )
    result = result[~all_zero].reset_index(drop=True)

    return result[COLS]


def transform_workbooks(
    gmv_rows: List[Tuple[Any, ...]], fr_rows: List[Tuple[Any, ...]]
) -> pd.DataFrame:
    df1 = parse_gmv_uv_sheet(gmv_rows)
    df2 = parse_first_repurchase_sheet(fr_rows)
    frames = [df for df in (df1, df2) if not df.empty]
    if not frames:
        return pd.DataFrame(columns=COLS)
    merged = pd.concat(frames, ignore_index=True)
    return _postprocess(merged)


def dataframe_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    import io
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buf.getvalue()
