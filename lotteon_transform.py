# -*- coding: utf-8 -*-
"""롯데온 외부광고 원본 2종 → 작업완료 롱포맷 변환."""

import io
from datetime import datetime
from pathlib import Path
from typing import Any, List, Mapping, Optional, Sequence, Tuple

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
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# 리포트 업로드용: BSA/SA/RT 채널 → 채널상세 비움 + 피벗(xlsx)
# ---------------------------------------------------------------------------

_BSA_SA_RT = frozenset({"BSA", "SA", "RT"})

# 리포트 업로드용 채널 인덱스 (구글 시트)
REPORT_INDEX_SPREADSHEET_ID = "18Gzpi_yeYQXbjqChlhm9EHT7z0Gi-65D0NCX7iC3SJ4"
REPORT_INDEX_GID = 1655143850


def load_default_bsa_sa_rt_channels() -> frozenset[str]:
    """저장소의 CSV(채널명 열)에서 기본 채널 목록을 읽는다."""
    path = Path(__file__).resolve().parent / "report_upload_bsa_sa_rt_channels.csv"
    if not path.exists():
        return frozenset()
    df = pd.read_csv(path, encoding="utf-8-sig", dtype=str)
    if df.empty:
        return frozenset()
    col = "채널명" if "채널명" in df.columns else df.columns[0]
    names = df[col].dropna().astype(str).str.strip()
    return frozenset(n for n in names if n and not n.startswith("#"))


def channel_names_from_index_dataframe(df: pd.DataFrame) -> frozenset[str]:
    """
    인덱스 시트(보낸 xlsx/csv)에서 I열(미디어)이 정확히 BSA·SA·RT인 행의 채널명만 모은다.
    열 이름이 있으면 '미디어', '채널명' 우선. 없으면 미디어=9번째 열(I), 채널명=두 번째 열(B).
    """
    if df.empty or len(df.columns) == 0:
        return frozenset()
    headers = {str(c).strip(): c for c in df.columns}
    media_col = headers.get("미디어")
    if media_col is None and len(df.columns) > 8:
        media_col = df.columns[8]
    if media_col is None:
        return frozenset()

    name_col = headers.get("채널명")
    if name_col is None and len(df.columns) > 1:
        name_col = df.columns[1]
    elif name_col is None:
        name_col = df.columns[0]

    media = df[media_col].astype(str).str.strip()
    mask = media.isin(_BSA_SA_RT)
    names = df.loc[mask, name_col].dropna().astype(str).str.strip()
    return frozenset(n for n in names if n)


def try_channel_names_from_published_google_sheet() -> Optional[frozenset[str]]:
    """
    시트를 「웹에 게시」한 경우(링크만으로 CSV 다운로드 가능) 인증 없이 읽는다.
    비공개 시트는 HTML 로그인 페이지가 오므로 None을 반환한다.
    """
    import urllib.error
    import urllib.request

    url = (
        f"https://docs.google.com/spreadsheets/d/{REPORT_INDEX_SPREADSHEET_ID}/export"
        f"?format=csv&gid={REPORT_INDEX_GID}"
    )
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "lotteon-report/1.0"})
        with urllib.request.urlopen(req, timeout=20) as resp:
            raw = resp.read()
    except (urllib.error.URLError, OSError, TimeoutError):
        return None
    if b"Sign in" in raw[:4000] or b"<html" in raw[:500].lower():
        return None
    for enc in ("utf-8-sig", "utf-8", "cp949"):
        try:
            df = pd.read_csv(io.BytesIO(raw), encoding=enc, dtype=str)
            names = channel_names_from_index_dataframe(df)
            return names if names else None
        except (UnicodeDecodeError, pd.errors.ParserError, ValueError):
            continue
    return None


def channel_names_from_google_sheet(service_account_info: Mapping[str, Any]) -> frozenset[str]:
    """Google Sheets API(서비스 계정)로 인덱스 탭을 읽는다. 시트는 해당 계정 이메일에 공유되어 있어야 한다."""
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(dict(service_account_info), scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(REPORT_INDEX_SPREADSHEET_ID)
    ws = sh.get_worksheet_by_id(int(REPORT_INDEX_GID))
    rows = ws.get_all_values()
    if not rows:
        return frozenset()
    header = [str(c).strip() for c in rows[0]]
    df = pd.DataFrame(rows[1:], columns=header, dtype=str)
    return channel_names_from_index_dataframe(df)


def resolve_report_upload_channel_names(
    streamlit_secrets: Optional[Mapping[str, Any]] = None,
) -> Tuple[frozenset[str], str]:
    """
    채널명 집합과, 어떤 경로로 읽었는지 짧은 설명을 반환한다.
    우선순위: 웹 게시 CSV → Secrets 서비스 계정 → 로컬 CSV
    """
    pub = try_channel_names_from_published_google_sheet()
    if pub:
        return pub, "구글 시트(웹 게시 CSV)"

    if streamlit_secrets is not None and "google_service_account" in streamlit_secrets:
        block = streamlit_secrets["google_service_account"]
        info = dict(block) if hasattr(block, "keys") else block
        return channel_names_from_google_sheet(info), "구글 시트(API·서비스 계정)"

    names = load_default_bsa_sa_rt_channels()
    return names, "로컬 report_upload_bsa_sa_rt_channels.csv"


def build_report_upload_long_df(df: pd.DataFrame, channel_names: frozenset[str]) -> pd.DataFrame:
    """
    작업완료 롱포맷과 동일 컬럼이나, channel_names에 해당하는 채널은 채널상세를 비운 뒤 동일 키로 합산한다.
    """
    out = df.copy()
    out["날짜"] = pd.to_datetime(out["날짜"])
    out["채널상세"] = out["채널상세"].fillna("")
    if channel_names:
        out.loc[out["채널명"].isin(channel_names), "채널상세"] = ""
    key = ["날짜", "채널명", "채널상세", "유입매체구분"]
    return (
        out.groupby(key, sort=False, as_index=False)
        .agg(
            UV=("UV", "sum"),
            구매자수=("구매자수", "sum"),
            판매매출=("판매매출", "sum"),
            첫구매=("첫구매", "sum"),
            재구매=("재구매", "sum"),
        )
    )


def report_upload_pivot_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    """
    지표별 시트로 날짜 피벗(열=날짜, 행=채널명·채널상세·유입매체구분).
    """
    work = df.copy()
    work["날짜"] = pd.to_datetime(work["날짜"])
    idx_cols = ["채널명", "채널상세", "유입매체구분"]
    metrics = ["UV", "구매자수", "판매매출", "첫구매", "재구매"]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for m in metrics:
            pt = work.pivot_table(
                index=idx_cols,
                columns="날짜",
                values=m,
                aggfunc="sum",
                fill_value=0,
            )
            pt = pt.sort_index(axis=1)
            pt.columns = [
                c.strftime("%Y-%m-%d") if hasattr(c, "strftime") else str(c)[:10]
                for c in pt.columns
            ]
            pt.reset_index().to_excel(writer, index=False, sheet_name=m[:31])
    return buf.getvalue()
