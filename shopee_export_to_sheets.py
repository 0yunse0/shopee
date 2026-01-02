import os
import re
import json
import time
import glob
import shutil
import datetime as dt
from typing import Optional, List, Dict

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError


# =========================
# 유틸: 날짜 파싱 (MM-DD-YYYY)
# =========================
DATE_FMT = "%m-%d-%Y"


def parse_mmddyyyy(s: str) -> Optional[dt.date]:
    s = (s or "").strip()
    try:
        return dt.datetime.strptime(s, DATE_FMT).date()
    except Exception:
        return None


def today_str():
    return dt.date.today().strftime("%Y%m%d")


# =========================
# Google Sheets
# =========================
def get_gspread_client(service_account_path: str):
    creds = Credentials.from_service_account_file(
        service_account_path,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return gspread.authorize(creds)


def get_last_date_in_sheet(ws) -> Optional[dt.date]:
    """
    너가 말한 구조:
    - 1행 비어있을 수 있음
    - 2행은 헤더(카테고리)
    - 실 데이터는 3행부터
    - 날짜는 A열
    """
    colA = ws.col_values(1)  # A열
    if len(colA) < 3:
        return None

    # 3행부터 훑어서 마지막 유효 날짜 찾기
    last = None
    for v in colA[2:]:
        d = parse_mmddyyyy(v)
        if d:
            last = d
    return last


def append_rows(ws, rows: List[List]):
    if not rows:
        return
    ws.append_rows(rows, value_input_option="USER_ENTERED")


# =========================
# Playwright: 로그인/세션
# =========================
def ensure_session(session_file: str, login_market_base: str):
    """
    세션 파일이 없으면 headless=False로 로그인 한번 시키고(2FA 포함),
    storage_state 저장
    """
    if os.path.exists(session_file):
        print(f"[OK] session exists: {session_file}")
        return

    print("[NEED] session file not found. Starting manual login flow...")
    print(" - 브라우저 열리면 로그인 + (2FA 있으면) 직접 처리하고 기다려줘.")
    print(" - 로그인 후 Seller Center 메인 화면 떠있으면 자동 저장됨.")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        # 로그인 페이지는 마켓별로 다를 수 있어서 base로 일단 진입
        page.goto(login_market_base, wait_until="domcontentloaded")
        page.wait_for_timeout(3000)

        # 사용자가 직접 로그인/2FA 처리
        # 너가 로그인 끝냈다는 기준을 "seller center 페이지가 로딩된 상태"로 잡고,
        # 2분 정도 여유 줌. 필요하면 늘려.
        for _ in range(240):  # 240 * 0.5s = 120s
            url = page.url
            if "seller.shopee" in url and ("portal" in url or "home" in url or "datacenter" in url):
                break
            page.wait_for_timeout(500)

        # 그래도 저장은 시도
        context.storage_state(path=session_file)
        print(f"[OK] session saved: {session_file}")

        browser.close()


# =========================
# Playwright: Data Center Export
# =========================
def export_datacenter_csv(
    market_base: str,
    session_file: str,
    download_dir: str,
    account_key: str,
) -> str:
    """
    Data Center Overview 페이지에서:
    - 기간 Past 7 days
    - Export 클릭
    - CSV 저장 경로 리턴
    """
    os.makedirs(download_dir, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            storage_state=session_file,
            accept_downloads=True
        )
        page = context.new_page()

        target_url = f"{market_base}/datacenter/overview"
        print(f"[{account_key}] goto: {target_url}")
        page.goto(target_url, wait_until="domcontentloaded")
        page.wait_for_timeout(2000)

        # ====== Past 7 days 선택 ======
        # UI가 계정/언어/개편에 따라 다르니, 가장 흔한 패턴을 몇 개 시도함.
        # 실패하면 아래 selector를 너 화면에 맞게 1~2개만 바꿔주면 됨.
        def try_click(selectors: List[str], desc: str):
            for sel in selectors:
                try:
                    page.click(sel, timeout=3000)
                    print(f"[{account_key}] clicked: {desc} ({sel})")
                    return True
                except Exception:
                    continue
            return False

        # 날짜/기간 드롭다운 열기
        opened = try_click(
            [
                "text=Date",
                "text=Period",
                "button:has-text('Date')",
                "button:has-text('Period')",
                "[data-testid*='date']",
            ],
            "open period dropdown"
        )

        if not opened:
            print(f"[{account_key}] WARN: period dropdown open failed (maybe already visible).")

        # Past 7 days 선택
        picked = try_click(
            [
                "text=Past 7 days",
                "text=Last 7 days",
                "li:has-text('Past 7 days')",
                "li:has-text('Last 7 days')",
                "button:has-text('Past 7 days')",
                "button:has-text('Last 7 days')",
            ],
            "select past 7 days"
        )
        if not picked:
            print(f"[{account_key}] WARN: couldn't click 'Past 7 days' - check UI wording.")
        page.wait_for_timeout(1500)

        # ====== Export 버튼 클릭 & 다운로드 ======
        export_clicked = False
        export_selectors = [
            "text=Export",
            "button:has-text('Export')",
            "[data-testid*='export']",
        ]
        for sel in export_selectors:
            try:
                with page.expect_download(timeout=15000) as d:
                    page.click(sel, timeout=5000)
                    export_clicked = True
                download = d.value
                break
            except Exception:
                continue

        if not export_clicked:
            raise RuntimeError(f"[{account_key}] Export click failed. UI selector needs update.")

        # 파일 저장
        suggested = download.suggested_filename or f"{account_key}_export.csv"
        safe_name = re.sub(r"[^a-zA-Z0-9._-]+", "_", suggested)
        out_path = os.path.join(download_dir, f"{account_key}_{today_str()}_{safe_name}")
        download.save_as(out_path)

        print(f"[{account_key}] downloaded: {out_path}")

        context.close()
        browser.close()

        return out_path


# =========================
# CSV 파싱 (컬럼명 차이 대응)
# =========================
def find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = {c.strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in cols:
            return cols[key]
    # 부분매칭도 시도
    lower_cols = [c.lower() for c in df.columns]
    for cand in candidates:
        for i, c in enumerate(lower_cols):
            if cand.lower() in c:
                return df.columns[i]
    return None


def parse_export_csv(csv_path: str, order_rule: str) -> pd.DataFrame:
    """
    Data Center export 파일 구조가 너 캡처처럼
    날짜별 합계 + 일자 행들로 나오기도 하고,
    혹은 주문 단위로 나오기도 해서
    둘 다 대응:

    - 만약 이미 "Date, Sales, Orders ..." 형태로 '일자별'이 들어있으면 그대로 사용
    - 아니면 주문/행 단위면 Date로 groupby 해서 집계
    """
    df = pd.read_csv(csv_path)
    df.columns = [str(c).strip() for c in df.columns]

    # 1) 이미 일자별 테이블인지 판단: Date + Sales + Orders 같은 컬럼이 있고, Date가 MM-DD-YYYY면
    date_col = find_col(df, ["Date"])
    sales_col = find_col(df, ["Sales", "Sales (SGD)", "Sales (MYR)", "Sales (PHP)"])
    orders_col = find_col(df, ["Orders", "Order", "Order Count"])

    if date_col and sales_col and orders_col:
        # date 파싱 가능한 행만 남김
        tmp = df.copy()
        tmp["__date"] = tmp[date_col].astype(str).apply(parse_mmddyyyy)
        tmp = tmp[tmp["__date"].notna()].copy()
        # 필요한 컬럼들 최대한 맞춰서 반환
        # (E,F,M,N,O 같은 컬럼도 export에 있으면 함께 사용 가능)
        return tmp

    # 2) 주문 단위로 나온 경우: status 컬럼 찾아서 Paid/Confirmed 필터 후 groupby
    status_col = find_col(df, ["Order Status", "Status"])
    if status_col:
        if order_rule == "PAID":
            df = df[df[status_col].astype(str).str.contains("Paid", case=False, na=False)]
        else:
            df = df[df[status_col].astype(str).str.contains("Confirm", case=False, na=False)]

    date_col = date_col or find_col(df, ["Order Date", "Date"])
    if not date_col:
        raise RuntimeError(f"Cannot find date column in CSV: {csv_path}")

    df["__date"] = df[date_col].astype(str).apply(parse_mmddyyyy)
    df = df[df["__date"].notna()].copy()

    # 집계 대상 컬럼들
    sales_col = sales_col or find_col(df, ["Sales", "Amount", "Total"])
    clicks_col = find_col(df, ["Product Clicks", "Clicks"])
    visitors_col = find_col(df, ["Visitors", "Visitor"])
    cancelled_orders_col = find_col(df, ["Cancelled Orders", "Canceled Orders"])
    cancelled_sales_col = find_col(df, ["Cancelled Sales", "Canceled Sales"])

    agg = {}
    if sales_col: agg[sales_col] = "sum"
    agg["__orders"] = "count"
    if clicks_col: agg[clicks_col] = "sum"
    if visitors_col: agg[visitors_col] = "sum"
    if cancelled_orders_col: agg[cancelled_orders_col] = "sum"
    if cancelled_sales_col: agg[cancelled_sales_col] = "sum"

    df["__orders"] = 1
    g = df.groupby("__date").agg(agg).reset_index()

    # 표준 컬럼명으로 정리
    out = pd.DataFrame()
    out["Date"] = g["__date"].apply(lambda d: d.strftime(DATE_FMT))
    if sales_col: out["Sales"] = g[sales_col]
    out["Orders"] = g["__orders"]
    if clicks_col: out["Product Clicks"] = g[clicks_col]
    if visitors_col: out["Visitors"] = g[visitors_col]
    if cancelled_orders_col: out["Cancelled Orders"] = g[cancelled_orders_col]
    if cancelled_sales_col: out["Cancelled Sales"] = g[cancelled_sales_col]

    return out


# =========================
# 시트 구조에 맞게 row 만들기
# =========================
def build_sheet_rows(df: pd.DataFrame, ws) -> List[List]:
    """
    시트는 너 캡처 기준:
    A: Date
    B: Sales
    C: Orders
    D: Sales per Order (수식 or 우리가 계산)
    E: Product Clicks
    F: Visitors
    ...
    M,N,O ... 도 있음

    여기선 "시트 2행 헤더"를 읽어서 그 순서대로 값을 맞춰 넣음.
    - 없는 컬럼은 빈칸("") 처리
    - D Sales per Order는 Sales/Orders로 계산해서 넣어줌 (원하면 수식으로 바꿔도 됨)
    """
    headers = ws.row_values(2)  # 2행 헤더
    headers = [h.strip() for h in headers]
    header_map = {h.lower(): i for i, h in enumerate(headers)}

    def get_val(row, name_candidates: List[str]):
        for n in name_candidates:
            if n in row:
                return row[n]
        return ""

    rows = []
    for _, r in df.iterrows():
        # 표준화 dict
        rowd = {c: r[c] for c in df.columns}

        sales = rowd.get("Sales", rowd.get("Sales (SGD)", ""))
        orders = rowd.get("Orders", "")

        # sales per order 계산
        spo = ""
        try:
            if pd.notna(sales) and pd.notna(orders) and float(orders) != 0:
                spo = float(sales) / float(orders)
        except Exception:
            spo = ""

        out = [""] * len(headers)

        def put(col_name: str, value):
            idx = header_map.get(col_name.lower())
            if idx is not None:
                out[idx] = value

        put("Date", rowd.get("Date", rowd.get("date", "")))
        put("Sales (SGD)", sales)   # 시트에 Sales (SGD) 라면
        put("Sales", sales)         # 시트에 Sales 라면
        put("Orders", orders)
        put("Sales per Order", spo)

        put("Product Clicks", rowd.get("Product Clicks", rowd.get("Clicks", "")))
        put("Visitors", rowd.get("Visitors", ""))

        put("Cancelled Orders", rowd.get("Cancelled Orders", rowd.get("Canceled Orders", "")))
        put("Cancelled Sales", rowd.get("Cancelled Sales", rowd.get("Canceled Sales", "")))

        # 나머지(E,F,M,N,O 등)은 export df에 있으면 자동으로 매칭되게 후보를 더 넣고 싶으면 여기 추가
        # 예: put("# of buyers", rowd.get("# of buyers",""))

        rows.append(out)

    return rows


def filter_new_dates(df: pd.DataFrame, last_date: Optional[dt.date]) -> pd.DataFrame:
    if last_date is None:
        return df.copy()
    tmp = df.copy()
    # df에 Date 컬럼이 있든 원래 date_col이 있든 통일
    if "Date" not in tmp.columns:
        # 혹시 date 컬럼명이 다르면
        dcol = None
        for c in tmp.columns:
            if str(c).strip().lower() == "date":
                dcol = c
        if dcol:
            tmp["Date"] = tmp[dcol]
        else:
            raise RuntimeError("No Date column for filtering")

    tmp["__date"] = tmp["Date"].astype(str).apply(parse_mmddyyyy)
    tmp = tmp[tmp["__date"].notna()].copy()
    tmp = tmp[tmp["__date"] > last_date].copy()
    tmp = tmp.drop(columns=["__date"], errors="ignore")
    return tmp


# =========================
# 메인
# =========================
def main():
    with open("config.json", "r", encoding="utf-8") as f:
        cfg = json.load(f)

    spreadsheet_id = cfg["spreadsheet_id"]
    sa_json = cfg["google_service_account_json"]
    session_file = cfg.get("session_file", "shopee_session.json")
    accounts = cfg["accounts"]

    # (1) 세션 없으면 로그인 1회
    # 로그인은 아무 마켓 베이스 하나로만 해도 보통 세션이 공유되지만,
    # 안되면 SG/MY/PH 각각 만들어야 함 (그땐 key별 session_file을 따로 두면 됨)
    ensure_session(session_file, login_market_base=accounts[0]["market_base"])

    # (2) 구글시트 연결
    gc = get_gspread_client(sa_json)
    ss = gc.open_by_key(spreadsheet_id)

    # (3) 각 계정별 Export -> Parse -> Append
    download_dir = os.path.join(os.getcwd(), "downloads")
    os.makedirs(download_dir, exist_ok=True)

    for acc in accounts:
        key = acc["key"]
        tab = acc["tab"]
        market_base = acc["market_base"]
        order_rule = acc["order_rule"]

        print(f"\n===== [{key}] START =====")
        ws = ss.worksheet(tab)

        last_date = get_last_date_in_sheet(ws)
        print(f"[{key}] last_date_in_sheet: {last_date}")

        csv_path = export_datacenter_csv(
            market_base=market_base,
            session_file=session_file,
            download_dir=download_dir,
            account_key=key
        )

        df = parse_export_csv(csv_path, order_rule=order_rule)
        df_new = filter_new_dates(df, last_date)

        if df_new.empty:
            print(f"[{key}] no new rows to append.")
            continue

        # 만약 parse_export_csv가 "일자별 테이블" 그대로 반환한 경우,
        # 컬럼명이 시트와 다를 수 있어서 Date/Sales/Orders 표준화 시도
        if "Date" not in df_new.columns:
            dcol = find_col(df_new, ["Date"])
            if dcol and dcol != "Date":
                df_new = df_new.rename(columns={dcol: "Date"})

        # Sales/Orders 표준화(가능하면)
        sc = find_col(df_new, ["Sales", "Sales (SGD)"])
        oc = find_col(df_new, ["Orders"])
        if sc and sc != "Sales":
            df_new = df_new.rename(columns={sc: "Sales"})
        if oc and oc != "Orders":
            df_new = df_new.rename(columns={oc: "Orders"})

        rows = build_sheet_rows(df_new, ws)
        append_rows(ws, rows)

        print(f"[{key}] appended {len(rows)} rows to {tab}")

    print("\nALL DONE ✅")


if __name__ == "__main__":
    main()
