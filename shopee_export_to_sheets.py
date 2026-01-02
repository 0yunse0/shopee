import os
import json
import re
import time
import datetime as dt

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright


# =========================
# 기본 설정
# =========================
CONFIG_FILE = "config.json"
SESSION_FILE = "shopee_session.json"
DOWNLOAD_DIR = "downloads"

DATE_FMT = "%d-%m-%Y"   # Export Date 형식 (DD-MM-YYYY)


# =========================
# 유틸
# =========================
def parse_date(s):
    try:
        return dt.datetime.strptime(str(s).strip(), DATE_FMT).date()
    except Exception:
        return None


def today_str():
    return dt.date.today().strftime("%Y%m%d")


# =========================
# Google Sheets
# =========================
def get_gspread_client(sa_json_path):
    creds = Credentials.from_service_account_file(
        sa_json_path,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return gspread.authorize(creds)


def get_last_date(ws):
    """
    시트 구조:
    1행 비어있음
    2행 헤더
    3행부터 데이터
    Date는 A열
    """
    colA = ws.col_values(1)
    last = None
    for v in colA[2:]:
        d = parse_date(v)
        if d:
            last = d
    return last


# =========================
# Playwright: 로그인 세션
# =========================
def ensure_session(login_base):
    if os.path.exists(SESSION_FILE):
        print("[OK] session exists")
        return

    print("[LOGIN] 세션 없음 → 브라우저 열림 (직접 로그인/2FA)")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        page.goto(login_base, wait_until="domcontentloaded")
        page.wait_for_timeout(120000)  # 2분 대기 (직접 로그인)

        context.storage_state(path=SESSION_FILE)
        browser.close()

    print("[OK] session saved")


# =========================
# Playwright: Data Center Export
# =========================
def export_csv(market_base, key):
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            storage_state=SESSION_FILE,
            accept_downloads=True
        )
        page = context.new_page()

        url = f"{market_base}/datacenter/overview"
        print(f"[{key}] goto {url}")
        page.goto(url, wait_until="domcontentloaded")
        page.wait_for_timeout(3000)

        # 기간 선택
        for sel in ["text=Date", "text=Period"]:
            try:
                page.click(sel, timeout=3000)
                break
            except:
                pass

        for sel in ["text=Past 7 days", "text=Last 7 days"]:
            try:
                page.click(sel, timeout=3000)
                break
            except:
                pass

        page.wait_for_timeout(1500)

        # Export
        with page.expect_download() as d:
            page.click("text=Export")
        download = d.value

        name = download.suggested_filename or f"{key}.xlsx"
        name = re.sub(r"[^a-zA-Z0-9._-]", "_", name)
        path = os.path.join(DOWNLOAD_DIR, f"{key}_{today_str()}_{name}")
        download.save_as(path)

        browser.close()

    print(f"[{key}] downloaded {path}")
    return path


# =========================
# Export 엑셀 파싱 (Row 3 header)
# =========================
def parse_export_excel(path):
    # Row 3 (index 2)를 헤더로 사용
    df = pd.read_excel(path, header=2)

    # Date 없는 행 제거
    df = df[df["Date"].notna()].copy()

    # Date 문자열 정리
    df["Date"] = df["Date"].astype(str).str.strip()

    # 숫자 컬럼 정리 (콤마 제거)
    for c in df.columns:
        if c != "Date":
            df[c] = (
                df[c]
                .astype(str)
                .str.replace(",", "", regex=False)
                .replace("nan", "")
            )

    return df


# =========================
# 메인 로직
# =========================
def main():
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    spreadsheet_id = cfg["spreadsheet_id"]
    sa_json = cfg["google_service_account_json"]
    accounts = cfg["accounts"]

    # 로그인 세션 확보 (SG 기준)
    ensure_session(accounts[0]["market_base"])

    # Google Sheets
    gc = get_gspread_client(sa_json)
    ss = gc.open_by_key(spreadsheet_id)

    for acc in accounts:
        key = acc["key"]
        tab = acc["tab"]
        market_base = acc["market_base"]

        print(f"\n===== {key} START =====")
        ws = ss.worksheet(tab)

        last_date = get_last_date(ws)
        print(f"[{key}] last_date = {last_date}")

        excel_path = export_csv(market_base, key)
        df = parse_export_excel(excel_path)

        # 날짜 필터
        df["__date"] = df["Date"].apply(parse_date)
        if last_date:
            df = df[df["__date"] > last_date]

        df = df.drop(columns=["__date"])

        if df.empty:
            print(f"[{key}] no new rows")
            continue

        # 그대로 append (컬럼 구조 동일)
        ws.append_rows(df.values.tolist(), value_input_option="USER_ENTERED")
        print(f"[{key}] appended {len(df)} rows")

    print("\nALL DONE ✅")


if __name__ == "__main__":
    main()
