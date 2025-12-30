import os
import json
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# ===============================
# Google Sheets 인증
# ===============================
service_account_info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_info(
    service_account_info,
    scopes=SCOPES
)

service = build("sheets", "v4", credentials=creds)

SPREADSHEET_ID = "1BF-AIj7KMYGYnOiTEaWQRO_EcXkOYDiEyz-Rl80jwt4"

# ===============================
# 유틸: 마지막 날짜 찾기
# ===============================
def get_last_date(sheet_name: str):
    """
    A3:A 범위에서 마지막 날짜를 찾아 datetime으로 반환
    데이터가 없으면 None
    """
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A3:A"
    ).execute()

    values = result.get("values", [])
    if not values:
        return None

    # 마지막 non-empty 값
    last_date_str = values[-1][0].strip()
    return datetime.strptime(last_date_str, "%m-%d-%Y")


# ===============================
# 유틸: 새 데이터 append
# ===============================
def append_new_rows(sheet_name: str, new_rows: list):
    """
    new_rows: A열(Date)부터 시작하는 row list
    """
    if not new_rows:
        print(f"ℹ️ {sheet_name}: append할 데이터 없음")
        return

    service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A1",
        valueInputOption="USER_ENTERED",
        body={"values": new_rows}
    ).execute()

    print(f"✅ {sheet_name}: {len(new_rows)} rows appended")


# ===============================
# 메인 로직 (탭별 처리)
# ===============================
def process_sheet(sheet_name: str, fetched_rows: list):
    """
    fetched_rows:
      [
        ["12-29-2025", ...],
        ["12-30-2025", ...],
        ...
      ]
    """
    last_date = get_last_date(sheet_name)

    if last_date:
        start_date = last_date + timedelta(days=1)
        print(f"{sheet_name}: last_date={last_date.strftime('%m-%d-%Y')}")
    else:
        # 시트에 데이터가 아예 없을 경우
        start_date = None
        print(f"{sheet_name}: no existing data")

    rows_to_append = []

    for row in fetched_rows:
        row_date = datetime.strptime(row[0], "%m-%d-%Y")

        if start_date is None or row_date >= start_date:
            rows_to_append.append(row)

    append_new_rows(sheet_name, rows_to_append)


# ===============================
# 예시 실행부 (Shopee 연동 전 테스트용)
# ===============================
if __name__ == "__main__":
    # 🔽 나중에 Shopee에서 가져온 데이터로 대체
    example_rows = [
        ["12-29-2025", 1626.96, 82, 19.84, 2238, 3285, "3.66%"],
        ["12-30-2025", 1701.12, 90, 18.90, 2400, 3400, "4.10%"],
    ]

    process_sheet("SG_Sales", example_rows)
