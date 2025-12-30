import os
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime

# 1. GitHub Secret에서 JSON 읽기
service_account_info = json.loads(
    os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
)

# 2. Google Sheets 인증
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_info(
    service_account_info,
    scopes=SCOPES
)

service = build("sheets", "v4", credentials=creds)

# 3. 시트 정보
SPREADSHEET_ID = "1BF-AIj7KMYGYnOiTEaWQRO_EcXkOYDiEyz-Rl80jwt4"
SHEET_NAME = "SG_Sales"  # 실제 시트 탭 이름과 정확히 동일

# 4. 입력할 테스트 데이터
now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
values = [[
    now,
    "TEST",
    "FROM",
    "GITHUB",
    "ACTIONS"
]]

# 5. 시트에 append
service.spreadsheets().values().append(
    spreadsheetId=SPREADSHEET_ID,
    range=f"{SHEET_NAME}!A1",
    valueInputOption="USER_ENTERED",
    body={"values": values}
).execute()

print("✅ Sheet write success")
