import gspread
from google.oauth2.service_account import Credentials

SERVICE_ACCOUNT_FILE = "dulcet-antler-462703-n8-d2fbdb362407.json"
SPREADSHEET_ID = "19v86VBcbzEzzF4I6g8RgCthostf3YFTAQaFFDjQuYlk"

scopes = ['https://www.googleapis.com/auth/spreadsheets']
credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=scopes
)
gc = gspread.authorize(credentials)

# 시트 열기 시도
try:
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.sheet1
    rows = ws.get_all_values()
    print(f"Row count: {len(rows)}")
except Exception as e:
    import traceback
    print(traceback.format_exc())
