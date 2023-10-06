from google.oauth2.service_account import Credentials
import gspread
import os
from datetime import datetime
import pytz
# 創建時區對象
tz = pytz.timezone('Asia/Taipei')  # 臺灣的時區，可以換成任何pytz支持的時區


scope = ['https://www.googleapis.com/auth/spreadsheets']

creds = Credentials.from_service_account_file(os.path.join("src", "sheet-linebot-6d33dca9428e.json"), scopes=scope)
gs = gspread.authorize(creds)

sheet = gs.open_by_url('https://docs.google.com/spreadsheets/d/12tsMYgX-gIF885hMDhgU9EK-WWwfXJp6P60o3KrYxuM/edit#gid=0')

# worksheet = sheet.get_worksheet(0)
# worksheet.update_value('A1', 'test')

# df = pd.read_csv('Billionaire.csv')
# worksheet.update([df.columns.values.tolist()] + df.values.tolist())
def get_or_create_worksheet(spreadsheet, worksheet_name):
    try:
        # Try to get the worksheet
        worksheet = spreadsheet.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        # If not found, create a new one
        worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows="100", cols="20")
    return worksheet

def save_message(_model, _uid, _question, _answer):
  get_or_create_worksheet(sheet, _uid)
  worksheet = sheet.worksheet(_uid)
  # .add_worksheet(title="新工作表"
  # print(f"{_answer = }")
  now = datetime.now(tz)

  # 取得表格的最後一行數字
  last_row = len(worksheet.get_all_values())
  # 插入新的一行資料
  worksheet.insert_row([str(now), _model, _uid, _question, _answer], last_row+1)
  
  # Python的gspread無法直接設定NumberFormat，因此這部分可能需要在Google Sheets內手動設定或者使用其他方式。

  return last_row + 1