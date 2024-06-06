import pandas as pd
import gspread
import datetime
import os
from dotenv import load_dotenv



from calculate_week_no import calculate_week_no

def update_gsheet():
    load_dotenv()
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    year = datetime.date.today().year
    week = datetime.date.today().isocalendar()[1]
    sheet_name = f'{year}-w{week}'

    gc = gspread.service_account(filename='credentials.json')
    sh = gc.open_by_key(sheet_id)
    worksheet = sh.add_worksheet(title=sheet_name, rows=1001, cols=11)

    week_no = calculate_week_no()
    # upload bestseller-{week_no}.xlsx to the worksheet
    df = pd.read_excel(f'bestseller-{week_no}.xlsx')
    df = df.astype(str)
    df['순위'] = df['순위'].astype(int)
    worksheet.update(range_name='A1', values=[df.columns.values.tolist()] + df.values.tolist())

    worksheet.format('A1:K1', {'textFormat': {'bold': True}})
    worksheet.freeze(rows=1)

    # open 

if __name__ == '__main__':
    update_gsheet()