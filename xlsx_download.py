import requests
import datetime
import pandas as pd
import glob
import os

def download_xlsx(week_no, page=1):
    url = 'https://www.yes24.com/Product/Category/BestSellerExcel'
    data = {
        'bestType': 'MONTHWEEK_BESTSELLER',
        'categoryNumber': '001001003',
        'sex': 'A',
        'age': '255',
        'goodsTp': '0',
        'addOptionTp': '0',
        'excludeTp': '2',
        'pageNumber': str(page),
        'pageSize': '120',
        'goodsStatGb': '06',
        'eBookTp': '0',
        'type': 'week',
        'saleYear': '2024',
        'saleMonth': '',
        'weekNo': '1065',
        'day': '0',
        'saleDts': ''
    }

    response = requests.post(url, data=data)
    with open(f'bestseller-{week_no}-page{page}.xlsx', 'wb') as f:
        f.write(response.content)

def calculate_week_no(date=datetime.date.today()):
    # 나중에 시작 주 번호 업데이트
    year = date.year
    return (year-2024)*52 + 1043 + date.isocalendar()[1]

def merge_xlsx(week_no):
    all_data = pd.DataFrame()
    for filename in glob.glob(f'bestseller-{week_no}-*.xlsx'):
        df = pd.read_excel(filename)
        df = df[['순위', 'ISBN', '상품번호', '상품명', '판매가', 'YES포인트', '저자', '출판사', '설명', '출고예상일', '관리분류']]
        df = df.astype(str)
        df['순위'] = df['순위'].astype(int)
        all_data = all_data._append(df, ignore_index=True)
        all_data = all_data.sort_values(by='순위', ascending=True)
    all_data.to_excel(f'bestseller-{week_no}.xlsx', index=False)

    for filename in glob.glob(f'bestseller-{week_no}-*.xlsx'):
        os.remove(filename)

if __name__ == '__main__':
    
    week_no=calculate_week_no()

    for i in range(1, 10):
        download_xlsx(week_no, page=i)
    merge_xlsx(week_no)