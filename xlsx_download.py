import requests
import datetime

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


if __name__ == '__main__':
    for i in range(1, 10):
        download_xlsx(week_no=calculate_week_no(), page=i)