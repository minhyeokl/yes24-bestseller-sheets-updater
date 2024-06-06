import datetime

def calculate_week_no(date=datetime.date.today()):
    # 나중에 시작 주 번호 업데이트
    year = date.year
    return (year-2024)*52 + 1043 + date.isocalendar()[1]