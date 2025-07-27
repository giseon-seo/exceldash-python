import pandas as pd
import numpy as np
from datetime import datetime, timedelta

def create_sample_sales_data():
    """판매 데이터 샘플 생성"""
    np.random.seed(42)
    
    # 날짜 범위 생성
    start_date = datetime(2023, 1, 1)
    end_date = datetime(2023, 12, 31)
    dates = pd.date_range(start_date, end_date, freq='D')
    
    # 데이터 생성
    data = {
        'Date': np.random.choice(dates, 1000),
        'Product': np.random.choice(['노트북', '태블릿', '스마트폰', '헤드폰', '키보드'], 1000),
        'Category': np.random.choice(['전자제품', '액세서리'], 1000),
        'Region': np.random.choice(['서울', '부산', '대구', '인천', '광주'], 1000),
        'Sales': np.random.randint(100000, 5000000, 1000),
        'Quantity': np.random.randint(1, 50, 1000),
        'Customer_Rating': np.random.uniform(1, 5, 1000).round(1)
    }
    
    return pd.DataFrame(data)

def create_sample_employee_data():
    """직원 데이터 샘플 생성"""
    np.random.seed(42)
    
    departments = ['개발팀', '마케팅팀', '영업팀', '인사팀', '재무팀']
    positions = ['사원', '대리', '과장', '차장', '부장']
    
    data = {
        'Employee_ID': range(1, 101),
        'Name': [f'직원{i}' for i in range(1, 101)],
        'Age': np.random.randint(25, 60, 100),
        'Department': np.random.choice(departments, 100),
        'Position': np.random.choice(positions, 100),
        'Salary': np.random.randint(30000000, 80000000, 100),
        'Years_Experience': np.random.randint(1, 20, 100),
        'Performance_Score': np.random.uniform(60, 100, 100).round(1)
    }
    
    return pd.DataFrame(data)

def create_sample_weather_data():
    """날씨 데이터 샘플 생성"""
    np.random.seed(42)
    
    # 날짜 범위 생성
    start_date = datetime(2023, 1, 1)
    end_date = datetime(2023, 12, 31)
    dates = pd.date_range(start_date, end_date, freq='D')
    
    # 계절별 온도 패턴
    def get_temp_by_date(date):
        day_of_year = date.timetuple().tm_yday
        # 사인 함수를 사용하여 계절적 변화 생성
        base_temp = 15 + 10 * np.sin(2 * np.pi * day_of_year / 365)
        return base_temp + np.random.normal(0, 5)
    
    data = {
        'Date': dates,
        'Temperature': [get_temp_by_date(date) for date in dates],
        'Humidity': np.random.uniform(30, 90, len(dates)),
        'Precipitation': np.random.exponential(5, len(dates)),
        'Wind_Speed': np.random.exponential(3, len(dates)),
        'Weather_Condition': np.random.choice(['맑음', '흐림', '비', '눈'], len(dates))
    }
    
    return pd.DataFrame(data)

def create_sample_stock_data():
    """주식 데이터 샘플 생성"""
    np.random.seed(42)
    
    # 날짜 범위 생성
    start_date = datetime(2023, 1, 1)
    end_date = datetime(2023, 12, 31)
    dates = pd.date_range(start_date, end_date, freq='D')
    
    # 주가 시뮬레이션 (랜덤 워크)
    def generate_stock_price(initial_price, days):
        returns = np.random.normal(0, 0.02, days)
        prices = [initial_price]
        for ret in returns:
            prices.append(prices[-1] * (1 + ret))
        return prices[1:]  # 첫 번째 값 제외
    
    companies = ['삼성전자', 'SK하이닉스', 'LG화학', '현대차', 'POSCO']
    
    data = []
    for company in companies:
        initial_price = np.random.randint(50000, 200000)
        prices = generate_stock_price(initial_price, len(dates))
        volumes = np.random.randint(1000000, 10000000, len(dates))
        
        for i, date in enumerate(dates):
            data.append({
                'Date': date,
                'Company': company,
                'Price': prices[i],
                'Volume': volumes[i],
                'Change': (prices[i] - initial_price) / initial_price * 100
            })
    
    return pd.DataFrame(data)

def save_sample_data():
    """모든 샘플 데이터를 엑셀 파일로 저장"""
    with pd.ExcelWriter('data/sample_data.xlsx', engine='openpyxl') as writer:
        create_sample_sales_data().to_excel(writer, sheet_name='Sales', index=False)
        create_sample_employee_data().to_excel(writer, sheet_name='Employees', index=False)
        create_sample_weather_data().to_excel(writer, sheet_name='Weather', index=False)
        create_sample_stock_data().to_excel(writer, sheet_name='Stocks', index=False)
    
    print("샘플 데이터가 'data/sample_data.xlsx'에 저장되었습니다.")

if __name__ == "__main__":
    save_sample_data() 