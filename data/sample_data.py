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

def create_sample_financial_data():
    """재무 데이터 샘플 생성 (회계사 지원용)"""
    np.random.seed(42)
    
    # 회사별 재무 데이터
    companies = ['A기업', 'B기업', 'C기업', 'D기업', 'E기업']
    years = [2021, 2022, 2023]
    
    data = []
    for company in companies:
        for year in years:
            # 매출액 (억원)
            revenue = np.random.randint(1000, 10000)
            # 영업이익 (억원)
            operating_income = revenue * np.random.uniform(0.05, 0.15)
            # 당기순이익 (억원)
            net_income = operating_income * np.random.uniform(0.7, 0.9)
            # 총자산 (억원)
            total_assets = revenue * np.random.uniform(1.5, 3.0)
            # 부채 (억원)
            total_liabilities = total_assets * np.random.uniform(0.3, 0.7)
            # 자본 (억원)
            equity = total_assets - total_liabilities
            
            data.append({
                'Company': company,
                'Year': year,
                'Revenue': revenue,
                'Operating_Income': operating_income,
                'Net_Income': net_income,
                'Total_Assets': total_assets,
                'Total_Liabilities': total_liabilities,
                'Equity': equity,
                'ROE': (net_income / equity) * 100,
                'ROA': (net_income / total_assets) * 100,
                'Debt_Ratio': (total_liabilities / total_assets) * 100
            })
    
    return pd.DataFrame(data)

def create_sample_sales_financial_data():
    """매출 및 재무 데이터 샘플 생성 (회계사 지원용)"""
    np.random.seed(42)
    
    # 월별 매출 데이터
    months = pd.date_range('2023-01-01', '2023-12-31', freq='M')
    departments = ['영업1팀', '영업2팀', '영업3팀', '영업4팀']
    products = ['제품A', '제품B', '제품C', '제품D', '제품E']
    
    data = []
    for month in months:
        for dept in departments:
            for product in products:
                # 기본 매출액
                base_sales = np.random.randint(50000000, 200000000)
                # 계절성 요인 추가
                seasonal_factor = 1 + 0.3 * np.sin(2 * np.pi * month.month / 12)
                # 월별 변동성 추가
                monthly_variation = np.random.uniform(0.8, 1.2)
                
                sales_amount = base_sales * seasonal_factor * monthly_variation
                cost_of_sales = sales_amount * np.random.uniform(0.6, 0.8)
                gross_profit = sales_amount - cost_of_sales
                
                # 세금 및 기타 비용
                tax_rate = np.random.uniform(0.2, 0.25)
                other_expenses = sales_amount * np.random.uniform(0.05, 0.15)
                net_profit = gross_profit - other_expenses - (gross_profit * tax_rate)
                
                data.append({
                    'Date': month,
                    'Department': dept,
                    'Product': product,
                    'Sales_Amount': sales_amount,
                    'Cost_of_Sales': cost_of_sales,
                    'Gross_Profit': gross_profit,
                    'Other_Expenses': other_expenses,
                    'Tax_Expense': gross_profit * tax_rate,
                    'Net_Profit': net_profit,
                    'Profit_Margin': (net_profit / sales_amount) * 100,
                    'Gross_Margin': (gross_profit / sales_amount) * 100
                })
    
    return pd.DataFrame(data)

def create_sample_accounting_data():
    """회계 데이터 샘플 생성 (회계사 지원용)"""
    np.random.seed(42)
    
    # 계정과목별 거래 데이터
    accounts = ['매출액', '매출원가', '판매비', '관리비', '영업외수익', '영업외비용']
    months = pd.date_range('2023-01-01', '2023-12-31', freq='M')
    
    data = []
    for month in months:
        for account in accounts:
            # 계정과목별 기본 금액
            base_amounts = {
                '매출액': np.random.randint(100000000, 500000000),
                '매출원가': np.random.randint(60000000, 300000000),
                '판매비': np.random.randint(20000000, 80000000),
                '관리비': np.random.randint(15000000, 60000000),
                '영업외수익': np.random.randint(5000000, 20000000),
                '영업외비용': np.random.randint(3000000, 15000000)
            }
            
            base_amount = base_amounts[account]
            # 월별 변동성
            variation = np.random.uniform(0.8, 1.2)
            amount = base_amount * variation
            
            data.append({
                'Date': month,
                'Account': account,
                'Amount': amount,
                'Account_Type': '수익' if account in ['매출액', '영업외수익'] else '비용'
            })
    
    return pd.DataFrame(data)

def save_sample_data():
    """모든 샘플 데이터를 엑셀 파일로 저장"""
    with pd.ExcelWriter('data/sample_data.xlsx', engine='openpyxl') as writer:
        create_sample_sales_data().to_excel(writer, sheet_name='Sales', index=False)
        create_sample_employee_data().to_excel(writer, sheet_name='Employees', index=False)
        create_sample_weather_data().to_excel(writer, sheet_name='Weather', index=False)
        create_sample_stock_data().to_excel(writer, sheet_name='Stocks', index=False)
        create_sample_financial_data().to_excel(writer, sheet_name='Financial', index=False)
        create_sample_sales_financial_data().to_excel(writer, sheet_name='Sales_Financial', index=False)
        create_sample_accounting_data().to_excel(writer, sheet_name='Accounting', index=False)
    
    print("샘플 데이터가 'data/sample_data.xlsx'에 저장되었습니다.")
    print("새로 추가된 시트:")
    print("- Financial: 기업별 재무제표 데이터")
    print("- Sales_Financial: 부서별/제품별 매출 및 수익성 데이터")
    print("- Accounting: 계정과목별 회계 거래 데이터")

if __name__ == "__main__":
    save_sample_data() 