import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Tuple
import streamlit as st


class ExcelReader:
    """엑셀 파일을 읽고 처리하는 클래스"""
    
    def __init__(self):
        self.data = None
        self.sheets = {}
        self.current_sheet = None
    
    def read_excel(self, file_path: str) -> Dict[str, pd.DataFrame]:
        """
        엑셀 파일을 읽어서 모든 시트를 딕셔너리로 반환
        
        Args:
            file_path (str): 엑셀 파일 경로
            
        Returns:
            Dict[str, pd.DataFrame]: 시트명을 키로 하는 데이터프레임 딕셔너리
        """
        try:
            # 모든 시트 읽기
            excel_file = pd.ExcelFile(file_path)
            sheets = {}
            
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                sheets[sheet_name] = df
            
            self.sheets = sheets
            return sheets
            
        except Exception as e:
            st.error(f"엑셀 파일 읽기 오류: {str(e)}")
            return {}
    
    def get_sheet_names(self) -> List[str]:
        """시트 이름 목록 반환"""
        return list(self.sheets.keys())
    
    def get_sheet_data(self, sheet_name: str) -> Optional[pd.DataFrame]:
        """특정 시트의 데이터 반환"""
        return self.sheets.get(sheet_name)
    
    def get_data_info(self, sheet_name: str) -> Dict:
        """데이터 정보 반환"""
        df = self.get_sheet_data(sheet_name)
        if df is None:
            return {}
        
        info = {
            'shape': df.shape,
            'columns': df.columns.tolist(),
            'dtypes': df.dtypes.to_dict(),
            'missing_values': df.isnull().sum().to_dict(),
            'numeric_columns': df.select_dtypes(include=[np.number]).columns.tolist(),
            'categorical_columns': df.select_dtypes(include=['object']).columns.tolist(),
            'date_columns': df.select_dtypes(include=['datetime64']).columns.tolist()
        }
        
        return info
    
    def clean_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """데이터 정리"""
        # 결측값 처리
        df = df.fillna(method='ffill').fillna(method='bfill')
        
        # 중복 행 제거
        df = df.drop_duplicates()
        
        return df
    
    def get_sample_data(self, sheet_name: str, n_rows: int = 5) -> pd.DataFrame:
        """샘플 데이터 반환"""
        df = self.get_sheet_data(sheet_name)
        if df is not None:
            return df.head(n_rows)
        return pd.DataFrame()


def create_sample_excel():
    """샘플 엑셀 데이터 생성"""
    import io
    
    # 샘플 데이터 생성
    sales_data = {
        'Date': pd.date_range('2023-01-01', periods=100, freq='D'),
        'Product': np.random.choice(['A', 'B', 'C', 'D'], 100),
        'Sales': np.random.randint(100, 1000, 100),
        'Quantity': np.random.randint(1, 50, 100),
        'Region': np.random.choice(['North', 'South', 'East', 'West'], 100)
    }
    
    employee_data = {
        'Name': ['John', 'Jane', 'Bob', 'Alice', 'Charlie', 'Diana', 'Eve', 'Frank'],
        'Age': [25, 30, 35, 28, 32, 27, 29, 31],
        'Salary': [50000, 60000, 70000, 55000, 65000, 58000, 62000, 68000],
        'Department': ['IT', 'HR', 'Sales', 'IT', 'Marketing', 'HR', 'Sales', 'IT']
    }
    
    # ExcelWriter로 여러 시트 생성
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame(sales_data).to_excel(writer, sheet_name='Sales', index=False)
        pd.DataFrame(employee_data).to_excel(writer, sheet_name='Employees', index=False)
    
    output.seek(0)
    return output 