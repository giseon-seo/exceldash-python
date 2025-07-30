import pandas as pd
import numpy as np
from scipy import stats
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
from sklearn.cluster import KMeans
from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score, mean_squared_error
import plotly.graph_objects as go
import plotly.express as px
from typing import Dict, List, Tuple, Optional
import warnings
warnings.filterwarnings('ignore')


class DataAnalyzer:
    """전문적인 데이터 분석 클래스"""
    
    def __init__(self):
        self.analysis_results = {}
    
    def descriptive_statistics(self, df: pd.DataFrame) -> Dict:
        """기술통계 분석"""
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        
        stats_dict = {}
        for col in numeric_cols:
            stats_dict[col] = {
                'count': df[col].count(),
                'mean': df[col].mean(),
                'median': df[col].median(),
                'std': df[col].std(),
                'min': df[col].min(),
                'max': df[col].max(),
                'q1': df[col].quantile(0.25),
                'q3': df[col].quantile(0.75),
                'skewness': df[col].skew(),
                'kurtosis': df[col].kurtosis(),
                'missing_count': df[col].isnull().sum(),
                'missing_percent': (df[col].isnull().sum() / len(df)) * 100
            }
        
        return stats_dict
    
    def correlation_analysis(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict]:
        """상관관계 분석"""
        numeric_df = df.select_dtypes(include=[np.number])
        
        # 상관계수 계산
        corr_matrix = numeric_df.corr()
        
        # 유의성 검정
        p_values = {}
        for i in numeric_df.columns:
            p_values[i] = {}
            for j in numeric_df.columns:
                if i != j:
                    correlation, p_value = stats.pearsonr(numeric_df[i].dropna(), numeric_df[j].dropna())
                    p_values[i][j] = p_value
        
        return corr_matrix, p_values
    
    def outlier_detection(self, df: pd.DataFrame, method: str = 'iqr') -> Dict:
        """이상치 탐지"""
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        outliers_dict = {}
        
        for col in numeric_cols:
            if method == 'iqr':
                Q1 = df[col].quantile(0.25)
                Q3 = df[col].quantile(0.75)
                IQR = Q3 - Q1
                lower_bound = Q1 - 1.5 * IQR
                upper_bound = Q3 + 1.5 * IQR
                
                outliers = df[(df[col] < lower_bound) | (df[col] > upper_bound)]
                
                outliers_dict[col] = {
                    'outlier_count': len(outliers),
                    'outlier_percent': (len(outliers) / len(df)) * 100,
                    'lower_bound': lower_bound,
                    'upper_bound': upper_bound,
                    'outlier_indices': outliers.index.tolist()
                }
            
            elif method == 'zscore':
                z_scores = np.abs(stats.zscore(df[col].dropna()))
                outliers = df[z_scores > 3]
                
                outliers_dict[col] = {
                    'outlier_count': len(outliers),
                    'outlier_percent': (len(outliers) / len(df)) * 100,
                    'z_score_threshold': 3,
                    'outlier_indices': outliers.index.tolist()
                }
        
        return outliers_dict
    
    def normality_test(self, df: pd.DataFrame) -> Dict:
        """정규성 검정"""
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        normality_results = {}
        
        for col in numeric_cols:
            data = df[col].dropna()
            if len(data) > 3:
                # Shapiro-Wilk 검정
                shapiro_stat, shapiro_p = stats.shapiro(data)
                # Kolmogorov-Smirnov 검정
                ks_stat, ks_p = stats.kstest(data, 'norm', args=(data.mean(), data.std()))
                
                normality_results[col] = {
                    'shapiro_statistic': shapiro_stat,
                    'shapiro_p_value': shapiro_p,
                    'ks_statistic': ks_stat,
                    'ks_p_value': ks_p,
                    'is_normal_shapiro': shapiro_p > 0.05,
                    'is_normal_ks': ks_p > 0.05
                }
        
        return normality_results
    
    def trend_analysis(self, df: pd.DataFrame, date_col: str, value_col: str) -> Dict:
        """시계열 트렌드 분석"""
        df_copy = df.copy()
        df_copy[date_col] = pd.to_datetime(df_copy[date_col])
        df_copy = df_copy.sort_values(date_col)
        
        # 선형 회귀 분석
        X = np.arange(len(df_copy)).reshape(-1, 1)
        y = df_copy[value_col].values
        
        model = LinearRegression()
        model.fit(X, y)
        y_pred = model.predict(X)
        
        # 추세 분석 결과
        trend_results = {
            'slope': model.coef_[0],
            'intercept': model.intercept_,
            'r_squared': r2_score(y, y_pred),
            'mse': mean_squared_error(y, y_pred),
            'trend_direction': 'increasing' if model.coef_[0] > 0 else 'decreasing',
            'trend_strength': abs(model.coef_[0])
        }
        
        return trend_results
    
    def seasonal_analysis(self, df: pd.DataFrame, date_col: str, value_col: str) -> Dict:
        """계절성 분석"""
        df_copy = df.copy()
        df_copy[date_col] = pd.to_datetime(df_copy[date_col])
        
        # 월별 평균
        df_copy['month'] = df_copy[date_col].dt.month
        monthly_avg = df_copy.groupby('month')[value_col].mean()
        
        # 분기별 평균
        df_copy['quarter'] = df_copy[date_col].dt.quarter
        quarterly_avg = df_copy.groupby('quarter')[value_col].mean()
        
        seasonal_results = {
            'monthly_pattern': monthly_avg.to_dict(),
            'quarterly_pattern': quarterly_avg.to_dict(),
            'seasonal_strength': monthly_avg.std() / monthly_avg.mean() if monthly_avg.mean() != 0 else 0
        }
        
        return seasonal_results
    
    def cluster_analysis(self, df: pd.DataFrame, n_clusters: int = 3) -> Dict:
        """군집 분석"""
        numeric_df = df.select_dtypes(include=[np.number])
        
        if len(numeric_df.columns) < 2:
            return {}
        
        # 데이터 정규화
        scaler = StandardScaler()
        scaled_data = scaler.fit_transform(numeric_df)
        
        # K-means 군집화
        kmeans = KMeans(n_clusters=n_clusters, random_state=42)
        clusters = kmeans.fit_predict(scaled_data)
        
        # 군집 분석 결과
        cluster_results = {
            'n_clusters': n_clusters,
            'cluster_labels': clusters.tolist(),
            'cluster_centers': kmeans.cluster_centers_.tolist(),
            'inertia': kmeans.inertia_,
            'cluster_sizes': np.bincount(clusters).tolist()
        }
        
        return cluster_results
    
    def pca_analysis(self, df: pd.DataFrame, n_components: int = 2) -> Dict:
        """주성분 분석"""
        numeric_df = df.select_dtypes(include=[np.number])
        
        if len(numeric_df.columns) < 2:
            return {}
        
        # 데이터 정규화
        scaler = StandardScaler()
        scaled_data = scaler.fit_transform(numeric_df)
        
        # PCA 분석
        pca = PCA(n_components=min(n_components, len(numeric_df.columns)))
        pca_result = pca.fit_transform(scaled_data)
        
        pca_results = {
            'explained_variance_ratio': pca.explained_variance_ratio_.tolist(),
            'cumulative_variance_ratio': np.cumsum(pca.explained_variance_ratio_).tolist(),
            'components': pca.components_.tolist(),
            'feature_names': numeric_df.columns.tolist()
        }
        
        return pca_results
    
    def financial_analysis(self, df: pd.DataFrame) -> Dict:
        """재무 분석"""
        financial_metrics = {}
        
        # 수익성 지표
        if 'Revenue' in df.columns and 'Net_Income' in df.columns:
            financial_metrics['profit_margin'] = (df['Net_Income'] / df['Revenue'] * 100).mean()
        
        if 'Revenue' in df.columns and 'Gross_Profit' in df.columns:
            financial_metrics['gross_margin'] = (df['Gross_Profit'] / df['Revenue'] * 100).mean()
        
        # 수익성 지표
        if 'Net_Income' in df.columns and 'Equity' in df.columns:
            financial_metrics['roe'] = (df['Net_Income'] / df['Equity'] * 100).mean()
        
        if 'Net_Income' in df.columns and 'Total_Assets' in df.columns:
            financial_metrics['roa'] = (df['Net_Income'] / df['Total_Assets'] * 100).mean()
        
        # 안정성 지표
        if 'Total_Liabilities' in df.columns and 'Total_Assets' in df.columns:
            financial_metrics['debt_ratio'] = (df['Total_Liabilities'] / df['Total_Assets'] * 100).mean()
        
        # 성장성 지표
        if 'Revenue' in df.columns:
            financial_metrics['revenue_growth'] = df['Revenue'].pct_change().mean() * 100
        
        return financial_metrics
    
    def create_analysis_report(self, df: pd.DataFrame) -> Dict:
        """종합 분석 리포트 생성"""
        report = {
            'data_overview': {
                'shape': df.shape,
                'columns': df.columns.tolist(),
                'dtypes': df.dtypes.to_dict(),
                'missing_data': df.isnull().sum().to_dict()
            },
            'descriptive_statistics': self.descriptive_statistics(df),
            'correlation_analysis': self.correlation_analysis(df),
            'outlier_analysis': self.outlier_detection(df),
            'normality_test': self.normality_test(df),
            'financial_analysis': self.financial_analysis(df)
        }
        
        # 시계열 분석 (날짜 컬럼이 있는 경우)
        date_cols = df.select_dtypes(include=['datetime64']).columns
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        
        if len(date_cols) > 0 and len(numeric_cols) > 0:
            date_col = date_cols[0]
            value_col = numeric_cols[0]
            
            report['trend_analysis'] = self.trend_analysis(df, date_col, value_col)
            report['seasonal_analysis'] = self.seasonal_analysis(df, date_col, value_col)
        
        # 군집 분석 (수치형 컬럼이 2개 이상인 경우)
        if len(numeric_cols) >= 2:
            report['cluster_analysis'] = self.cluster_analysis(df)
            report['pca_analysis'] = self.pca_analysis(df)
        
        return report 