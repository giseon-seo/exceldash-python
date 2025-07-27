import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import seaborn as sns
from typing import Dict, List, Optional, Tuple
import streamlit as st


class ChartCreator:
    """다양한 차트를 생성하는 클래스"""
    
    def __init__(self):
        self.chart_types = {
            'bar': '막대그래프',
            'line': '선그래프',
            'pie': '파이차트',
            'scatter': '산점도',
            'histogram': '히스토그램',
            'box': '박스플롯',
            'heatmap': '히트맵',
            'area': '영역차트'
        }
    
    def create_bar_chart(self, df: pd.DataFrame, x_col: str, y_col: str, 
                         color_col: Optional[str] = None, title: str = "막대그래프") -> go.Figure:
        """막대그래프 생성"""
        fig = px.bar(df, x=x_col, y=y_col, color=color_col, title=title)
        fig.update_layout(
            xaxis_title=x_col,
            yaxis_title=y_col,
            template="plotly_white"
        )
        return fig
    
    def create_line_chart(self, df: pd.DataFrame, x_col: str, y_col: str,
                         color_col: Optional[str] = None, title: str = "선그래프") -> go.Figure:
        """선그래프 생성"""
        fig = px.line(df, x=x_col, y=y_col, color=color_col, title=title)
        fig.update_layout(
            xaxis_title=x_col,
            yaxis_title=y_col,
            template="plotly_white"
        )
        return fig
    
    def create_pie_chart(self, df: pd.DataFrame, values_col: str, names_col: str,
                         title: str = "파이차트") -> go.Figure:
        """파이차트 생성"""
        fig = px.pie(df, values=values_col, names=names_col, title=title)
        fig.update_layout(template="plotly_white")
        return fig
    
    def create_scatter_plot(self, df: pd.DataFrame, x_col: str, y_col: str,
                           color_col: Optional[str] = None, size_col: Optional[str] = None,
                           title: str = "산점도") -> go.Figure:
        """산점도 생성"""
        fig = px.scatter(df, x=x_col, y=y_col, color=color_col, size=size_col, title=title)
        fig.update_layout(
            xaxis_title=x_col,
            yaxis_title=y_col,
            template="plotly_white"
        )
        return fig
    
    def create_histogram(self, df: pd.DataFrame, column: str, bins: int = 30,
                        title: str = "히스토그램") -> go.Figure:
        """히스토그램 생성"""
        fig = px.histogram(df, x=column, nbins=bins, title=title)
        fig.update_layout(
            xaxis_title=column,
            yaxis_title="빈도",
            template="plotly_white"
        )
        return fig
    
    def create_box_plot(self, df: pd.DataFrame, x_col: str, y_col: str,
                       title: str = "박스플롯") -> go.Figure:
        """박스플롯 생성"""
        fig = px.box(df, x=x_col, y=y_col, title=title)
        fig.update_layout(
            xaxis_title=x_col,
            yaxis_title=y_col,
            template="plotly_white"
        )
        return fig
    
    def create_heatmap(self, df: pd.DataFrame, columns: Optional[List[str]] = None,
                      title: str = "히트맵") -> go.Figure:
        """히트맵 생성"""
        if columns is None:
            # 수치형 컬럼만 선택
            numeric_df = df.select_dtypes(include=[np.number])
        else:
            numeric_df = df[columns].select_dtypes(include=[np.number])
        
        if numeric_df.empty:
            st.warning("히트맵을 생성할 수 있는 수치형 데이터가 없습니다.")
            return go.Figure()
        
        corr_matrix = numeric_df.corr()
        
        fig = px.imshow(
            corr_matrix,
            text_auto=True,
            aspect="auto",
            title=title,
            color_continuous_scale="RdBu"
        )
        fig.update_layout(template="plotly_white")
        return fig
    
    def create_area_chart(self, df: pd.DataFrame, x_col: str, y_col: str,
                         color_col: Optional[str] = None, title: str = "영역차트") -> go.Figure:
        """영역차트 생성"""
        fig = px.area(df, x=x_col, y=y_col, color=color_col, title=title)
        fig.update_layout(
            xaxis_title=x_col,
            yaxis_title=y_col,
            template="plotly_white"
        )
        return fig
    
    def create_summary_stats(self, df: pd.DataFrame) -> pd.DataFrame:
        """요약 통계 생성"""
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) == 0:
            return pd.DataFrame()
        
        summary = df[numeric_cols].describe()
        return summary
    
    def get_chart_options(self, df: pd.DataFrame) -> Dict[str, List[str]]:
        """데이터에 따른 차트 옵션 반환"""
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
        date_cols = df.select_dtypes(include=['datetime64']).columns.tolist()
        
        options = {
            'bar': {
                'x': categorical_cols + numeric_cols,
                'y': numeric_cols,
                'color': categorical_cols
            },
            'line': {
                'x': date_cols + categorical_cols,
                'y': numeric_cols,
                'color': categorical_cols
            },
            'pie': {
                'values': numeric_cols,
                'names': categorical_cols
            },
            'scatter': {
                'x': numeric_cols,
                'y': numeric_cols,
                'color': categorical_cols,
                'size': numeric_cols
            },
            'histogram': {
                'column': numeric_cols
            },
            'box': {
                'x': categorical_cols,
                'y': numeric_cols
            },
            'area': {
                'x': date_cols + categorical_cols,
                'y': numeric_cols,
                'color': categorical_cols
            }
        }
        
        return options
    
    def create_dashboard_layout(self, charts: List[go.Figure], titles: List[str]) -> go.Figure:
        """대시보드 레이아웃 생성"""
        if len(charts) == 0:
            return go.Figure()
        
        # 서브플롯 생성
        rows = (len(charts) + 1) // 2
        cols = min(2, len(charts))
        
        fig = go.Figure()
        
        for i, (chart, title) in enumerate(zip(charts, titles)):
            row = i // 2 + 1
            col = i % 2 + 1
            
            # 차트를 서브플롯에 추가
            for trace in chart.data:
                trace.update(xaxis=f'x{i+1}', yaxis=f'y{i+1}')
                fig.add_trace(trace)
        
        # 레이아웃 설정
        fig.update_layout(
            title="대시보드",
            template="plotly_white",
            showlegend=False
        )
        
        return fig 