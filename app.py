import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from utils.excel_reader import ExcelReader, create_sample_excel
from utils.chart_creator import ChartCreator
from utils.data_analyzer import DataAnalyzer
import os

# 페이지 설정
st.set_page_config(
    page_title="Excel Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS 스타일
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 0.5rem 0;
    }
    .chart-container {
        border: 1px solid #e0e0e0;
        border-radius: 0.5rem;
        padding: 1rem;
        margin: 1rem 0;
    }
    .dashboard-section {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    .analysis-card {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        border-radius: 0.5rem;
        padding: 1rem;
        margin: 0.5rem 0;
    }
</style>
""", unsafe_allow_html=True)

def create_dashboard_charts(df, chart_creator):
    """대시보드용 차트들을 생성"""
    charts = []
    
    # 수치형 컬럼들
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
    date_cols = df.select_dtypes(include=['datetime64']).columns.tolist()
    
    # 1. 상관관계 히트맵 (수치형 컬럼이 2개 이상인 경우)
    if len(numeric_cols) >= 2:
        fig_heatmap = chart_creator.create_heatmap(df, numeric_cols, "상관관계 히트맵")
        charts.append(("상관관계 히트맵", fig_heatmap))
    
    # 2. 주요 수치형 데이터 분포 (히스토그램)
    if numeric_cols:
        selected_numeric = numeric_cols[0]
        fig_hist = chart_creator.create_histogram(df, selected_numeric, 20, f"{selected_numeric} 분포")
        charts.append((f"{selected_numeric} 분포", fig_hist))
    
    # 3. 범주별 수치 분포 (박스플롯)
    if categorical_cols and numeric_cols:
        cat_col = categorical_cols[0]
        num_col = numeric_cols[0]
        fig_box = chart_creator.create_box_plot(df, cat_col, num_col, f"{cat_col}별 {num_col} 분포")
        charts.append((f"{cat_col}별 {num_col} 분포", fig_box))
    
    # 4. 범주별 집계 (막대그래프)
    if categorical_cols and numeric_cols:
        cat_col = categorical_cols[0]
        num_col = numeric_cols[0]
        fig_bar = chart_creator.create_bar_chart(df, cat_col, num_col, title=f"{cat_col}별 {num_col}")
        charts.append((f"{cat_col}별 {num_col}", fig_bar))
    
    # 5. 비율 분석 (파이차트)
    if categorical_cols and numeric_cols:
        cat_col = categorical_cols[0]
        num_col = numeric_cols[0]
        pie_data = df.groupby(cat_col)[num_col].sum().reset_index()
        fig_pie = chart_creator.create_pie_chart(pie_data, num_col, cat_col, f"{cat_col}별 {num_col} 비율")
        charts.append((f"{cat_col}별 {num_col} 비율", fig_pie))
    
    # 6. 변수 간 관계 (산점도)
    if len(numeric_cols) >= 2:
        x_col = numeric_cols[0]
        y_col = numeric_cols[1]
        fig_scatter = chart_creator.create_scatter_plot(df, x_col, y_col, title=f"{x_col} vs {y_col}")
        charts.append((f"{x_col} vs {y_col}", fig_scatter))
    
    # 7. 시계열 분석 (날짜 데이터가 있는 경우)
    if date_cols and numeric_cols:
        date_col = date_cols[0]
        num_col = numeric_cols[0]
        
        # 월별 집계
        df_copy = df.copy()
        df_copy['Month'] = df_copy[date_col].dt.to_period('M')
        monthly_data = df_copy.groupby('Month')[num_col].agg(['sum', 'mean']).reset_index()
        monthly_data['Month'] = monthly_data['Month'].astype(str)
        
        fig_monthly = chart_creator.create_bar_chart(monthly_data, 'Month', 'sum', title=f"월별 {num_col} 합계")
        charts.append((f"월별 {num_col} 합계", fig_monthly))
    
    # 8. 상위 분석
    if categorical_cols and numeric_cols:
        cat_col = categorical_cols[0]
        num_col = numeric_cols[0]
        
        top_data = df.groupby(cat_col)[num_col].sum().sort_values(ascending=False).head(10).reset_index()
        fig_top = chart_creator.create_bar_chart(top_data, cat_col, num_col, title=f"상위 10개 {cat_col}별 {num_col}")
        charts.append((f"상위 10개 {cat_col}별 {num_col}", fig_top))
    
    return charts

def display_dashboard(df, chart_creator):
    """개선된 대시보드 표시"""
    st.header("📊 데이터 대시보드")
    
    # 데이터 요약 정보
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("총 행 수", len(df))
    with col2:
        st.metric("총 열 수", len(df.columns))
    with col3:
        numeric_count = len(df.select_dtypes(include=['number']).columns)
        st.metric("수치형 컬럼", numeric_count)
    with col4:
        categorical_count = len(df.select_dtypes(include=['object']).columns)
        st.metric("범주형 컬럼", categorical_count)
    
    # 대시보드 차트 생성
    dashboard_charts = create_dashboard_charts(df, chart_creator)
    
    if dashboard_charts:
        st.markdown("---")
        st.subheader("📈 주요 분석 차트")
        
        # 차트를 2x2 그리드로 배치
        for i in range(0, len(dashboard_charts), 2):
            col1, col2 = st.columns(2)
            
            with col1:
                title, fig = dashboard_charts[i]
                st.markdown(f"**{title}**")
                st.plotly_chart(fig, use_container_width=True, height=400)
            
            if i + 1 < len(dashboard_charts):
                with col2:
                    title, fig = dashboard_charts[i + 1]
                    st.markdown(f"**{title}**")
                    st.plotly_chart(fig, use_container_width=True, height=400)
            else:
                with col2:
                    st.empty()
    else:
        st.warning("대시보드를 생성할 수 있는 충분한 데이터가 없습니다.")

def display_advanced_analysis(df, analyzer):
    """고급 분석 결과 표시"""
    st.header("🔬 고급 데이터 분석")
    
    # 분석 리포트 생성
    with st.spinner("전문적인 데이터 분석을 수행하고 있습니다..."):
        report = analyzer.create_analysis_report(df)
    
    # 탭으로 분석 결과 구분
    tab1, tab2, tab3, tab4 = st.tabs(["📊 기술통계", "🔗 상관관계", "⚠️ 이상치/정규성", "📈 시계열/군집"])
    
    with tab1:
        st.subheader("📊 기술통계 분석")
        desc_stats = report['descriptive_statistics']
        
        for col, stats in desc_stats.items():
            with st.expander(f"{col} 기술통계"):
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("평균", f"{stats['mean']:.2f}")
                    st.metric("중앙값", f"{stats['median']:.2f}")
                with col2:
                    st.metric("표준편차", f"{stats['std']:.2f}")
                    st.metric("분산", f"{stats['std']**2:.2f}")
                with col3:
                    st.metric("최소값", f"{stats['min']:.2f}")
                    st.metric("최대값", f"{stats['max']:.2f}")
                with col4:
                    st.metric("왜도", f"{stats['skewness']:.2f}")
                    st.metric("첨도", f"{stats['kurtosis']:.2f}")
    
    with tab2:
        st.subheader("🔗 상관관계 분석")
        corr_matrix, p_values = report['correlation_analysis']
        
        if not corr_matrix.empty:
            st.write("**상관계수 행렬**")
            st.dataframe(corr_matrix.round(3))
            
            # 유의한 상관관계 표시
            significant_correlations = []
            for i in corr_matrix.columns:
                for j in corr_matrix.columns:
                    if i != j and p_values.get(i, {}).get(j, 1) < 0.05:
                        significant_correlations.append({
                            '변수1': i,
                            '변수2': j,
                            '상관계수': corr_matrix.loc[i, j],
                            'p값': p_values[i][j]
                        })
            
            if significant_correlations:
                st.write("**유의한 상관관계 (p < 0.05)**")
                sig_df = pd.DataFrame(significant_correlations)
                st.dataframe(sig_df.round(4))
    
    with tab3:
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("⚠️ 이상치 분석")
            outlier_analysis = report['outlier_analysis']
            
            for col, outlier_info in outlier_analysis.items():
                with st.expander(f"{col} 이상치"):
                    st.metric("이상치 개수", outlier_info['outlier_count'])
                    st.metric("이상치 비율", f"{outlier_info['outlier_percent']:.2f}%")
        
        with col2:
            st.subheader("📈 정규성 검정")
            normality_test = report['normality_test']
            
            for col, test_results in normality_test.items():
                with st.expander(f"{col} 정규성"):
                    st.write(f"**Shapiro-Wilk**: {'정규분포' if test_results['is_normal_shapiro'] else '비정규분포'}")
                    st.write(f"**KS 검정**: {'정규분포' if test_results['is_normal_ks'] else '비정규분포'}")
    
    with tab4:
        col1, col2 = st.columns(2)
        
        with col1:
            if 'trend_analysis' in report:
                st.subheader("📈 시계열 분석")
                trend_analysis = report['trend_analysis']
                st.metric("기울기", f"{trend_analysis['slope']:.4f}")
                st.metric("R²", f"{trend_analysis['r_squared']:.4f}")
                st.metric("추세 방향", trend_analysis['trend_direction'])
        
        with col2:
            if 'cluster_analysis' in report:
                st.subheader("🎯 군집 분석")
                cluster_analysis = report['cluster_analysis']
                st.metric("군집 수", cluster_analysis['n_clusters'])
                st.metric("Inertia", f"{cluster_analysis['inertia']:.2f}")
    
    # 재무 분석 (별도 섹션)
    if 'financial_analysis' in report and report['financial_analysis']:
        st.markdown("---")
        st.subheader("💰 재무 분석")
        financial_analysis = report['financial_analysis']
        
        metrics_cols = st.columns(len(financial_analysis))
        for i, (metric, value) in enumerate(financial_analysis.items()):
            with metrics_cols[i]:
                st.metric(metric.replace('_', ' ').title(), f"{value:.2f}%")

def main():
    # 헤더
    st.markdown('<h1 class="main-header">📊 Excel Dashboard</h1>', unsafe_allow_html=True)
    
    # 사이드바
    with st.sidebar:
        st.header("📁 데이터 업로드")
        
        # 파일 업로드
        uploaded_file = st.file_uploader(
            "엑셀 파일을 선택하세요",
            type=['xlsx', 'xls'],
            help="Excel 파일(.xlsx, .xls)을 업로드하세요"
        )
        
        # 샘플 데이터 다운로드
        st.markdown("---")
        st.header("📋 샘플 데이터")
        sample_data = create_sample_excel()
        st.download_button(
            label="샘플 엑셀 파일 다운로드",
            data=sample_data.getvalue(),
            file_name="sample_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # 메인 컨텐츠
    if uploaded_file is not None:
        # 엑셀 파일 읽기
        excel_reader = ExcelReader()
        chart_creator = ChartCreator()
        data_analyzer = DataAnalyzer()
        
        # 임시 파일로 저장
        with open("temp_file.xlsx", "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # 엑셀 파일 읽기
        sheets = excel_reader.read_excel("temp_file.xlsx")
        
        if sheets:
            # 시트 선택
            sheet_names = excel_reader.get_sheet_names()
            selected_sheet = st.selectbox("시트 선택", sheet_names)
            
            if selected_sheet:
                df = excel_reader.get_sheet_data(selected_sheet)
                data_info = excel_reader.get_data_info(selected_sheet)
                
                # 탭 생성
                tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["📊 대시보드", "📈 차트 생성", "📋 데이터 보기", "📈 요약 통계", "🔍 데이터 분석", "🔬 고급 분석"])
                
                with tab1:
                    display_dashboard(df, chart_creator)
                
                with tab2:
                    st.header("차트 생성")
                    
                    # 차트 타입 선택
                    chart_type = st.selectbox(
                        "차트 타입 선택",
                        options=list(chart_creator.chart_types.keys()),
                        format_func=lambda x: chart_creator.chart_types[x]
                    )
                    
                    # 차트 옵션 가져오기
                    chart_options = chart_creator.get_chart_options(df)
                    
                    if chart_type in chart_options:
                        options = chart_options[chart_type]
                        
                        # 차트 생성
                        if chart_type == 'bar':
                            x_col = st.selectbox("X축", options.get('x', []))
                            y_col = st.selectbox("Y축", options.get('y', []))
                            color_col = st.selectbox("색상 구분", ['없음'] + options.get('color', []))
                            
                            if x_col and y_col:
                                color_col = None if color_col == '없음' else color_col
                                fig = chart_creator.create_bar_chart(df, x_col, y_col, color_col)
                                st.plotly_chart(fig, use_container_width=True)
                        
                        elif chart_type == 'line':
                            x_col = st.selectbox("X축", options.get('x', []))
                            y_col = st.selectbox("Y축", options.get('y', []))
                            color_col = st.selectbox("색상 구분", ['없음'] + options.get('color', []))
                            
                            if x_col and y_col:
                                color_col = None if color_col == '없음' else color_col
                                fig = chart_creator.create_line_chart(df, x_col, y_col, color_col)
                                st.plotly_chart(fig, use_container_width=True)
                        
                        elif chart_type == 'pie':
                            values_col = st.selectbox("값", options.get('values', []))
                            names_col = st.selectbox("이름", options.get('names', []))
                            
                            if values_col and names_col:
                                fig = chart_creator.create_pie_chart(df, values_col, names_col)
                                st.plotly_chart(fig, use_container_width=True)
                        
                        elif chart_type == 'scatter':
                            x_col = st.selectbox("X축", options.get('x', []))
                            y_col = st.selectbox("Y축", options.get('y', []))
                            color_col = st.selectbox("색상 구분", ['없음'] + options.get('color', []))
                            size_col = st.selectbox("크기", ['없음'] + options.get('size', []))
                            
                            if x_col and y_col:
                                color_col = None if color_col == '없음' else color_col
                                size_col = None if size_col == '없음' else size_col
                                fig = chart_creator.create_scatter_plot(df, x_col, y_col, color_col, size_col)
                                st.plotly_chart(fig, use_container_width=True)
                        
                        elif chart_type == 'histogram':
                            column = st.selectbox("컬럼", options.get('column', []))
                            bins = st.slider("구간 수", 5, 100, 30)
                            
                            if column:
                                fig = chart_creator.create_histogram(df, column, bins)
                                st.plotly_chart(fig, use_container_width=True)
                        
                        elif chart_type == 'box':
                            x_col = st.selectbox("X축", options.get('x', []))
                            y_col = st.selectbox("Y축", options.get('y', []))
                            
                            if x_col and y_col:
                                fig = chart_creator.create_box_plot(df, x_col, y_col)
                                st.plotly_chart(fig, use_container_width=True)
                        
                        elif chart_type == 'heatmap':
                            fig = chart_creator.create_heatmap(df)
                            st.plotly_chart(fig, use_container_width=True)
                        
                        elif chart_type == 'area':
                            x_col = st.selectbox("X축", options.get('x', []))
                            y_col = st.selectbox("Y축", options.get('y', []))
                            color_col = st.selectbox("색상 구분", ['없음'] + options.get('color', []))
                            
                            if x_col and y_col:
                                color_col = None if color_col == '없음' else color_col
                                fig = chart_creator.create_area_chart(df, x_col, y_col, color_col)
                                st.plotly_chart(fig, use_container_width=True)
                
                with tab3:
                    st.header("데이터 보기")
                    
                    # 데이터 필터링
                    st.subheader("데이터 필터링")
                    filter_col = st.selectbox("필터링할 컬럼", ['없음'] + df.columns.tolist())
                    
                    if filter_col != '없음':
                        unique_values = df[filter_col].unique()
                        selected_values = st.multiselect("값 선택", unique_values, default=unique_values)
                        filtered_df = df[df[filter_col].isin(selected_values)]
                    else:
                        filtered_df = df
                    
                    # 데이터 표시
                    st.dataframe(filtered_df, use_container_width=True)
                    
                    # 데이터 다운로드
                    csv = filtered_df.to_csv(index=False)
                    st.download_button(
                        label="필터링된 데이터 다운로드 (CSV)",
                        data=csv,
                        file_name=f"filtered_{selected_sheet}.csv",
                        mime="text/csv"
                    )
                
                with tab4:
                    st.header("요약 통계")
                    
                    # 수치형 데이터 요약
                    summary_stats = chart_creator.create_summary_stats(df)
                    if not summary_stats.empty:
                        st.subheader("수치형 데이터 요약")
                        st.dataframe(summary_stats, use_container_width=True)
                    
                    # 범주형 데이터 요약
                    categorical_cols = data_info.get('categorical_columns', [])
                    if categorical_cols:
                        st.subheader("범주형 데이터 요약")
                        for col in categorical_cols:
                            st.write(f"**{col}**")
                            value_counts = df[col].value_counts()
                            st.bar_chart(value_counts)
                
                with tab5:
                    st.header("데이터 분석")
                    
                    # 결측값 분석
                    missing_data = data_info.get('missing_values', {})
                    if any(missing_data.values()):
                        st.subheader("결측값 분석")
                        missing_df = pd.DataFrame(list(missing_data.items()), columns=['컬럼', '결측값 수'])
                        st.dataframe(missing_df)
                    
                    # 데이터 타입 정보
                    st.subheader("데이터 타입 정보")
                    dtype_info = pd.DataFrame(list(data_info.get('dtypes', {}).items()), 
                                           columns=['컬럼', '데이터 타입'])
                    st.dataframe(dtype_info)
                
                with tab6:
                    display_advanced_analysis(df, data_analyzer)
        
        # 임시 파일 삭제
        if os.path.exists("temp_file.xlsx"):
            os.remove("temp_file.xlsx")
    
    else:
        # 시작 화면
        st.markdown("""
        <div style="text-align: center; padding: 2rem;">
            <h2>📊 Excel Dashboard에 오신 것을 환영합니다!</h2>
            <p>왼쪽 사이드바에서 엑셀 파일을 업로드하여 데이터를 시각화해보세요.</p>
            <br>
            <h3>지원하는 기능:</h3>
            <ul style="text-align: left; display: inline-block;">
                <li>📈 자동 대시보드 생성 (여러 차트 동시 표시)</li>
                <li>📊 다양한 차트 타입 (막대그래프, 선그래프, 파이차트, 산점도 등)</li>
                <li>📋 데이터 필터링 및 탐색</li>
                <li>📊 요약 통계 및 분석</li>
                <li>🔬 고급 데이터 분석 (통계 검정, 군집 분석, PCA 등)</li>
                <li>💾 차트 및 데이터 다운로드</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main() 