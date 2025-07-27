import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from utils.excel_reader import ExcelReader, create_sample_excel
from utils.chart_creator import ChartCreator
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
</style>
""", unsafe_allow_html=True)

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
                
                # 데이터 정보 표시
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("행 수", data_info.get('shape', [0, 0])[0])
                with col2:
                    st.metric("열 수", data_info.get('shape', [0, 0])[1])
                with col3:
                    st.metric("수치형 컬럼", len(data_info.get('numeric_columns', [])))
                with col4:
                    st.metric("범주형 컬럼", len(data_info.get('categorical_columns', [])))
                
                # 탭 생성
                tab1, tab2, tab3, tab4 = st.tabs(["📊 차트 생성", "📋 데이터 보기", "📈 요약 통계", "🔍 데이터 분석"])
                
                with tab1:
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
                
                with tab2:
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
                
                with tab3:
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
                
                with tab4:
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
                <li>📈 다양한 차트 타입 (막대그래프, 선그래프, 파이차트, 산점도 등)</li>
                <li>📋 데이터 필터링 및 탐색</li>
                <li>📊 요약 통계 및 분석</li>
                <li>💾 차트 및 데이터 다운로드</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main() 