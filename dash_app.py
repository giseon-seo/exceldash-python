import dash
from dash import dcc, html, Input, Output, State, callback_context
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from utils.excel_reader import ExcelReader, create_sample_excel
from utils.chart_creator import ChartCreator
import base64
import io
import json

# Dash 앱 초기화
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
app.title = "Excel Dashboard"

def create_dashboard_charts(df, chart_creator):
    """대시보드용 차트들을 생성"""
    charts = []
    
    # 수치형 컬럼들
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
    date_cols = df.select_dtypes(include=['datetime64']).columns.tolist()
    
    if len(numeric_cols) >= 2:
        # 1. 상관관계 히트맵
        if len(numeric_cols) > 1:
            fig_heatmap = chart_creator.create_heatmap(df, numeric_cols, "상관관계 히트맵")
            charts.append(("상관관계 히트맵", fig_heatmap))
        
        # 2. 수치형 데이터 분포 (히스토그램)
        if numeric_cols:
            selected_numeric = numeric_cols[0] if numeric_cols else None
            if selected_numeric:
                fig_hist = chart_creator.create_histogram(df, selected_numeric, 20, f"{selected_numeric} 분포")
                charts.append((f"{selected_numeric} 분포", fig_hist))
        
        # 3. 박스플롯 (범주형 컬럼이 있는 경우)
        if categorical_cols and numeric_cols:
            cat_col = categorical_cols[0]
            num_col = numeric_cols[0]
            fig_box = chart_creator.create_box_plot(df, cat_col, num_col, f"{cat_col}별 {num_col} 분포")
            charts.append((f"{cat_col}별 {num_col} 분포", fig_box))
    
    # 4. 막대그래프 (범주형 데이터)
    if categorical_cols and numeric_cols:
        cat_col = categorical_cols[0]
        num_col = numeric_cols[0]
        fig_bar = chart_creator.create_bar_chart(df, cat_col, num_col, title=f"{cat_col}별 {num_col}")
        charts.append((f"{cat_col}별 {num_col}", fig_bar))
    
    # 5. 파이차트 (범주형 데이터)
    if categorical_cols and numeric_cols:
        cat_col = categorical_cols[0]
        num_col = numeric_cols[0]
        # 파이차트는 값의 합계를 사용
        pie_data = df.groupby(cat_col)[num_col].sum().reset_index()
        fig_pie = chart_creator.create_pie_chart(pie_data, num_col, cat_col, f"{cat_col}별 {num_col} 비율")
        charts.append((f"{cat_col}별 {num_col} 비율", fig_pie))
    
    # 6. 산점도 (두 수치형 컬럼)
    if len(numeric_cols) >= 2:
        x_col = numeric_cols[0]
        y_col = numeric_cols[1]
        fig_scatter = chart_creator.create_scatter_plot(df, x_col, y_col, title=f"{x_col} vs {y_col}")
        charts.append((f"{x_col} vs {y_col}", fig_scatter))
    
    # 7. 월별/분기별 집계 차트 (날짜 데이터가 있는 경우)
    if date_cols and numeric_cols:
        date_col = date_cols[0]
        num_col = numeric_cols[0]
        
        # 월별 집계
        df_copy = df.copy()
        df_copy['Month'] = df_copy[date_col].dt.to_period('M')
        monthly_data = df_copy.groupby('Month')[num_col].agg(['sum', 'mean', 'count']).reset_index()
        monthly_data['Month'] = monthly_data['Month'].astype(str)
        
        fig_monthly = chart_creator.create_bar_chart(monthly_data, 'Month', 'sum', title=f"월별 {num_col} 합계")
        charts.append((f"월별 {num_col} 합계", fig_monthly))
        
        # 분기별 집계
        df_copy['Quarter'] = df_copy[date_col].dt.to_period('Q')
        quarterly_data = df_copy.groupby('Quarter')[num_col].agg(['sum', 'mean']).reset_index()
        quarterly_data['Quarter'] = quarterly_data['Quarter'].astype(str)
        
        fig_quarterly = chart_creator.create_bar_chart(quarterly_data, 'Quarter', 'sum', title=f"분기별 {num_col} 합계")
        charts.append((f"분기별 {num_col} 합계", fig_quarterly))
    
    # 8. 상위/하위 분석
    if categorical_cols and numeric_cols:
        cat_col = categorical_cols[0]
        num_col = numeric_cols[0]
        
        # 상위 10개 분석
        top_data = df.groupby(cat_col)[num_col].sum().sort_values(ascending=False).head(10).reset_index()
        fig_top = chart_creator.create_bar_chart(top_data, cat_col, num_col, title=f"상위 10개 {cat_col}별 {num_col}")
        charts.append((f"상위 10개 {cat_col}별 {num_col}", fig_top))
    
    # 9. 데이터 분포 분석 (수치형 데이터가 2개 이상인 경우)
    if len(numeric_cols) >= 2:
        # 첫 번째 수치형 컬럼의 분포를 다른 컬럼으로 그룹화
        x_col = numeric_cols[0]
        y_col = numeric_cols[1]
        
        # 구간별 분류
        df_copy = df.copy()
        df_copy['Range'] = pd.cut(df_copy[x_col], bins=5, labels=['매우 낮음', '낮음', '보통', '높음', '매우 높음'])
        range_data = df_copy.groupby('Range')[y_col].mean().reset_index()
        
        fig_range = chart_creator.create_bar_chart(range_data, 'Range', y_col, title=f"{x_col} 구간별 {y_col} 평균")
        charts.append((f"{x_col} 구간별 {y_col} 평균", fig_range))
    
    # 10. 성장률 분석 (날짜 데이터가 있는 경우)
    if date_cols and numeric_cols:
        date_col = date_cols[0]
        num_col = numeric_cols[0]
        
        # 월별 성장률 계산
        df_copy = df.copy()
        df_copy['Month'] = df_copy[date_col].dt.to_period('M')
        monthly_sum = df_copy.groupby('Month')[num_col].sum().reset_index()
        monthly_sum['Growth_Rate'] = monthly_sum[num_col].pct_change() * 100
        
        # 성장률이 있는 데이터만 필터링
        growth_data = monthly_sum[monthly_sum['Growth_Rate'].notna()]
        if not growth_data.empty:
            growth_data['Month'] = growth_data['Month'].astype(str)
            fig_growth = chart_creator.create_bar_chart(growth_data, 'Month', 'Growth_Rate', title=f"월별 {num_col} 성장률 (%)")
            charts.append((f"월별 {num_col} 성장률 (%)", fig_growth))
    
    # 11. 평균 vs 중앙값 비교 (수치형 데이터)
    if numeric_cols:
        num_col = numeric_cols[0]
        mean_val = df[num_col].mean()
        median_val = df[num_col].median()
        
        comparison_data = pd.DataFrame({
            '통계': ['평균', '중앙값'],
            '값': [mean_val, median_val]
        })
        
        fig_comparison = chart_creator.create_bar_chart(comparison_data, '통계', '값', title=f"{num_col} 평균 vs 중앙값")
        charts.append((f"{num_col} 평균 vs 중앙값", fig_comparison))
    
    # 12. 표준편차 분석 (수치형 데이터가 2개 이상인 경우)
    if len(numeric_cols) >= 2:
        std_data = df[numeric_cols].std().reset_index()
        std_data.columns = ['변수', '표준편차']
        
        fig_std = chart_creator.create_bar_chart(std_data, '변수', '표준편차', title="변수별 표준편차")
        charts.append(("변수별 표준편차", fig_std))
    
    # 13. 범주별 평균 비교 (범주형 데이터가 2개 이상인 경우)
    if len(categorical_cols) >= 2 and numeric_cols:
        cat1 = categorical_cols[0]
        cat2 = categorical_cols[1]
        num_col = numeric_cols[0]
        
        # 두 범주형 변수의 조합별 평균
        combo_data = df.groupby([cat1, cat2])[num_col].mean().reset_index()
        combo_data['조합'] = combo_data[cat1] + ' - ' + combo_data[cat2]
        
        fig_combo = chart_creator.create_bar_chart(combo_data, '조합', num_col, title=f"{cat1} x {cat2} 조합별 {num_col} 평균")
        charts.append((f"{cat1} x {cat2} 조합별 {num_col} 평균", fig_combo))
    
    # 14. 분위수 분석 (수치형 데이터)
    if numeric_cols:
        num_col = numeric_cols[0]
        quantiles = df[num_col].quantile([0.25, 0.5, 0.75, 0.9, 0.95, 0.99]).reset_index()
        quantiles.columns = ['분위수', '값']
        quantiles['분위수'] = quantiles['분위수'] * 100
        
        fig_quantile = chart_creator.create_bar_chart(quantiles, '분위수', '값', title=f"{num_col} 분위수 분석")
        charts.append((f"{num_col} 분위수 분석", fig_quantile))
    
    # 15. 이상치 탐지 (박스플롯 기반)
    if numeric_cols:
        num_col = numeric_cols[0]
        Q1 = df[num_col].quantile(0.25)
        Q3 = df[num_col].quantile(0.75)
        IQR = Q3 - Q1
        lower_bound = Q1 - 1.5 * IQR
        upper_bound = Q3 + 1.5 * IQR
        
        outliers = df[(df[num_col] < lower_bound) | (df[num_col] > upper_bound)]
        normal_data = df[(df[num_col] >= lower_bound) & (df[num_col] <= upper_bound)]
        
        outlier_summary = pd.DataFrame({
            '구분': ['정상 데이터', '이상치'],
            '개수': [len(normal_data), len(outliers)],
            '비율': [len(normal_data)/len(df)*100, len(outliers)/len(df)*100]
        })
        
        fig_outlier = chart_creator.create_bar_chart(outlier_summary, '구분', '개수', title=f"{num_col} 이상치 분석")
        charts.append((f"{num_col} 이상치 분석", fig_outlier))
    
    return charts

# 레이아웃
app.layout = dbc.Container([
    dbc.Row([
        dbc.Col([
            html.H1("📊 Excel Dashboard", className="text-center mb-4"),
            html.Hr()
        ])
    ]),
    
    dbc.Row([
        dbc.Col([
            dbc.Card([
                dbc.CardHeader("📁 파일 업로드"),
                dbc.CardBody([
                    dcc.Upload(
                        id='upload-data',
                        children=html.Div([
                            '드래그 앤 드롭 또는 ',
                            html.A('파일 선택')
                        ]),
                        style={
                            'width': '100%',
                            'height': '60px',
                            'lineHeight': '60px',
                            'borderWidth': '1px',
                            'borderStyle': 'dashed',
                            'borderRadius': '5px',
                            'textAlign': 'center',
                            'margin': '10px'
                        },
                        multiple=False
                    ),
                    html.Div(id='upload-status')
                ])
            ])
        ], width=12)
    ]),
    
    dbc.Row([
        dbc.Col([
            html.Div(id='data-info', style={'display': 'none'})
        ])
    ]),
    
    dbc.Row([
        dbc.Col([
            html.Div(id='dashboard-output', style={'display': 'none'})
        ])
    ]),
    
    dbc.Row([
        dbc.Col([
            html.Div(id='chart-controls', style={'display': 'none'})
        ])
    ]),
    
    dbc.Row([
        dbc.Col([
            html.Div(id='chart-output')
        ])
    ]),
    
    dbc.Row([
        dbc.Col([
            html.Div(id='data-table', style={'display': 'none'})
        ])
    ])
], fluid=True)

# 파일 업로드 콜백
@app.callback(
    [Output('upload-status', 'children'),
     Output('data-info', 'children'),
     Output('data-info', 'style'),
     Output('dashboard-output', 'children'),
     Output('dashboard-output', 'style'),
     Output('chart-controls', 'children'),
     Output('chart-controls', 'style')],
    [Input('upload-data', 'contents')],
    [State('upload-data', 'filename')]
)
def update_output(contents, filename):
    if contents is None:
        return "파일을 업로드하세요", "", {'display': 'none'}, "", {'display': 'none'}, "", {'display': 'none'}
    
    # 파일 내용 디코딩
    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    
    try:
        # 엑셀 파일 읽기
        excel_reader = ExcelReader()
        
        # 임시 파일로 저장
        with open("temp_dash_file.xlsx", "wb") as f:
            f.write(decoded)
        
        sheets = excel_reader.read_excel("temp_dash_file.xlsx")
        
        if sheets:
            sheet_names = excel_reader.get_sheet_names()
            selected_sheet = sheet_names[0]  # 첫 번째 시트 선택
            df = excel_reader.get_sheet_data(selected_sheet)
            data_info = excel_reader.get_data_info(selected_sheet)
            
            # 데이터 정보 표시
            info_content = dbc.Card([
                dbc.CardHeader("📊 데이터 정보"),
                dbc.CardBody([
                    dbc.Row([
                        dbc.Col([
                            dbc.Card([
                                dbc.CardBody([
                                    html.H4(data_info.get('shape', [0, 0])[0], className="text-center"),
                                    html.P("행 수", className="text-center text-muted")
                                ])
                            ])
                        ], width=3),
                        dbc.Col([
                            dbc.Card([
                                dbc.CardBody([
                                    html.H4(data_info.get('shape', [0, 0])[1], className="text-center"),
                                    html.P("열 수", className="text-center text-muted")
                                ])
                            ])
                        ], width=3),
                        dbc.Col([
                            dbc.Card([
                                dbc.CardBody([
                                    html.H4(len(data_info.get('numeric_columns', [])), className="text-center"),
                                    html.P("수치형 컬럼", className="text-center text-muted")
                                ])
                            ])
                        ], width=3),
                        dbc.Col([
                            dbc.Card([
                                dbc.CardBody([
                                    html.H4(len(data_info.get('categorical_columns', [])), className="text-center"),
                                    html.P("범주형 컬럼", className="text-center text-muted")
                                ])
                            ])
                        ], width=3)
                    ])
                ])
            ])
            
            # 대시보드 생성
            chart_creator = ChartCreator()
            dashboard_charts = create_dashboard_charts(df, chart_creator)
            
            dashboard_content = []
            if dashboard_charts:
                dashboard_content.append(html.H3("📊 자동 생성된 대시보드", className="mb-4"))
                
                for i in range(0, len(dashboard_charts), 2):
                    row = dbc.Row([
                        dbc.Col([
                            html.H5(dashboard_charts[i][0], className="text-center"),
                            dcc.Graph(figure=dashboard_charts[i][1])
                        ], width=6)
                    ])
                    
                    if i + 1 < len(dashboard_charts):
                        row.children.append(
                            dbc.Col([
                                html.H5(dashboard_charts[i + 1][0], className="text-center"),
                                dcc.Graph(figure=dashboard_charts[i + 1][1])
                            ], width=6)
                        )
                    
                    dashboard_content.append(row)
            else:
                dashboard_content.append(html.Div("대시보드를 생성할 수 있는 충분한 데이터가 없습니다.", className="text-center text-muted"))
            
            # 차트 컨트롤
            chart_options = chart_creator.get_chart_options(df)
            
            controls_content = dbc.Card([
                dbc.CardHeader("📈 개별 차트 생성"),
                dbc.CardBody([
                    dbc.Row([
                        dbc.Col([
                            html.Label("차트 타입"),
                            dcc.Dropdown(
                                id='chart-type-dropdown',
                                options=[{'label': v, 'value': k} for k, v in chart_creator.chart_types.items()],
                                value='bar'
                            )
                        ], width=4),
                        dbc.Col([
                            html.Label("시트 선택"),
                            dcc.Dropdown(
                                id='sheet-dropdown',
                                options=[{'label': name, 'value': name} for name in sheet_names],
                                value=selected_sheet
                            )
                        ], width=4),
                        dbc.Col([
                            html.Label("X축"),
                            dcc.Dropdown(id='x-axis-dropdown', options=[])
                        ], width=4)
                    ], className="mb-3"),
                    dbc.Row([
                        dbc.Col([
                            html.Label("Y축"),
                            dcc.Dropdown(id='y-axis-dropdown', options=[])
                        ], width=4),
                        dbc.Col([
                            html.Label("색상 구분"),
                            dcc.Dropdown(id='color-dropdown', options=[])
                        ], width=4),
                        dbc.Col([
                            html.Label("크기"),
                            dcc.Dropdown(id='size-dropdown', options=[])
                        ], width=4)
                    ])
                ])
            ])
            
            return f"✅ {filename} 업로드 완료", info_content, {'display': 'block'}, dashboard_content, {'display': 'block'}, controls_content, {'display': 'block'}
        
        else:
            return "❌ 파일 읽기 오류", "", {'display': 'none'}, "", {'display': 'none'}, "", {'display': 'none'}
    
    except Exception as e:
        return f"❌ 오류: {str(e)}", "", {'display': 'none'}, "", {'display': 'none'}, "", {'display': 'none'}

# 차트 타입 변경 콜백
@app.callback(
    [Output('x-axis-dropdown', 'options'),
     Output('y-axis-dropdown', 'options'),
     Output('color-dropdown', 'options'),
     Output('size-dropdown', 'options')],
    [Input('chart-type-dropdown', 'value'),
     Input('sheet-dropdown', 'value')]
)
def update_chart_options(chart_type, sheet_name):
    if not chart_type or not sheet_name:
        return [], [], [], []
    
    try:
        excel_reader = ExcelReader()
        sheets = excel_reader.read_excel("temp_dash_file.xlsx")
        df = excel_reader.get_sheet_data(sheet_name)
        chart_creator = ChartCreator()
        chart_options = chart_creator.get_chart_options(df)
        
        if chart_type in chart_options:
            options = chart_options[chart_type]
            
            x_options = [{'label': col, 'value': col} for col in options.get('x', [])]
            y_options = [{'label': col, 'value': col} for col in options.get('y', [])]
            color_options = [{'label': '없음', 'value': 'none'}] + [{'label': col, 'value': col} for col in options.get('color', [])]
            size_options = [{'label': '없음', 'value': 'none'}] + [{'label': col, 'value': col} for col in options.get('size', [])]
            
            return x_options, y_options, color_options, size_options
        
        return [], [], [], []
    
    except:
        return [], [], [], []

# 차트 생성 콜백
@app.callback(
    Output('chart-output', 'children'),
    [Input('chart-type-dropdown', 'value'),
     Input('sheet-dropdown', 'value'),
     Input('x-axis-dropdown', 'value'),
     Input('y-axis-dropdown', 'value'),
     Input('color-dropdown', 'value'),
     Input('size-dropdown', 'value')]
)
def update_chart(chart_type, sheet_name, x_col, y_col, color_col, size_col):
    if not all([chart_type, sheet_name, x_col, y_col]):
        return ""
    
    try:
        excel_reader = ExcelReader()
        sheets = excel_reader.read_excel("temp_dash_file.xlsx")
        df = excel_reader.get_sheet_data(sheet_name)
        chart_creator = ChartCreator()
        
        # 색상 및 크기 컬럼 처리
        color_col = None if color_col == 'none' else color_col
        size_col = None if size_col == 'none' else size_col
        
        if chart_type == 'bar':
            fig = chart_creator.create_bar_chart(df, x_col, y_col, color_col)
        elif chart_type == 'line':
            fig = chart_creator.create_line_chart(df, x_col, y_col, color_col)
        elif chart_type == 'scatter':
            fig = chart_creator.create_scatter_plot(df, x_col, y_col, color_col, size_col)
        elif chart_type == 'area':
            fig = chart_creator.create_area_chart(df, x_col, y_col, color_col)
        elif chart_type == 'box':
            fig = chart_creator.create_box_plot(df, x_col, y_col)
        elif chart_type == 'histogram':
            fig = chart_creator.create_histogram(df, x_col)
        elif chart_type == 'heatmap':
            fig = chart_creator.create_heatmap(df)
        else:
            return ""
        
        return dcc.Graph(figure=fig)
    
    except Exception as e:
        return html.Div(f"차트 생성 오류: {str(e)}", style={'color': 'red'})

if __name__ == '__main__':
    app.run_server(debug=True, host='0.0.0.0', port=8050) 