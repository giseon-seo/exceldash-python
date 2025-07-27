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
     Output('chart-controls', 'children'),
     Output('chart-controls', 'style')],
    [Input('upload-data', 'contents')],
    [State('upload-data', 'filename')]
)
def update_output(contents, filename):
    if contents is None:
        return "파일을 업로드하세요", "", {'display': 'none'}, "", {'display': 'none'}
    
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
            
            # 차트 컨트롤
            chart_creator = ChartCreator()
            chart_options = chart_creator.get_chart_options(df)
            
            controls_content = dbc.Card([
                dbc.CardHeader("📈 차트 생성"),
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
            
            return f"✅ {filename} 업로드 완료", info_content, {'display': 'block'}, controls_content, {'display': 'block'}
        
        else:
            return "❌ 파일 읽기 오류", "", {'display': 'none'}, "", {'display': 'none'}
    
    except Exception as e:
        return f"❌ 오류: {str(e)}", "", {'display': 'none'}, "", {'display': 'none'}

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