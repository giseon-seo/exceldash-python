# Excel Dashboard Python

엑셀 파일을 업로드하여 데이터를 시각화하는 Python 프로젝트입니다.

## 기능

- 엑셀 파일 업로드 및 읽기
- 다양한 차트 타입 지원 (막대그래프, 선그래프, 파이차트, 히트맵 등)
- 인터랙티브 대시보드
- 데이터 필터링 및 분석

## 설치 방법

1. 가상환경 생성 및 활성화:
```bash
python -m venv venv
source venv/bin/activate  # macOS/Linux
# 또는
venv\Scripts\activate  # Windows
```

2. 필요한 패키지 설치:
```bash
pip install -r requirements.txt
```

## 사용 방법

### Streamlit 앱 실행
```bash
streamlit run app.py
```

### Dash 앱 실행
```bash
python dash_app.py
```

## 프로젝트 구조

```
exceldash-python/
├── app.py                 # Streamlit 메인 앱
├── dash_app.py           # Dash 대시보드 앱
├── utils/
│   ├── __init__.py
│   ├── excel_reader.py   # 엑셀 파일 읽기 유틸리티
│   └── chart_creator.py  # 차트 생성 유틸리티
├── data/                 # 샘플 데이터 및 업로드된 파일
├── requirements.txt      # Python 패키지 의존성
└── README.md           # 프로젝트 설명서
```

## 지원하는 차트 타입

- 막대그래프 (Bar Chart)
- 선그래프 (Line Chart)
- 파이차트 (Pie Chart)
- 히트맵 (Heatmap)
- 산점도 (Scatter Plot)
- 박스플롯 (Box Plot)
- 히스토그램 (Histogram)

## 라이선스

MIT License 