# Excel Dashboard Python

엑셀 파일을 업로드하여 데이터를 시각화하는 Python 프로젝트입니다.

## 기능

- 엑셀 파일 업로드 및 읽기
- **자동 대시보드 생성** (여러 차트 동시 표시)
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
- 영역차트 (Area Chart)

## 자동 대시보드 기능

엑셀 파일을 업로드하면 자동으로 다음과 같은 유의미한 차트들이 생성됩니다:

1. **상관관계 히트맵**: 수치형 데이터 간의 상관관계 분석
2. **히스토그램**: 주요 수치형 데이터의 분포 분석
3. **박스플롯**: 범주별 수치 데이터 분포 및 이상치 탐지
4. **막대그래프**: 범주별 수치 데이터 비교
5. **파이차트**: 범주별 비율 분석
6. **산점도**: 두 수치형 변수 간의 관계 분석
7. **월별/분기별 집계**: 시간별 데이터 집계 및 트렌드 분석
8. **상위 10개 분석**: 가장 높은 값을 가진 항목들 분석
9. **구간별 분포 분석**: 데이터를 구간별로 분류하여 패턴 분석
10. **성장률 분석**: 월별 성장률 변화 추이
11. **평균 vs 중앙값 비교**: 데이터의 중심 경향성 분석
12. **표준편차 분석**: 변수별 분산 정도 비교
13. **범주별 조합 분석**: 두 범주형 변수의 조합별 평균 분석
14. **분위수 분석**: 데이터의 분포 위치 분석
15. **이상치 탐지**: 정상 데이터와 이상치 구분 분석

## 샘플 데이터

프로젝트에는 다음과 같은 샘플 데이터가 포함되어 있습니다:
- **Sales**: 판매 데이터 (제품, 지역, 매출 등)
- **Employees**: 직원 데이터 (부서, 급여, 성과 등)
- **Weather**: 날씨 데이터 (온도, 습도, 강수량 등)
- **Stocks**: 주식 데이터 (가격, 거래량, 변동률 등)

## 라이선스

MIT License 