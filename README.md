# Doc Converter Windows

한글(HWP) 문서 변환 윈도우 서버 애플리케이션

## 기능
- HWP/HWPX 파일 처리
- 빨간색 텍스트 추출 및 동의어 문제 생성
  - 표에서 빨간색 텍스트를 추출하여 동의어 문제 생성
  - 문제와 답안을 자동으로 포맷팅

## 요구사항
- Windows OS
- Python 3.x
- 한글(HWP) 프로그램
- Flask
- pywin32

## 설치 방법
1. 저장소 클론
```bash
git clone https://github.com/[사용자명]/doc-converter-windows.git
```

2. 가상환경 생성 및 활성화
```bash
python -m venv doc-env
doc-env\Scripts\activate
```

3. 의존성 설치
```bash
pip install -r requirements.txt
```

## 사용 방법
1. 서버 실행
```bash
python run.py
```

2. API 엔드포인트
- POST `/api/synonym`: 동의어 문제 생성
  - 요청: multipart/form-data 형식으로 파일 업로드
  - 지원 파일 형식: .hwp, .hwpx
  - 응답: 생성된 동의어 문제가 포함된 HWP 파일

## 동의어 문제 생성 규칙
1. 표에서 빨간색으로 표시된 텍스트를 추출
2. 각 행의 첫 번째 단어를 문제로 사용
3. 나머지 단어들을 동의어로 처리
4. 문제와 답안을 자동으로 포맷팅하여 새로운 문서 생성

## 라이선스
MIT License
