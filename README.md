# Doc Converter Windows

한글(HWP)과 워드(DOCX) 문서 변환 윈도우 서버 애플리케이션

## 기능
- HWP/HWPX 파일을 DOCX로 변환
- DOCX 파일을 HWP로 변환
- 빨간색 텍스트 추출 및 동의어 문제 생성

## 요구사항
- Windows OS
- Python 3.x
- 한글(HWP) 프로그램

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
- POST `/api/synonym`: 문서 변환 및 동의어 문제 생성

## 라이선스
MIT License
