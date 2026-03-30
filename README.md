# HWP → DOCX 변환기

HWP 파일을 DOCX로 변환하는 모바일 PWA + FastAPI 백엔드.

## 특징

- **LibreOffice 기반** — 고품질 변환, API key 불필요
- **로컬 처리** — 파일이 외부 서버로 전송되지 않음
- **모바일 PWA** — 홈화면 추가로 앱처럼 사용 가능
- **드래그 & 드롭** — 간편한 파일 업로드
- **최대 30MB** 지원

## 설치 및 실행

### 요구사항

- Python 3.9+
- LibreOffice (`sudo apt install libreoffice`)

### 실행

```bash
pip install -r requirements.txt
uvicorn main:app --host 0.0.0.0 --port 8000
```

브라우저에서 `http://localhost:8000` 접속

### 모바일에서 사용

같은 네트워크에서:
```
http://<서버IP>:8000
```

iOS Safari / Android Chrome → 홈화면에 추가 → 앱처럼 실행

## 배포 (Render.com 무료 플랜)

```bash
# render.yaml 설정 후
git push origin main
```

## API

```
POST /convert
  - body: multipart/form-data { file: .hwp }
  - response: { job_id, filename, download_url }

GET /download/{job_id}/{filename}
  - response: .docx file

GET /health
  - response: { status, soffice }
```

## 기술 스택

- **Backend**: FastAPI + Uvicorn
- **변환 엔진**: LibreOffice (soffice --headless)
- **Frontend**: Vanilla HTML/CSS/JS (PWA)
- **아이콘**: 자동 생성
