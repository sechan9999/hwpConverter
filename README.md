# HWP → DOCX 변환기

HWP/HWPX 파일을 DOCX로 변환하는 모바일 PWA + FastAPI 백엔드.
**LibreOffice 불필요** — 순수 Python으로 동작합니다.

## 특징

- **LibreOffice 불필요** — `pyhwp` + `python-docx` 순수 Python 변환
- **HWP / HWPX 모두 지원** — 구형 OLE 바이너리(.hwp) & 신형 ZIP+XML(.hwpx)
- **로컬 처리** — 파일이 외부 서버로 전송되지 않음
- **모바일 PWA** — 홈화면 추가로 앱처럼 사용 가능
- **드래그 & 드롭** — 간편한 파일 업로드
- **최대 30MB** 지원

## 변환 방식

| 파일 | 포맷 감지 | 변환 방법 |
|------|----------|----------|
| `.hwpx` | 확장자 | ZIP 언패킹 → section XML 파싱 → python-docx |
| `.hwp` (신형) | 매직바이트 `PK` | ZIP 언패킹 → section XML 파싱 → python-docx |
| `.hwp` (구형 OLE) | 매직바이트 `D0CF` | pyhwp `HWPTAG_PARA_TEXT` 추출 → python-docx |

## 설치 및 실행

### 요구사항

- Python 3.9+
- LibreOffice **불필요**

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

> Render 빌드 시 LibreOffice 설치 단계가 제거되어 **빌드 속도가 빨라집니다.**

## API

```
POST /convert
  - body: multipart/form-data { file: .hwp 또는 .hwpx }
  - response: { job_id, filename, download_url }

GET /download/{job_id}/{filename}
  - response: .docx file

GET /health
  - response: { status, engine }
```

## 기술 스택

- **Backend**: FastAPI + Uvicorn
- **변환 엔진**: pyhwp (HWP5 OLE 파싱) + python-docx (DOCX 생성)
- **HWPX 파싱**: zipfile + xml.etree (표준 라이브러리)
- **Frontend**: Vanilla HTML/CSS/JS (PWA)
