"""
HWP to DOCX Converter – LibreOffice 없이 순수 Python으로 변환
- .hwpx (ZIP+XML 신형): zipfile + xml.etree 직접 파싱
- .hwp  (OLE 바이너리 구형): pyhwp(hwp5) 파싱
"""

import os
import re
import uuid
import logging
import tempfile
import zipfile
import asyncio
import xml.etree.ElementTree as ET
from io import BytesIO
from pathlib import Path
from contextlib import asynccontextmanager

from fastapi import FastAPI, File, UploadFile, HTTPException, Path as FPath
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from docx import Document

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

BASE   = Path(__file__).parent
OUTPUT = BASE / "outputs"
STATIC = BASE / "static"
OUTPUT.mkdir(exist_ok=True)

MAX_MB          = int(os.environ.get("MAX_MB", 30))
MAX_ZIP_MB      = int(os.environ.get("MAX_ZIP_MB", 200))   # 압축 해제 크기 상한
CORS_ORIGINS    = os.environ.get("CORS_ORIGINS", "*").split(",")
FILE_TTL_HOURS  = int(os.environ.get("FILE_TTL_HOURS", 1))  # 출력 파일 보존 시간

# ─────────────────────────────────────────
# 출력 파일 정리 (lifespan)
# ─────────────────────────────────────────

async def _cleanup_loop():
    """FILE_TTL_HOURS 경과한 outputs/ 파일을 주기적으로 삭제."""
    import time
    while True:
        await asyncio.sleep(600)  # 10분마다 실행
        cutoff = time.time() - FILE_TTL_HOURS * 3600
        for f in OUTPUT.iterdir():
            try:
                if f.is_file() and f.stat().st_mtime < cutoff:
                    f.unlink()
                    log.info(f"Cleaned up: {f.name}")
            except OSError:
                pass


@asynccontextmanager
async def lifespan(app: FastAPI):
    task = asyncio.create_task(_cleanup_loop())
    yield
    task.cancel()


app = FastAPI(title="HWP to DOCX Converter", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=CORS_ORIGINS,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)


# ─────────────────────────────────────────
# 경로 보안 헬퍼
# ─────────────────────────────────────────

_SAFE_JOB_ID = re.compile(r"^[0-9a-f]{32}$")


def safe_filename(name: str) -> str:
    """파일명에서 경로 구분자 및 위험 문자를 제거."""
    return os.path.basename(name).replace("..", "").strip() or "output"


def safe_output_path(job_id: str, filename: str) -> Path:
    """job_id와 filename을 검증 후 OUTPUT 하위 경로 반환."""
    if not _SAFE_JOB_ID.match(job_id):
        raise HTTPException(400, "잘못된 job_id 형식입니다.")
    name = safe_filename(filename)
    path = (OUTPUT / f"{job_id}_{name}").resolve()
    if path.parent.resolve() != OUTPUT.resolve():
        raise HTTPException(400, "잘못된 파일 경로입니다.")
    return path


# ─────────────────────────────────────────
# HWPX (ZIP + XML) 변환
# ─────────────────────────────────────────

def _iter_text_from_xml(xml_bytes: bytes):
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return

    def local(tag):
        return tag.split("}")[-1] if "}" in tag else tag

    for elem in root.iter():
        if local(elem.tag) == "P":
            has_content = False
            for child in elem.iter():
                tag = local(child.tag)
                if tag in ("CHAR", "T") and child.text:
                    yield (child.text, False, False)
                    has_content = True
            if has_content:
                yield ("\n", False, False)


def _hwpx_to_docx_bytes(hwp_bytes: bytes) -> bytes:
    doc = Document()
    max_bytes = MAX_ZIP_MB * 1024 * 1024

    with zipfile.ZipFile(BytesIO(hwp_bytes)) as zf:
        # Zip Bomb 방어: 압축 해제 전 총 크기 검증
        total_size = sum(i.file_size for i in zf.infolist())
        if total_size > max_bytes:
            raise ValueError(f"압축 해제 크기가 {MAX_ZIP_MB}MB를 초과합니다.")

        section_files = sorted(
            n for n in zf.namelist()
            if re.search(r"section\d+\.xml$", n, re.IGNORECASE)
        ) or [n for n in zf.namelist() if n.endswith(".xml")]

        for sec_file in section_files:
            para = doc.add_paragraph()
            for text, bold, italic in _iter_text_from_xml(zf.read(sec_file)):
                if text == "\n":
                    para = doc.add_paragraph()
                else:
                    run = para.add_run(text)
                    run.bold = bold
                    run.italic = italic

    out = BytesIO()
    doc.save(out)
    return out.getvalue()


# ─────────────────────────────────────────
# HWP5 (OLE 바이너리) 변환
# ─────────────────────────────────────────

def _hwp5_to_docx_bytes(hwp_bytes: bytes) -> bytes:
    try:
        from hwp5.hwp5txt import Hwp5File  # type: ignore
    except ImportError as e:
        raise RuntimeError("pyhwp 패키지가 필요합니다: pip install pyhwp") from e

    doc = Document()
    with tempfile.NamedTemporaryFile(suffix=".hwp", delete=False) as tmp:
        tmp.write(hwp_bytes)
        tmp_path = tmp.name

    try:
        hwpfile = Hwp5File(tmp_path)
        for section in hwpfile.text.sections:
            for item in section.models():
                if item.get("tagname") == "HWPTAG_PARA_TEXT":
                    payload = item.get("payload", b"")
                    try:
                        text = payload.decode("utf-16-le")
                        text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
                        doc.add_paragraph(text.rstrip("\r\n"))
                    except (UnicodeDecodeError, ValueError):
                        pass
    finally:
        os.unlink(tmp_path)

    out = BytesIO()
    doc.save(out)
    return out.getvalue()


# ─────────────────────────────────────────
# 공통 변환 진입점
# ─────────────────────────────────────────

def _convert_hwp_to_docx(filename: str, file_bytes: bytes) -> bytes:
    ext = os.path.splitext(filename)[1].lower()
    # 매직 바이트로 실제 포맷 판별
    is_zip = file_bytes[:4] == b"PK\x03\x04"
    is_ole = file_bytes[:8] == bytes.fromhex("d0cf11e0a1b11ae1")

    if ext == ".hwpx" or (ext == ".hwp" and is_zip):
        return _hwpx_to_docx_bytes(file_bytes)
    if ext == ".hwp" and is_ole:
        return _hwp5_to_docx_bytes(file_bytes)
    if ext == ".hwp":
        # 매직바이트 불명확 — OLE 경로 시도
        return _hwp5_to_docx_bytes(file_bytes)
    raise ValueError(f"지원하지 않는 파일 형식: {ext}")


# ─────────────────────────────────────────
# API 엔드포인트
# ─────────────────────────────────────────

@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    filename = safe_filename(file.filename or "upload.hwp")
    ext = os.path.splitext(filename)[1].lower()
    if ext not in (".hwp", ".hwpx"):
        raise HTTPException(400, "HWP 또는 HWPX 파일만 업로드 가능합니다.")

    content = await file.read()
    size_mb = len(content) / 1024 / 1024
    if size_mb > MAX_MB:
        raise HTTPException(413, f"파일 크기가 {MAX_MB}MB를 초과합니다. ({size_mb:.1f}MB)")

    log.info("Received %s (%.2f MB)", filename, size_mb)

    try:
        # CPU-bound 변환을 스레드풀로 오프로드 (이벤트 루프 블로킹 방지)
        docx_bytes = await asyncio.to_thread(_convert_hwp_to_docx, filename, content)
    except Exception as e:
        log.exception("변환 실패: %s", filename)
        raise HTTPException(500, "변환 중 오류가 발생했습니다. 파일을 확인해 주세요.")

    job_id   = uuid.uuid4().hex
    out_name = os.path.splitext(filename)[0] + ".docx"
    out_path = safe_output_path(job_id, out_name)
    out_path.write_bytes(docx_bytes)

    log.info("Converted → %s", out_path.name)
    return JSONResponse({
        "job_id": job_id,
        "filename": out_name,
        "download_url": f"/download/{job_id}/{out_name}",
    })


@app.get("/download/{job_id}/{filename}")
def download(
    job_id: str = FPath(..., pattern=r"^[0-9a-f]{32}$"),
    filename: str = FPath(...),
):
    path = safe_output_path(job_id, filename)
    if not path.exists():
        raise HTTPException(404, "파일을 찾을 수 없습니다.")
    safe_name = safe_filename(filename)
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=safe_name,
        headers={"Content-Disposition": f'attachment; filename="{safe_name}"'},
    )


@app.get("/health")
def health():
    return {"status": "ok", "engine": "pyhwp + python-docx (no LibreOffice)"}


app.mount("/", StaticFiles(directory=str(STATIC), html=True), name="static")
