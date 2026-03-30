"""
HWP to DOCX Converter – LibreOffice 없이 순수 Python으로 변환
- .hwpx (ZIP+XML 신형): zipfile + xml.etree 직접 파싱
- .hwp  (OLE 바이너리 구형): pyhwp(hwp5) 파싱
"""

import os, uuid, logging, re, tempfile, zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
from pathlib import Path

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from docx import Document

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

app = FastAPI(title="HWP to DOCX Converter")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE   = Path(__file__).parent
OUTPUT = BASE / "outputs"
STATIC = BASE / "static"
OUTPUT.mkdir(exist_ok=True)

MAX_MB = 30


# ─────────────────────────────────────────
# HWPX (ZIP + XML) 변환
# ─────────────────────────────────────────

def _iter_text_from_xml(xml_bytes: bytes):
    """HWPX section XML에서 (텍스트, bold, italic) 토큰 생성."""
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
    with zipfile.ZipFile(BytesIO(hwp_bytes)) as zf:
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

def convert_hwp_to_docx(filename: str, file_bytes: bytes) -> bytes:
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".hwpx":
        return _hwpx_to_docx_bytes(file_bytes)
    if ext == ".hwp":
        # 매직 바이트로 실제 포맷 판별 (ZIP이면 HWPX)
        if file_bytes[:4] == b"PK\x03\x04":
            return _hwpx_to_docx_bytes(file_bytes)
        return _hwp5_to_docx_bytes(file_bytes)
    raise ValueError(f"지원하지 않는 파일 형식: {ext}")


# ─────────────────────────────────────────
# API 엔드포인트
# ─────────────────────────────────────────

@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    filename = file.filename or ""
    ext = os.path.splitext(filename)[1].lower()
    if ext not in (".hwp", ".hwpx"):
        raise HTTPException(400, "HWP 또는 HWPX 파일만 업로드 가능합니다.")

    content = await file.read()
    size_mb = len(content) / 1024 / 1024
    if size_mb > MAX_MB:
        raise HTTPException(413, f"파일 크기가 {MAX_MB}MB를 초과합니다. ({size_mb:.1f}MB)")

    log.info(f"Received {filename} ({size_mb:.2f} MB)")

    try:
        docx_bytes = convert_hwp_to_docx(filename, content)
    except Exception as e:
        log.exception("변환 실패")
        raise HTTPException(500, f"변환 실패: {e}")

    job_id   = uuid.uuid4().hex
    out_name = os.path.splitext(filename)[0] + ".docx"
    out_path = OUTPUT / f"{job_id}_{out_name}"
    out_path.write_bytes(docx_bytes)

    log.info(f"Converted → {out_path.name}")
    return JSONResponse({
        "job_id": job_id,
        "filename": out_name,
        "download_url": f"/download/{job_id}/{out_name}",
    })


@app.get("/download/{job_id}/{filename}")
def download(job_id: str, filename: str):
    path = OUTPUT / f"{job_id}_{filename}"
    if not path.exists():
        raise HTTPException(404, "파일을 찾을 수 없습니다.")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=filename,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.get("/health")
def health():
    return {"status": "ok", "engine": "pyhwp + python-docx (no LibreOffice)"}


app.mount("/", StaticFiles(directory=str(STATIC), html=True), name="static")
