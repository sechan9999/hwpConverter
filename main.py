import os, uuid, shutil, subprocess, logging
from pathlib import Path
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

app = FastAPI(title="HWP to DOCX Converter")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE    = Path(__file__).parent
UPLOAD  = BASE / "uploads"
OUTPUT  = BASE / "outputs"
STATIC  = BASE / "static"

UPLOAD.mkdir(exist_ok=True)
OUTPUT.mkdir(exist_ok=True)

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")
MAX_MB  = 30


@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".hwp"):
        raise HTTPException(400, "HWP 파일만 업로드 가능합니다.")

    content = await file.read()
    size_mb = len(content) / 1024 / 1024
    if size_mb > MAX_MB:
        raise HTTPException(413, f"파일 크기가 {MAX_MB}MB를 초과합니다. ({size_mb:.1f}MB)")

    job_id   = uuid.uuid4().hex
    job_dir  = UPLOAD / job_id
    job_dir.mkdir()
    hwp_path = job_dir / file.filename

    try:
        hwp_path.write_bytes(content)
        log.info(f"Received {file.filename} ({size_mb:.2f} MB), job={job_id}")

        result = subprocess.run(
            [SOFFICE,
             "--headless",
             "--norestore",
             "--convert-to", "docx",
             "--outdir", str(OUTPUT),
             str(hwp_path)],
            capture_output=True,
            text=True,
            timeout=120
        )

        log.info(f"soffice stdout: {result.stdout}")
        if result.stderr:
            log.warning(f"soffice stderr: {result.stderr}")

        stem   = hwp_path.stem
        docx   = OUTPUT / f"{stem}.docx"

        if not docx.exists():
            # LibreOffice sometimes appends job info — search
            matches = list(OUTPUT.glob(f"{stem}*.docx"))
            if matches:
                docx = matches[0]
            else:
                raise HTTPException(500, f"변환 실패: {result.stderr or result.stdout}")

        out_name = f"{stem}.docx"
        final    = OUTPUT / f"{job_id}_{out_name}"
        docx.rename(final)

        log.info(f"Converted → {final.name}")
        return JSONResponse({"job_id": job_id, "filename": out_name, "download_url": f"/download/{job_id}/{out_name}"})

    except subprocess.TimeoutExpired:
        raise HTTPException(504, "변환 시간이 초과되었습니다. 파일을 확인해 주세요.")
    except HTTPException:
        raise
    except Exception as e:
        log.exception("Unexpected error")
        raise HTTPException(500, str(e))
    finally:
        shutil.rmtree(job_dir, ignore_errors=True)


@app.get("/download/{job_id}/{filename}")
def download(job_id: str, filename: str):
    path = OUTPUT / f"{job_id}_{filename}"
    if not path.exists():
        raise HTTPException(404, "파일을 찾을 수 없습니다.")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=filename,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


@app.get("/health")
def health():
    return {"status": "ok", "soffice": SOFFICE}


app.mount("/", StaticFiles(directory=str(STATIC), html=True), name="static")
