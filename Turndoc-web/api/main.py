import os
import shutil
import uuid
import zipfile
from pathlib import Path

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from converter import TurnDocConverter


# ============================================================
# TurnDoc Web API
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

# main.py:
# Turndoc-web/api/main.py
#
# 上一層：
# Turndoc-web/
#
# 因此 index.html 位於：
# Turndoc-web/index.html
PROJECT_DIR = BASE_DIR.parent

UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# FastAPI
# ============================================================

app = FastAPI(
    title="TurnDoc API",
    description="TurnDoc 文件轉換 Web API",
    version="1.0.0"
)


# ============================================================
# CORS
# ============================================================

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ============================================================
# 靜態網站
# ============================================================

ASSETS_DIR = PROJECT_DIR / "assets"

if ASSETS_DIR.exists() and ASSETS_DIR.is_dir():
    app.mount(
        "/assets",
        StaticFiles(directory=str(ASSETS_DIR)),
        name="assets"
    )


# ============================================================
# 首頁
# ============================================================

@app.get("/")
def root():
    index_file = PROJECT_DIR / "index.html"

    if not index_file.exists():
        raise HTTPException(
            status_code=404,
            detail="找不到首頁 index.html"
        )

    return FileResponse(
        path=str(index_file),
        media_type="text/html"
    )


# ============================================================
# 支援工具
# ============================================================

SUPPORTED_TOOLS = {
    "word-to-pdf",
    "pdf-to-word",
    "ppt-to-pdf",
    "excel-to-pdf",
    "pdf-to-ppt",
    "pdf-to-excel",
    "merge",
    "compress",
    "pdf-to-image",
    "image-to-pdf",
    "image-convert",
    "image-compress",
    "gif-split",
}


# ============================================================
# 檔名安全處理
# ============================================================

def safe_filename(filename: str) -> str:

    if not filename:
        return "file"

    filename = Path(filename).name

    invalid_chars = '<>:"/\\|?*'

    for char in invalid_chars:
        filename = filename.replace(char, "")

    return filename or "file"


# ============================================================
# Job
# ============================================================

def create_job():

    job_id = uuid.uuid4().hex

    job_upload_dir = UPLOAD_DIR / job_id
    job_output_dir = OUTPUT_DIR / job_id

    job_upload_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    job_output_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    return (
        job_id,
        job_upload_dir,
        job_output_dir
    )


def cleanup_job(job_id: str):

    upload_path = UPLOAD_DIR / job_id
    output_path = OUTPUT_DIR / job_id

    try:
        shutil.rmtree(
            upload_path,
            ignore_errors=True
        )
    except Exception:
        pass

    try:
        shutil.rmtree(
            output_path,
            ignore_errors=True
        )
    except Exception:
        pass


# ============================================================
# Converter
# ============================================================

def converter_run(
    tool,
    input_paths,
    output_dir,
    output_format="PNG",
    quality=80
):

    converter = TurnDocConverter(
        output_dir=str(output_dir)
    )

    if tool == "word-to-pdf":
        return converter.word_to_pdf(
            input_paths[0]
        )

    if tool == "pdf-to-word":
        return converter.pdf_to_word(
            input_paths[0]
        )

    if tool == "ppt-to-pdf":
        return converter.ppt_to_pdf(
            input_paths[0]
        )

    if tool == "excel-to-pdf":
        return converter.excel_to_pdf(
            input_paths[0]
        )

    if tool == "pdf-to-ppt":
        return converter.pdf_to_ppt(
            input_paths[0]
        )

    if tool == "pdf-to-excel":
        return converter.pdf_to_excel(
            input_paths[0]
        )

    if tool == "merge":
        return converter.merge_pdfs(
            input_paths
        )

    if tool == "compress":
        return converter.compress_pdf(
            input_paths[0],
            quality
        )

    if tool == "pdf-to-image":
        return converter.pdf_to_images(
            input_paths[0],
            output_format
        )

    if tool == "image-to-pdf":
        return converter.images_to_pdf(
            input_paths
        )

    if tool == "image-convert":
        return converter.convert_image(
            input_paths[0],
            output_format
        )

    if tool == "image-compress":
        return converter.compress_image(
            input_paths[0],
            quality,
            None
            if output_format == "original"
            else output_format
        )

    if tool == "gif-split":
        return converter.split_gif(
            input_paths[0],
            output_format
        )

    raise ValueError(
        f"不支援工具: {tool}"
    )


# ============================================================
# Health Check
# ============================================================

@app.get("/api/health")
def health():

    return {
        "success": True,
        "service": "TurnDoc API",
        "version": "1.0.0",
        "status": "online"
    }


# ============================================================
# 工具列表
# ============================================================

@app.get("/api/tools")
def tools():

    return {
        "success": True,
        "tools": sorted(
            list(SUPPORTED_TOOLS)
        )
    }


# ============================================================
# 主要轉換 API
# ============================================================

@app.post("/api/convert")
async def convert(
    tool: str = Form(...),
    files: list[UploadFile] = File(...),
    outputFormat: str = Form("PNG"),
    quality: int = Form(80),
):

    if tool not in SUPPORTED_TOOLS:

        raise HTTPException(
            status_code=400,
            detail=f"不支援的工具: {tool}"
        )

    if not files:

        raise HTTPException(
            status_code=400,
            detail="請至少上傳一個檔案"
        )

    if quality < 1 or quality > 100:

        raise HTTPException(
            status_code=400,
            detail="quality 必須介於 1 到 100"
        )

    (
        job_id,
        job_upload_dir,
        job_output_dir
    ) = create_job()

    input_paths = []

    try:

        # ----------------------------------------------------
        # 儲存上傳檔案
        # ----------------------------------------------------

        for index, upload in enumerate(files):

            filename = safe_filename(
                upload.filename
            )

            if not filename:
                filename = f"file_{index}"

            file_path = (
                job_upload_dir /
                filename
            )

            with open(
                file_path,
                "wb"
            ) as buffer:

                while True:

                    chunk = await upload.read(
                        1024 * 1024
                    )

                    if not chunk:
                        break

                    buffer.write(chunk)

            input_paths.append(
                str(file_path)
            )

        # ----------------------------------------------------
        # 執行轉換
        # ----------------------------------------------------

        result = converter_run(
            tool=tool,
            input_paths=input_paths,
            output_dir=job_output_dir,
            output_format=outputFormat,
            quality=quality
        )

        if isinstance(result, list):
            output_paths = result
        else:
            output_paths = [result]

        # ----------------------------------------------------
        # 確認輸出
        # ----------------------------------------------------

        files_result = []

        for output_path in output_paths:

            output_path = Path(
                output_path
            )

            if not output_path.exists():
                continue

            filename = output_path.name

            files_result.append({
                "name": filename,
                "url": (
                    f"/api/download/"
                    f"{job_id}/"
                    f"{filename}"
                )
            })

        if not files_result:

            raise Exception(
                "轉換完成，但找不到輸出檔案"
            )

        # ----------------------------------------------------
        # 單一檔案
        # ----------------------------------------------------

        if len(files_result) == 1:

            return {
                "success": True,
                "jobId": job_id,
                "files": files_result,
                "download": (
                    files_result[0]["url"]
                )
            }

        # ----------------------------------------------------
        # 多個檔案 → ZIP
        # ----------------------------------------------------

        zip_name = "TurnDoc_result.zip"

        zip_path = (
            job_output_dir /
            zip_name
        )

        with zipfile.ZipFile(
            zip_path,
            "w",
            zipfile.ZIP_DEFLATED
        ) as zip_file:

            for output_path in output_paths:

                output_path = Path(
                    output_path
                )

                if output_path.exists():

                    zip_file.write(
                        output_path,
                        arcname=output_path.name
                    )

        return {
            "success": True,
            "jobId": job_id,
            "files": files_result,
            "zip": {
                "name": zip_name,
                "url": (
                    f"/api/download/"
                    f"{job_id}/"
                    f"{zip_name}"
                )
            }
        }

    except Exception as e:

        cleanup_job(job_id)

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )

    finally:

        # 上傳檔案刪除
        # 輸出檔案保留給下載

        try:

            shutil.rmtree(
                job_upload_dir,
                ignore_errors=True
            )

        except Exception:
            pass


# ============================================================
# 下載
# ============================================================

@app.get(
    "/api/download/{job_id}/{filename}"
)
def download(
    job_id: str,
    filename: str
):

    filename = safe_filename(
        filename
    )

    output_path = (
        OUTPUT_DIR /
        job_id /
        filename
    )

    if not output_path.exists():

        raise HTTPException(
            status_code=404,
            detail="找不到檔案，可能已過期"
        )

    return FileResponse(
        path=str(output_path),
        filename=filename,
        media_type="application/octet-stream"
    )


# ============================================================
# 靜態檔案（必須放在所有 /api 路由之後）
# ============================================================

@app.get("/{file_path:path}")
def static_files(file_path: str):

    # API 路徑不由這裡處理
    if file_path.startswith("api/"):
        raise HTTPException(
            status_code=404,
            detail="Not Found"
        )

    requested_file = PROJECT_DIR / file_path

    # 防止路徑穿越
    try:
        requested_file.resolve().relative_to(
            PROJECT_DIR.resolve()
        )
    except ValueError:
        raise HTTPException(
            status_code=404,
            detail="Not Found"
        )

    if not requested_file.is_file():
        raise HTTPException(
            status_code=404,
            detail="Not Found"
        )

    return FileResponse(
        path=str(requested_file)
    )


# ============================================================
# 啟動
# ============================================================

if __name__ == "__main__":

    import uvicorn

    port = int(
        os.environ.get(
            "PORT",
            "8000"
        )
    )

    uvicorn.run(
        app,
        host="0.0.0.0",
        port=port,
        reload=False
    )
