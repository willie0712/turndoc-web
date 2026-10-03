import os
import shutil
import uuid
import zipfile
from pathlib import Path

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware

from converter import TurnDocConverter


# ============================================================
# TurnDoc Web API
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


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
# 首頁
# ============================================================

@app.get("/")
def root():
    return {
        "name": "TurnDoc API",
        "status": "online",
        "version": "1.0.0",
        "docs": "/docs",
        "health": "/api/health",
        "tools": "/api/tools"
    }


# ============================================================
# 工具
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


def safe_filename(filename: str) -> str:
    """
    避免使用者上傳奇怪檔名造成路徑問題。
    """
    if not filename:
        return "file"

    filename = Path(filename).name

    invalid_chars = '<>:"/\\|?*'

    for char in invalid_chars:
        filename = filename.replace(char, "")

    return filename or "file"


def create_job():
    """
    每次轉換建立獨立工作目錄。
    """
    job_id = uuid.uuid4().hex

    job_upload_dir = UPLOAD_DIR / job_id
    job_output_dir = OUTPUT_DIR / job_id

    job_upload_dir.mkdir(parents=True, exist_ok=True)
    job_output_dir.mkdir(parents=True, exist_ok=True)

    return job_id, job_upload_dir, job_output_dir


def cleanup_job(job_id: str):
    """
    清理暫存檔。
    """
    upload_path = UPLOAD_DIR / job_id
    output_path = OUTPUT_DIR / job_id

    try:
        shutil.rmtree(upload_path, ignore_errors=True)
    except Exception:
        pass

    try:
        shutil.rmtree(output_path, ignore_errors=True)
    except Exception:
        pass


def converter_run(
    tool,
    input_paths,
    output_dir,
    output_format="PNG",
    quality=80
):
    """
    使用你原本的 TurnDocConverter。
    """

    converter = TurnDocConverter(
        output_dir=str(output_dir)
    )

    if tool == "word-to-pdf":
        return converter.word_to_pdf(input_paths[0])

    if tool == "pdf-to-word":
        return converter.pdf_to_word(input_paths[0])

    if tool == "ppt-to-pdf":
        return converter.ppt_to_pdf(input_paths[0])

    if tool == "excel-to-pdf":
        return converter.excel_to_pdf(input_paths[0])

    if tool == "pdf-to-ppt":
        return converter.pdf_to_ppt(input_paths[0])

    if tool == "pdf-to-excel":
        return converter.pdf_to_excel(input_paths[0])

    if tool == "merge":
        return converter.merge_pdfs(input_paths)

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
            None if output_format == "original" else output_format
        )

    if tool == "gif-split":
        return converter.split_gif(
            input_paths[0],
            output_format
        )

    raise ValueError(f"不支援工具: {tool}")


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
        "tools": sorted(list(SUPPORTED_TOOLS))
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

    job_id, job_upload_dir, job_output_dir = create_job()

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

            file_path = job_upload_dir / filename

            with open(file_path, "wb") as buffer:
                while True:
                    chunk = await upload.read(1024 * 1024)

                    if not chunk:
                        break

                    buffer.write(chunk)

            input_paths.append(str(file_path))

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

            output_path = Path(output_path)

            if not output_path.exists():
                continue

            filename = output_path.name

            files_result.append({
                "name": filename,
                "url": f"/api/download/{job_id}/{filename}"
            })

        if not files_result:
            raise Exception("轉換完成，但找不到輸出檔案")

        # ----------------------------------------------------
        # 如果只有一個檔案
        # ----------------------------------------------------

        if len(files_result) == 1:

            return {
                "success": True,
                "jobId": job_id,
                "files": files_result,
                "download": files_result[0]["url"]
            }

        # ----------------------------------------------------
        # 多個檔案 → 自動 ZIP
        # ----------------------------------------------------

        zip_name = "TurnDoc_result.zip"
        zip_path = job_output_dir / zip_name

        with zipfile.ZipFile(
            zip_path,
            "w",
            zipfile.ZIP_DEFLATED
        ) as zip_file:

            for output_path in output_paths:

                output_path = Path(output_path)

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
                "url": f"/api/download/{job_id}/{zip_name}"
            }
        }

    except Exception as e:

        cleanup_job(job_id)

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )

    finally:

        # 上傳檔案可以刪除
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

@app.get("/api/download/{job_id}/{filename}")
def download(
    job_id: str,
    filename: str
):

    filename = safe_filename(filename)

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
# 啟動
# ============================================================

if __name__ == "__main__":
    import uvicorn

    port = int(os.environ.get("PORT", "8000"))

    uvicorn.run(
        app,
        host="0.0.0.0",
        port=port,
        reload=False
    )
