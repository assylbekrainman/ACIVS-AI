import os
import uuid
import shutil
from pathlib import Path
from fastapi import FastAPI, File, UploadFile, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
import aiofiles

from .config import settings
from .models import CompressionMode, CompressionResult
from .compressor import PDFCompressor

# Создаем директории
settings.UPLOAD_DIR.mkdir(exist_ok=True)
settings.COMPRESSED_DIR.mkdir(exist_ok=True)

app = FastAPI(
    title=settings.APP_NAME,
    description="🐰 Сверхмощный AI-агент для профессионального сжатия PDF",
    version="2.2.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Статические файлы
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/")
async def root():
    return FileResponse("static/index.html")

@app.post("/compress", response_model=CompressionResult)
async def compress_pdf(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    mode: CompressionMode = Form(CompressionMode.BALANCED),
    target_size_mb: float = Form(10.0)
):
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(400, "🐰 Я сжимаю только PDF! Другие форматы — не моя стихия.")

    file_id = str(uuid.uuid4())[:8]
    input_path = settings.UPLOAD_DIR / f"{file_id}_{file.filename}"
    output_path = settings.COMPRESSED_DIR / f"{file_id}_compressed.pdf"

    # Сохраняем загруженный файл
    try:
        content = await file.read()
        if len(content) > settings.MAX_FILE_SIZE:
            raise HTTPException(400, f"🐰 Файл слишком тяжёлый! Максимум {settings.MAX_FILE_SIZE / 1024 / 1024:.0f} МБ")

        async with aiofiles.open(input_path, 'wb') as f:
            await f.write(content)
    except Exception as e:
        raise HTTPException(500, f"Ошибка сохранения: {str(e)}")

    # Сжимаем
    compressor = PDFCompressor(mode=mode)
    success, result = compressor.compress(str(input_path), str(output_path), target_size_mb)

    if not success:
        raise HTTPException(500, result.get("message", "Unknown error"))

    # Добавляем URL для скачивания
    result["download_url"] = f"/download/{file_id}"

    # Очистка в фоне (через час)
    background_tasks.add_task(cleanup_files, str(input_path), str(output_path))

    return CompressionResult(**result)

@app.get("/download/{file_id}")
async def download(file_id: str):
    files = list(settings.COMPRESSED_DIR.glob(f"{file_id}_compressed.pdf"))
    if not files:
        raise HTTPException(404, "🐰 Файл не найден! Кролик, кажется, потерял его...")

    return FileResponse(
        files[0],
        media_type="application/pdf",
        filename="compressed.pdf"
    )

@app.get("/health")
async def health():
    return {"status": "🐰 Кролик бодр и готов к сжатию!"}

def cleanup_files(input_path: str, output_path: str):
    import time
    time.sleep(3600)
    for path in [input_path, output_path]:
        if os.path.exists(path):
            os.remove(path)
