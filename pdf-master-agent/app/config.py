from pydantic_settings import BaseSettings
from pathlib import Path

class Settings(BaseSettings):
    APP_NAME: str = "PDF Master Agent v2.2 🐰"
    MAX_FILE_SIZE: int = 500 * 1024 * 1024
    UPLOAD_DIR: Path = Path("uploads")
    COMPRESSED_DIR: Path = Path("compressed")
    TARGET_SIZE_MB: float = 10.0

    DOWNSAMPLE_DPI_BALANCED: int = 150
    DOWNSAMPLE_DPI_QUALITY: int = 200
    DOWNSAMPLE_DPI_STRONG: int = 100
    DOWNSAMPLE_DPI_EXTREME: int = 72

    JPEG_QUALITY_BALANCED: int = 85
    JPEG_QUALITY_QUALITY: int = 95
    JPEG_QUALITY_STRONG: int = 70
    JPEG_QUALITY_EXTREME: int = 50

settings = Settings()
