from pydantic import BaseModel
from enum import Enum
from typing import Optional, List

class CompressionMode(str, Enum):
    BALANCED = "balanced"
    QUALITY_FIRST = "quality_first"
    STRONG = "strong"
    EXTREME = "extreme"

class CompressionResult(BaseModel):
    success: bool
    original_size_mb: float
    compressed_size_mb: float
    savings_mb: float
    savings_percent: float
    target_reached: bool
    mode: str
    techniques_applied: List[str]
    download_url: Optional[str] = None
    message: str
    pages: int
    duration_seconds: float
