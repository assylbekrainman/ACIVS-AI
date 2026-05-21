import fitz
import os
import time
import io
import subprocess
from pathlib import Path
from PIL import Image
from typing import Tuple, List, Optional

from .config import settings
from .models import CompressionMode

class PDFCompressor:
    def __init__(self, mode: CompressionMode = CompressionMode.BALANCED):
        self.mode = mode
        self.techniques: List[str] = []
        self.pages = 0

    def _get_dpi(self) -> int:
        mapping = {
            CompressionMode.BALANCED: settings.DOWNSAMPLE_DPI_BALANCED,
            CompressionMode.QUALITY_FIRST: settings.DOWNSAMPLE_DPI_QUALITY,
            CompressionMode.STRONG: settings.DOWNSAMPLE_DPI_STRONG,
            CompressionMode.EXTREME: settings.DOWNSAMPLE_DPI_EXTREME,
        }
        return mapping.get(self.mode, 150)

    def _get_jpeg_quality(self) -> int:
        mapping = {
            CompressionMode.BALANCED: settings.JPEG_QUALITY_BALANCED,
            CompressionMode.QUALITY_FIRST: settings.JPEG_QUALITY_QUALITY,
            CompressionMode.STRONG: settings.JPEG_QUALITY_STRONG,
            CompressionMode.EXTREME: settings.JPEG_QUALITY_EXTREME,
        }
        return mapping.get(self.mode, 85)

    def analyze_pdf(self, file_path: str) -> dict:
        doc = fitz.open(file_path)
        info = {
            "pages": len(doc),
            "size_mb": round(os.path.getsize(file_path) / (1024 * 1024), 2),
            "has_images": False,
            "has_text": False,
            "image_count": 0,
        }

        for page_num in range(len(doc)):
            page = doc[page_num]
            if page.get_text().strip():
                info["has_text"] = True
            images = page.get_images(full=True)
            info["image_count"] += len(images)
            if images:
                info["has_images"] = True

        doc.close()
        return info

    def _compress_with_fitz_rasterize(self, input_path: str, output_path: str) -> bool:
        """
        Основной метод: рендерим каждую страницу в изображение с нужным DPI,
        сжимаем через Pillow и собираем новый PDF.
        Гарантированно уменьшает размер сканов и тяжелых PDF.
        """
        try:
            doc = fitz.open(input_path)
            self.pages = len(doc)
            dpi = self._get_dpi()
            quality = self._get_jpeg_quality()

            new_doc = fitz.open()

            for page_num in range(len(doc)):
                page = doc[page_num]
                rect = page.rect

                # Рендерим страницу с целевым DPI
                mat = fitz.Matrix(dpi/72, dpi/72)
                pix = page.get_pixmap(matrix=mat)

                # Конвертируем в PIL Image
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                # Сжимаем в JPEG
                buffer = io.BytesIO()
                img.save(buffer, format='JPEG', quality=quality, optimize=True, progressive=True)
                buffer.seek(0)

                # Создаем новую страницу оригинального размера
                new_page = new_doc.new_page(width=rect.width, height=rect.height)
                # Вставляем сжатое изображение на всю страницу
                new_page.insert_image(rect, stream=buffer.getvalue())

            # Сохраняем с максимальной оптимизацией
            new_doc.save(
                output_path,
                garbage=4,
                deflate=True,
                clean=True,
                linear=True
            )
            new_doc.close()
            doc.close()

            self.techniques.append(f"Page rasterization (DPI={dpi}, JPEG={quality})")
            self.techniques.append("PyMuPDF optimization (garbage=4, deflate)")
            return True

        except Exception as e:
            return False

    def _compress_with_ghostscript(self, input_path: str, output_path: str) -> bool:
        """Fallback через Ghostscript"""
        try:
            quality_map = {
                CompressionMode.BALANCED: "/ebook",
                CompressionMode.QUALITY_FIRST: "/prepress",
                CompressionMode.STRONG: "/ebook",
                CompressionMode.EXTREME: "/screen"
            }
            pdf_settings = quality_map.get(self.mode, "/ebook")

            cmd = [
                "gs",
                "-sDEVICE=pdfwrite",
                "-dCompatibilityLevel=1.4",
                f"-dPDFSETTINGS={pdf_settings}",
                "-dNOPAUSE",
                "-dQUIET",
                "-dBATCH",
                "-dColorImageResolution=150",
                "-dGrayImageResolution=150",
                "-dMonoImageResolution=300",
                f"-sOutputFile={output_path}",
                input_path
            ]

            result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
            if result.returncode == 0 and os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                self.techniques.append("Ghostscript optimization")
                return True
            return False
        except Exception:
            return False

    def compress(self, input_path: str, output_path: str, target_mb: float = None) -> Tuple[bool, dict]:
        start_time = time.time()
        target = target_mb or settings.TARGET_SIZE_MB

        # Анализ
        analysis = self.analyze_pdf(input_path)
        self.pages = analysis["pages"]
        original_size = analysis["size_mb"]

        # Если уже маленький — просто копируем
        if original_size <= target:
            import shutil
            shutil.copy(input_path, output_path)
            return True, {
                "success": True,
                "original_size_mb": original_size,
                "compressed_size_mb": original_size,
                "savings_mb": 0.0,
                "savings_percent": 0.0,
                "target_reached": True,
                "mode": self.mode.value,
                "techniques_applied": ["Файл уже лёгкий — скопирован как есть"],
                "pages": self.pages,
                "duration_seconds": round(time.time() - start_time, 2),
                "message": "🐰 Файл уже лёгкий! Скопировал как есть.",
                "download_url": None
            }

        # Попытка 1: PyMuPDF rasterization (самый надёжный)
        success = self._compress_with_fitz_rasterize(input_path, output_path)

        # Если не получилось или результат всё ещё больше цели в 2 раза — пробуем Ghostscript
        if not success or (os.path.exists(output_path) and os.path.getsize(output_path)/(1024*1024) > target * 2):
            gs_path = output_path + ".gs.pdf"
            if self._compress_with_ghostscript(input_path, gs_path):
                if os.path.exists(gs_path) and os.path.getsize(gs_path) > 0:
                    import shutil
                    shutil.move(gs_path, output_path)
            # Cleanup
            if os.path.exists(gs_path):
                os.remove(gs_path)

        if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
            return False, {
                "success": False,
                "message": "🐰 Не удалось создать выходной файл",
                "pages": self.pages,
                "duration_seconds": round(time.time() - start_time, 2)
            }

        compressed_size = os.path.getsize(output_path) / (1024 * 1024)
        savings = original_size - compressed_size
        percent = (savings / original_size * 100) if original_size > 0 else 0
        target_reached = compressed_size <= target

        if not target_reached and self.mode != CompressionMode.EXTREME:
            msg = f"⚠️ Не удалось достичь {target} МБ без потерь. Попробуй режим 'extreme'!"
        elif not target_reached:
            msg = "⚠️ Даже Extreme не уложился в лимит. Файл слишком тяжёлый."
        else:
            msg = "🎉 Кролик победил! Файл стал лёгким и счастливым!"

        return True, {
            "success": True,
            "original_size_mb": round(original_size, 2),
            "compressed_size_mb": round(compressed_size, 2),
            "savings_mb": round(savings, 2),
            "savings_percent": round(percent, 1),
            "target_reached": target_reached,
            "mode": self.mode.value,
            "techniques_applied": self.techniques,
            "pages": self.pages,
            "duration_seconds": round(time.time() - start_time, 2),
            "message": msg,
            "download_url": None
        }
