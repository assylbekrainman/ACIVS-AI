# 🐰 PDF Master Agent v2.2 «Кролик-Сжиматор»

Сверхмощный, умный и дружелюбный AI-агент для профессионального сжатия PDF-файлов.

## ✨ Возможности

- 📄 Сжатие PDF от 1 МБ до 500 МБ
- 🎯 Целевой размер ≤ 10 МБ (настраиваемый)
- ⚖️ 4 режима: Balanced, Quality First, Strong, Extreme
- 📸 Адаптивная оптимизация изображений через PyMuPDF + Pillow
- 👻 Fallback на Ghostscript для сложных случаев
- 🎨 Красивый веб-интерфейс с прогресс-барами и ASCII-кроликом
- 🔒 Автоудаление файлов через 1 час

## 🚀 Быстрый старт

### Docker (рекомендуется)

```bash
# Распаковать архив
cd pdf-master-agent

# Запустить
docker-compose up -d --build

# Открыть http://localhost:8000
```

### Локально (без Docker)

```bash
# Установить системные зависимости
# Ubuntu/Debian:
sudo apt-get update && sudo apt-get install -y ghostscript qpdf libmagic1 libgl1 libglib2.0-0

# Python зависимости
pip install -r requirements.txt

# Запуск
uvicorn app.main:app --reload

# Открыть http://localhost:8000
```

## 📡 API

### POST `/compress`

```bash
curl -X POST "http://localhost:8000/compress" \
  -F "file=@document.pdf" \
  -F "mode=balanced" \
  -F "target_size_mb=10"
```

**Параметры:**
- `file` — PDF файл
- `mode` — `balanced` | `quality_first` | `strong` | `extreme`
- `target_size_mb` — целевой размер в МБ (по умолчанию 10)

**Ответ:**
```json
{
  "success": true,
  "original_size_mb": 68.4,
  "compressed_size_mb": 9.4,
  "savings_mb": 59.0,
  "savings_percent": 86.3,
  "target_reached": true,
  "mode": "balanced",
  "techniques_applied": [
    "Page rasterization (DPI=150, JPEG=85)",
    "PyMuPDF optimization (garbage=4, deflate)"
  ],
  "pages": 42,
  "duration_seconds": 3.45,
  "message": "🎉 Кролик победил! Файл стал лёгким и счастливым!",
  "download_url": "/download/abc12345"
}
```

### GET `/download/{file_id}`

Скачивание сжатого файла.

### GET `/health`

Проверка состояния сервиса.

## 🛠 Технологии

- **FastAPI** — веб-фреймворк
- **PyMuPDF (fitz)** — рендеринг и оптимизация PDF
- **Pillow** — сжатие изображений
- **Ghostscript** — fallback для сложных PDF
- **Docker** — контейнеризация

## 🐰 Лицензия

MIT — Кролик делится добром!
