# Обновление дашборда РННТД раз в квартал

1. Скачать файл «Показатели проектов на ДД.ММ.ГГГГ.xlsx» с Google Drive (через коннектор Google Drive: `download_file_content`, base64 → .xlsx).
2. `python3 tools/rnntd_sync.py <файл.xlsx> tools/rnntd_attrs.csv <папка>` → `projects.csv`, `quarterly.csv`, `meta.json` (предупреждения об ошибках формул смотреть в `meta.json`).
3. Загрузить оба CSV в дашборд и обновить `url` и `updated` у источников `projects` и `quarterly`.
4. Справки: `python3 tools/rnntd_spravka.py <папка с CSV> spravki --variant all --format all [--year 2024]`.
