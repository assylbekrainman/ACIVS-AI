# Обновление дашборда РННТД раз в квартал

1. Скачать файл «Показатели проектов на ДД.ММ.ГГГГ.xlsx» с Google Drive (коннектор Google Drive: `download_file_content`, base64 → .xlsx).
2. Реестр атрибутов проектов (`tools/rnntd_registry.csv`) при необходимости пересобрать:
   `python3 tools/rnntd_registry.py <реестр.xlsx> tools/rnntd_registry.csv`
3. `python3 tools/rnntd_sync.py <показатели.xlsx> tools/rnntd_registry.csv <папка>` → `projects.csv` (все проекты, объединение реестра и Drive),
   `history.csv`, `quarterly.csv`, `quality.csv`, `recon.csv`, `quality.json`.
4. Загрузить CSV в дашборд и обновить `url` и `updated` у источников `projects`, `history`, `quarterly`, `quality`, `recon`.
5. Справки: `python3 tools/rnntd_spravka.py <папка с CSV> spravki --variant all --format all [--year 2024]`.
