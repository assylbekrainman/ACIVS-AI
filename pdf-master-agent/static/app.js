class RabbitCompressor {
    constructor() {
        this.dropZone = document.getElementById('dropZone');
        this.fileInput = document.getElementById('fileInput');
        this.fileInfo = document.getElementById('fileInfo');
        this.optionsPanel = document.getElementById('optionsPanel');
        this.compressBtn = document.getElementById('compressBtn');
        this.progressCard = document.getElementById('progressCard');
        this.resultCard = document.getElementById('resultCard');
        this.uploadCard = document.getElementById('uploadCard');
        this.logTerminal = document.getElementById('logTerminal');
        this.progressBar = document.getElementById('progressBar');
        this.statusText = document.getElementById('statusText');
        this.rabbitAnim = document.getElementById('rabbitAnim');
        this.featuresSection = document.getElementById('featuresSection');
        this.currentFile = null;
        this.API_BASE = ''; // Относительные пути — тот же хост

        this.init();
    }

    init() {
        // Drag & Drop
        this.dropZone.addEventListener('click', () => this.fileInput.click());

        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            this.dropZone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
            });
        });

        ['dragenter', 'dragover'].forEach(eventName => {
            this.dropZone.addEventListener(eventName, () => {
                this.dropZone.classList.add('dragover');
            });
        });

        ['dragleave', 'drop'].forEach(eventName => {
            this.dropZone.addEventListener(eventName, () => {
                this.dropZone.classList.remove('dragover');
            });
        });

        this.dropZone.addEventListener('drop', (e) => {
            const files = e.dataTransfer.files;
            if (files.length) this.handleFile(files[0]);
        });

        this.fileInput.addEventListener('change', (e) => {
            if (e.target.files.length) this.handleFile(e.target.files[0]);
        });

        this.compressBtn.addEventListener('click', () => this.startCompression());
        document.getElementById('againBtn').addEventListener('click', () => this.reset());
    }

    handleFile(file) {
        if (!file.name.toLowerCase().endsWith('.pdf')) {
            this.log('❌ Кролик сжимает только PDF! Попробуй другой файл.', true);
            this.shake(this.dropZone);
            return;
        }

        this.currentFile = file;
        document.getElementById('fileName').textContent = file.name;
        document.getElementById('fileSize').textContent = this.formatBytes(file.size);
        this.fileInfo.classList.add('visible');

        this.optionsPanel.style.display = 'grid';
        this.compressBtn.style.display = 'flex';

        this.log(`🐰 Кролик-Сжиматор проснулся и схватил твой тяжёлый PDF! (${this.formatBytes(file.size)})`);
    }

    async startCompression() {
        if (!this.currentFile) return;

        const mode = document.getElementById('mode').value;
        const targetSize = parseFloat(document.getElementById('targetSize').value);

        this.optionsPanel.style.display = 'none';
        this.compressBtn.style.display = 'none';
        this.progressCard.classList.add('visible');
        this.featuresSection.style.display = 'none';
        this.scrollTo(this.progressCard);

        const formData = new FormData();
        formData.append('file', this.currentFile);
        formData.append('mode', mode);
        formData.append('target_size_mb', targetSize);

        // Симулируем прогресс пока ждем ответ
        let progress = 0;
        const progressInterval = setInterval(() => {
            progress = Math.min(progress + Math.random() * 12, 90);
            this.updateProgress(progress, this.getRandomStatus(progress));
            this.log(this.getRandomLog(progress));
            this.animateRabbit(progress);
        }, 1500);

        try {
            const response = await fetch(`${this.API_BASE}/compress`, {
                method: 'POST',
                body: formData
            });

            clearInterval(progressInterval);

            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.detail || 'Ошибка сервера');
            }

            const data = await response.json();
            this.updateProgress(100, '✅ Готово, чемпион!');
            this.log('Сжатие завершено успешно!');
            this.animateRabbit(100);

            await this.sleep(800);
            this.showResult(data);

        } catch (error) {
            clearInterval(progressInterval);
            this.updateProgress(100, '❌ Ошибка!', true);
            this.log(`❌ Ошибка: ${error.message}`, true);
            this.statusText.textContent = '🐰 Кролик споткнулся... Попробуй ещё раз!';

            await this.sleep(2000);
            this.reset();
        }
    }

    showResult(data) {
        this.progressCard.style.display = 'none';
        this.resultCard.classList.add('visible');
        this.scrollTo(this.resultCard);

        document.getElementById('resultMessage').textContent = data.message;

        const emoji = data.target_reached ? '🎉' : '⚠️';
        const title = data.target_reached ? 'Готово, чемпион!' : 'Сжатие выполнено';
        document.getElementById('resultEmoji').textContent = emoji;
        document.getElementById('resultTitle').textContent = title;

        const grid = document.getElementById('statsGrid');
        grid.innerHTML = `
            <div class="stat-card">
                <div class="stat-value">${data.pages}</div>
                <div class="stat-label">📄 Страниц</div>
            </div>
            <div class="stat-card">
                <div class="stat-value">${data.original_size_mb} МБ</div>
                <div class="stat-label">Исходный размер</div>
            </div>
            <div class="stat-card ${data.compressed_size_mb <= 10 ? 'positive' : ''}">
                <div class="stat-value">${data.compressed_size_mb} МБ</div>
                <div class="stat-label">Новый размер</div>
            </div>
            <div class="stat-card positive">
                <div class="stat-value">-${data.savings_mb} МБ</div>
                <div class="stat-label">Экономия (${data.savings_percent}%)</div>
            </div>
            <div class="stat-card ${data.target_reached ? 'positive' : 'negative'}">
                <div class="stat-value">${data.target_reached ? '✅ Да!' : '⚠️ Нет'}</div>
                <div class="stat-label">Цель ≤ 10 МБ</div>
            </div>
            <div class="stat-card">
                <div class="stat-value">${data.duration_seconds}с</div>
                <div class="stat-label">Время работы</div>
            </div>
        `;

        if (data.techniques_applied && data.techniques_applied.length > 0) {
            document.getElementById('techniquesBlock').style.display = 'block';
            const list = document.getElementById('techniquesList');
            list.innerHTML = data.techniques_applied.map(t => 
                `<li>✨ ${t}</li>`
            ).join('');
        } else {
            document.getElementById('techniquesBlock').style.display = 'none';
        }

        const downloadBtn = document.getElementById('downloadBtn');
        if (data.download_url) {
            downloadBtn.href = this.API_BASE + data.download_url;
            downloadBtn.style.display = 'flex';
        } else {
            downloadBtn.style.display = 'none';
        }
    }

    updateProgress(percent, text, isError = false) {
        this.progressBar.style.width = percent + '%';
        this.statusText.textContent = text;
        if (isError) {
            this.statusText.style.color = 'var(--error)';
        } else {
            this.statusText.style.color = 'var(--dark)';
        }
    }

    animateRabbit(percent) {
        let intensity = '';
        if (percent > 30) intensity = '  💨';
        if (percent > 60) intensity = '  🔥';
        if (percent >= 100) intensity = '  ✨';

        this.rabbitAnim.innerHTML = `   🐰${intensity}<br>  /|\<br> / | \<br>   |<br>  / \\`;
    }

    getRandomStatus(progress) {
        const stages = [
            { max: 15, msgs: ['🔍 Кролик внимательно изучает файл...', '📄 Читаю структуру PDF...'] },
            { max: 35, msgs: ['📸 Запихиваю картинки в маленький чемодан... 😤', '🎨 Оптимизирую изображения...'] },
            { max: 55, msgs: ['🗜 Кролик давит папки лапками изо всех сил!', '💪 Сжимаю потоки данных...'] },
            { max: 75, msgs: ['🧹 Вычищаю мусор и дубликаты...', '🔧 Дедупликация объектов...'] },
            { max: 90, msgs: ['✨ Финальная полировка...', '📦 Упаковываю результат...'] },
            { max: 100, msgs: ['⏳ Почти готово...', '🐰 Финишная прямая!'] }
        ];

        for (const stage of stages) {
            if (progress <= stage.max) {
                return stage.msgs[Math.floor(Math.random() * stage.msgs.length)];
            }
        }
        return '⏳ Кролик усердно работает...';
    }

    getRandomLog(progress) {
        const logs = [
            { max: 20, msgs: ['Анализ структуры PDF...', 'Обнаружено страниц: загрузка...'] },
            { max: 40, msgs: ['Downsampling изображений...', 'JPEG recompression...'] },
            { max: 60, msgs: ['Удаление метаданных...', 'Очистка JavaScript...'] },
            { max: 80, msgs: ['Object deduplication...', 'Linearization...'] },
            { max: 100, msgs: ['Финальная оптимизация...', 'Проверка целостности...'] }
        ];

        for (const stage of logs) {
            if (progress <= stage.max) {
                return stage.msgs[Math.floor(Math.random() * stage.msgs.length)];
            }
        }
        return `Прогресс: ${Math.round(progress)}%`;
    }

    log(message, isError = false) {
        const entry = document.createElement('div');
        entry.className = 'log-entry';
        entry.style.color = isError ? '#ff7675' : '#00ff88';
        entry.textContent = `> ${message}`;
        this.logTerminal.appendChild(entry);
        this.logTerminal.scrollTop = this.logTerminal.scrollHeight;
    }

    reset() {
        this.currentFile = null;
        this.fileInput.value = '';

        this.fileInfo.classList.remove('visible');
        this.optionsPanel.style.display = 'none';
        this.compressBtn.style.display = 'none';
        this.progressCard.classList.remove('visible');
        this.progressCard.style.display = '';
        this.resultCard.classList.remove('visible');
        this.resultCard.style.display = '';
        this.featuresSection.style.display = 'grid';

        this.progressBar.style.width = '0%';
        this.statusText.textContent = '';
        this.statusText.style.color = '';
        this.logTerminal.innerHTML = '<div class="log-entry">> Инициализация Кролика-Сжиматора v2.2...</div>';

        document.getElementById('dropText').textContent = 'Перетащи PDF сюда или кликни для выбора';
        document.getElementById('techniquesBlock').style.display = 'none';

        this.scrollTo(this.uploadCard);
    }

    shake(element) {
        element.style.animation = 'shake 0.5s';
        setTimeout(() => {
            element.style.animation = '';
        }, 500);
    }

    scrollTo(element) {
        element.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    formatBytes(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

document.addEventListener('DOMContentLoaded', () => {
    new RabbitCompressor();
});
