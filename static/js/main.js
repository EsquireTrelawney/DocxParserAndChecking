document.addEventListener('DOMContentLoaded', () => {
    // Инициализация темы
    initTheme();
    
    // Визуальная обратная связь при загрузке файла
    const fileInput = document.getElementById('docx_file');
    if (fileInput) {
        fileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            
            const fileName = file.name;
            const fileSize = formatFileSize(file.size);
            
            const fileStatus = document.createElement('div');
            fileStatus.className = 'file-status';
            fileStatus.innerHTML = `
                <span class="file-icon">📄</span>
                <div>
                    <strong>${fileName}</strong>
                    <div>${fileSize}</div>
                </div>
            `;
            
            // Удалить предыдущий статус, если он есть
            const previousStatus = document.querySelector('.file-status');
            if (previousStatus) {
                previousStatus.remove();
            }
            
            // Вставить новый статус после input-wrapper
            const fileInputWrapper = document.querySelector('.file-input-wrapper');
            if (fileInputWrapper) {
                fileInputWrapper.parentNode.appendChild(fileStatus);
            }
            
            // Активируем кнопку проверки
            const checkButton = document.querySelector('.check-button');
            if (checkButton) {
                checkButton.disabled = false;
            }
        });
    }
    
    // Обработка отправки формы - показать индикатор загрузки
    const form = document.querySelector('form');
    if (form) {
        form.addEventListener('submit', () => {
            // Создаем контейнер для сообщения о загрузке
            const loadingContainer = document.createElement('div');
            loadingContainer.className = 'loading-container';
            loadingContainer.innerHTML = `
                <div class="loading-spinner"></div>
                <p>Выполняется проверка документа...</p>
                <p class="loading-details">Это может занять несколько секунд</p>
            `;
            
            // Добавляем в интерфейс
            document.body.appendChild(loadingContainer);
            
            // Блокируем кнопку отправки
            const checkButton = document.querySelector('.check-button');
            if (checkButton) {
                checkButton.disabled = true;
                checkButton.textContent = 'Проверяем...';
            }
        });
    }
    
    // Обработчик переключения темы
    const themeToggle = document.querySelector('.theme-toggle');
    if (themeToggle) {
        themeToggle.addEventListener('click', () => {
            toggleTheme();
        });
    }
});

// Функция для форматирования размера файла
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Байт';
    
    const k = 1024;
    const sizes = ['Байт', 'КБ', 'МБ', 'ГБ'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Инициализация темы
function initTheme() {
    // Проверяем сохраненную тему или системные предпочтения
    const savedTheme = localStorage.getItem('theme');
    const prefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
    
    if (savedTheme === 'dark' || (!savedTheme && prefersDark)) {
        document.documentElement.setAttribute('data-theme', 'dark');
        updateThemeIcon('dark');
    } else {
        document.documentElement.setAttribute('data-theme', 'light');
        updateThemeIcon('light');
    }
}

// Переключение темы
function toggleTheme() {
    const currentTheme = document.documentElement.getAttribute('data-theme') || 'light';
    const newTheme = currentTheme === 'light' ? 'dark' : 'light';
    
    document.documentElement.setAttribute('data-theme', newTheme);
    localStorage.setItem('theme', newTheme);
    
    updateThemeIcon(newTheme);
}

// Обновление иконки переключателя темы
function updateThemeIcon(theme) {
    const themeToggle = document.querySelector('.theme-toggle');
    if (!themeToggle) return;
    
    if (theme === 'dark') {
        themeToggle.innerHTML = `
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
                <path d="M12 9c1.65 0 3 1.35 3 3s-1.35 3-3 3-3-1.35-3-3 1.35-3 3-3m0-2c-2.76 0-5 2.24-5 5s2.24 5 5 5 5-2.24 5-5-2.24-5-5-5zM2 13h2c.55 0 1-.45 1-1s-.45-1-1-1H2c-.55 0-1 .45-1 1s.45 1 1 1zm18 0h2c.55 0 1-.45 1-1s-.45-1-1-1h-2c-.55 0-1 .45-1 1s.45 1 1 1zM11 2v2c0 .55.45 1 1 1s1-.45 1-1V2c0-.55-.45-1-1-1s-1 .45-1 1zm0 18v2c0 .55.45 1 1 1s1-.45 1-1v-2c0-.55-.45-1-1-1s-1 .45-1 1zM5.99 4.58c-.39-.39-1.03-.39-1.41 0-.39.39-.39 1.03 0 1.41l1.06 1.06c.39.39 1.03.39 1.41 0 .39-.39.39-1.03 0-1.41L5.99 4.58zm12.37 12.37c-.39-.39-1.03-.39-1.41 0-.39.39-.39 1.03 0 1.41l1.06 1.06c.39.39 1.03.39 1.41 0 .39-.39.39-1.03 0-1.41l-1.06-1.06zm1.06-10.96c.39-.39.39-1.03 0-1.41-.39-.39-1.03-.39-1.41 0l-1.06 1.06c-.39.39-.39 1.03 0 1.41.39.39 1.03.39 1.41 0l1.06-1.06zM7.05 18.36c.39-.39.39-1.03 0-1.41-.39-.39-1.03-.39-1.41 0l-1.06 1.06c-.39.39-.39 1.03 0 1.41.39.39 1.03.39 1.41 0l1.06-1.06z"/>
            </svg>
        `;
    } else {
        themeToggle.innerHTML = `
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
                <path d="M9.37 5.51c-.18.64-.27 1.31-.27 1.99 0 4.08 3.32 7.4 7.4 7.4.68 0 1.35-.09 1.99-.27C17.45 17.19 14.93 19 12 19c-3.86 0-7-3.14-7-7 0-2.93 1.81-5.45 4.37-6.49zM12 3c-4.97 0-9 4.03-9 9s4.03 9 9 9 9-4.03 9-9c0-.46-.04-.92-.1-1.36-.98 1.37-2.58 2.26-4.4 2.26-2.98 0-5.4-2.42-5.4-5.4 0-1.81.89-3.42 2.26-4.4-.44-.06-.9-.1-1.36-.1z"/>
            </svg>
        `;
    }
} 