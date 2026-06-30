// ===== AI转Word助手 — UI 组件模块 =====

import { THEME_CONFIG } from './constants.js';
import { saveState } from './state.js';

// ===== Toast 通知 =====
/**
 * 显示 Toast 通知
 * @param {string} message - 消息内容
 * @param {string} type - 类型 ('info' | 'success' | 'error' | 'warning')
 */
function showToast(message, type = 'info') {
    const container = document.querySelector('.toast-container') || createToastContainer();
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 3000);
}

function createToastContainer() {
    const container = document.createElement('div');
    container.className = 'toast-container';
    document.body.appendChild(container);
    return container;
}

// ===== 主题切换 =====
/**
 * 循环切换主题
 */
function cycleTheme() {
    const { themes, icons } = THEME_CONFIG;
    const currentTheme = window.__state?.currentTheme || 'light';
    const idx = themes.indexOf(currentTheme);
    const nextTheme = themes[(idx + 1) % themes.length];
    
    if (window.__state) {
        window.__state.currentTheme = nextTheme;
    }
    saveState();
    applyTheme();
}

/**
 * 应用主题
 */
function applyTheme() {
    const currentTheme = window.__state?.currentTheme || 'light';
    const icons = THEME_CONFIG.icons;
    
    document.documentElement.setAttribute('data-theme', currentTheme);
    
    const themeIcon = document.querySelector('#theme-icon');
    if (themeIcon) {
        themeIcon.textContent = icons[currentTheme] || icons.light;
    }
}

// ===== 分割面板拖拽 =====
let isDragging = false;

/**
 * 设置分割面板拖拽
 */
function setupDividerDrag() {
    const divider = document.querySelector('#divider');
    if (!divider) return;

    divider.addEventListener('mousedown', (e) => {
        isDragging = true;
        document.body.style.cursor = 'col-resize';
        document.body.style.userSelect = 'none';
    });

    document.addEventListener('mousemove', (e) => {
        if (!isDragging) return;
        const container = document.querySelector('.split-container');
        if (!container) return;
        const rect = container.getBoundingClientRect();
        const ratio = (e.clientX - rect.left) / rect.width;
        const splitRatio = Math.max(0.25, Math.min(0.75, ratio));
        container.style.gridTemplateColumns = `${splitRatio * 100}% ${((1 - splitRatio) * 100).toFixed(2)}%`;
        
        if (window.__state) {
            window.__state.splitRatio = splitRatio;
        }
    });

    document.addEventListener('mouseup', () => {
        if (isDragging) {
            isDragging = false;
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
            saveState();
        }
    });
}

// ===== 拖拽上传 =====
/**
 * 设置拖拽上传
 */
function setupDragDrop() {
    const dz = document.querySelector('#drop-zone');
    if (!dz) return;

    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(evt => {
        dz.addEventListener(evt, (e) => { e.preventDefault(); e.stopPropagation(); });
    });

    dz.addEventListener('dragover', () => dz.classList.add('drag-over'));
    dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
    dz.addEventListener('drop', (e) => {
        dz.classList.remove('drag-over');
        const files = e.dataTransfer.files;
        if (files.length > 0 && window.__fileHandler) {
            window.__fileHandler(files[0]);
        }
    });

    // Click to upload
    dz.addEventListener('click', () => {
        const fileInput = document.querySelector('#file-input');
        if (fileInput) fileInput.click();
    });
}

// ===== 文件处理 =====
/**
 * 处理文件
 * @param {File} file - 文件对象
 */
function processFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    const reader = new FileReader();
    
    reader.onload = (e) => {
        const markdownInput = document.querySelector('#markdown-input');
        
        if (ext === 'docx') {
            extractDocxText(e.target.result, markdownInput);
        } else {
            if (markdownInput) {
                markdownInput.value = e.target.result;
                markdownInput.dispatchEvent(new Event('input'));
            }
            showToast('文件已加载', 'success');
        }
    };
    
    if (ext === 'docx') {
        reader.readAsArrayBuffer(file);
    } else {
        reader.readAsText(file);
    }
}

/**
 * 从 docx 提取文本
 * @param {ArrayBuffer} arrayBuffer - 文件内容
 * @param {HTMLTextAreaElement} markdownInput - 目标输入框
 */
function extractDocxText(arrayBuffer, markdownInput) {
    showToast('正在解析文档...', 'info');

    const reader = new FileReader();
    reader.onload = (e) => {
        const bytes = new Uint8Array(e.target.result);
        let text = '';

        const decoder = new TextDecoder('latin1');
        const str = decoder.decode(bytes);

        // Extract text between <w:t> and </w:t>
        const matches = str.match(/<w:t[^>]*>([^<]*)<\/w:t>/g);
        if (matches) {
            text = matches.map(m => {
                const inner = m.match(/>([^<]*)</);
                return inner ? inner[1] : '';
            }).join('');
        }

        if (text && markdownInput) {
            markdownInput.value = text;
            markdownInput.dispatchEvent(new Event('input'));
            showToast('文档已加载', 'success');
        } else {
            showToast('无法读取文档内容，请尝试上传为 .txt 或 .md 文件', 'error');
        }
    };
    reader.readAsArrayBuffer(arrayBuffer);
}

// ===== 预览和统计 =====
/**
 * 更新预览
 * @param {string} markdownInputId - 输入框 ID
 * @param {string} previewContentId - 预览区 ID
 */
function updatePreview(markdownInputId, previewContentId) {
    const input = document.querySelector(markdownInputId);
    const preview = document.querySelector(previewContentId);
    if (!input || !preview) return;

    const text = input.value.trim();
    if (!text) {
        preview.innerHTML = `
            <div class="empty-preview">
                <div class="icon">📝</div>
                <p>粘贴AI文案后点击"导出Word"查看效果</p>
                <p style="font-size: 0.8rem; margin-top: 0.5rem; opacity: 0.7;">支持 Markdown 格式：标题、加粗、列表、表格等</p>
            </div>`;
        return;
    }
    
    try {
        preview.innerHTML = marked.parse(text);
    } catch(e) {
        preview.textContent = text;
    }
}

/**
 * 更新统计
 * @param {string} markdownInputId - 输入框 ID
 * @param {Object} statElements - 统计元素对象
 */
function updateStats(markdownInputId, statElements) {
    const input = document.querySelector(markdownInputId);
    if (!input) return;

    const text = input.value.trim();
    const { words, chars, headings, tables } = statElements;
    
    if (!text) {
        if (words) words.textContent = '0';
        if (chars) chars.textContent = '0';
        if (headings) headings.textContent = '0';
        if (tables) tables.textContent = '0';
        return;
    }

    const wordCount = text.replace(/\s/g, '').length;
    const charCount = text.length;
    const headingCount = (text.match(/^#{1,6}\s/gm) || []).length;
    const tableCount = (text.match(/^\|/gm) || []).length;

    if (words) words.textContent = wordCount;
    if (chars) chars.textContent = charCount;
    if (headings) headings.textContent = headingCount;
    if (tables) tables.textContent = Math.floor(tableCount / 2);
}

export {
    showToast,
    cycleTheme,
    applyTheme,
    setupDividerDrag,
    setupDragDrop,
    processFile,
    extractDocxText,
    updatePreview,
    updateStats
};
