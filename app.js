// ===== AI转Word助手 — 主入口 (v3.1 模块化重构) =====

import { DEFAULT_TEMPLATES, MODE_NAMES, THEME_CONFIG } from './js/constants.js';
import { exportToWord, batchExport } from './js/exporter.js';
import { showToast, cycleTheme, applyTheme, setupDividerDrag, setupDragDrop, processFile, updatePreview, updateStats } from './js/ui.js';
import { handlePaste } from './js/textProcessor.js';

// ===== 全局状态 (通过 window 共享) =====
window.__state = {
    currentTemplate: 'official',
    currentMode: 'purify',
    currentTheme: 'light',
    customSettings: null,
    customTemplates: [],
    splitRatio: 0.5
};

const STORAGE_KEY = 'ai-word-converter-v3';

// ===== DOM 选择器 =====
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

// ===== 初始化 =====
function init() {
    loadState();
    renderTemplateGrid();
    renderCustomTemplateList();
    setupEventListeners();
    applyTheme();
    updateTemplateSelection();
    updateModeSelection();
    updatePreview('#markdown-input', '#preview-content');
    updateStats('#markdown-input', {
        words: $('#stat-words'),
        chars: $('#stat-chars'),
        headings: $('#stat-headings'),
        tables: $('#stat-tables')
    });
    
    // 同步状态到 window
    syncStateToWindow();
}

// ===== 状态持久化 =====
function loadState() {
    try {
        const saved = localStorage.getItem(STORAGE_KEY);
        if (saved) {
            const s = JSON.parse(saved);
            window.__state.currentTemplate = s.template || 'official';
            window.__state.currentMode = s.mode || 'purify';
            window.__state.currentTheme = s.theme || 'light';
            window.__state.customSettings = s.customSettings || null;
            window.__state.customTemplates = s.customTemplates || [];
            window.__state.splitRatio = s.splitRatio ?? 0.5;
        }
    } catch(e) { console.warn('Load state failed:', e); }
}

function saveState() {
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify({
            template: window.__state.currentTemplate,
            mode: window.__state.currentMode,
            theme: window.__state.currentTheme,
            customSettings: window.__state.customSettings,
            customTemplates: window.__state.customTemplates,
            splitRatio: window.__state.splitRatio
        }));
    } catch(e) { console.warn('Save state failed:', e); }
}

function syncStateToWindow() {
    // 同步到 window 以便其他模块访问
    window.__state.currentTemplate = window.__state.currentTemplate;
    window.__state.currentMode = window.__state.currentMode;
    window.__state.currentTheme = window.__state.currentTheme;
}

// ===== 事件绑定 =====
function setupEventListeners() {
    // Theme
    $('#theme-cycle')?.addEventListener('click', () => {
        cycleTheme();
    });

    // Template settings modal
    $('#template-settings')?.addEventListener('click', () => openSettingsModal());
    $('#close-settings')?.addEventListener('click', () => closeSettingsModal());
    $('#template-modal')?.addEventListener('click', (e) => {
        if (e.target === $('#template-modal')) closeSettingsModal();
    });

    // Save/Reset settings
    $('#save-template')?.addEventListener('click', saveCurrentSettings);
    $('#save-template-modal')?.addEventListener('click', saveCurrentSettings);
    $('#reset-template')?.addEventListener('click', resetSettingsForm);

    // Template cards
    $('#template-grid')?.addEventListener('click', (e) => {
        const card = e.target.closest('.template-card');
        if (!card) return;
        selectTemplate(card.dataset.template);
    });

    // Mode selection
    $$('input[name="mode"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            window.__state.currentMode = e.target.value;
            saveState();
            updatePreview('#markdown-input', '#preview-content');
            updateStats('#markdown-input', {
                words: $('#stat-words'),
                chars: $('#stat-chars'),
                headings: $('#stat-headings'),
                tables: $('#stat-tables')
            });
        });
    });

    // File input
    $('#upload-btn')?.addEventListener('click', () => $('#file-input')?.click());
    $('#file-input')?.addEventListener('change', handleFileUpload);

    // Batch upload
    $('#batch-upload-btn')?.addEventListener('click', () => $('#batch-file-input')?.click());
    $('#batch-file-input')?.addEventListener('change', handleBatchFileUpload);

    // Drag & Drop
    setupDragDrop();

    // Textarea input
    $('#markdown-input')?.addEventListener('input', () => {
        updatePreview('#markdown-input', '#preview-content');
        updateStats('#markdown-input', {
            words: $('#stat-words'),
            chars: $('#stat-chars'),
            headings: $('#stat-headings'),
            tables: $('#stat-tables')
        });
    });

    // Paste interception for purify mode
    $('#markdown-input')?.addEventListener('paste', (e) => {
        const result = handlePaste(e);
        if (result !== null) {
            setTimeout(() => {
                updatePreview('#markdown-input', '#preview-content');
                updateStats('#markdown-input', {
                    words: $('#stat-words'),
                    chars: $('#stat-chars'),
                    headings: $('#stat-headings'),
                    tables: $('#stat-tables')
                });
            }, 50);
        }
    });

    // Export
    $('#export-btn')?.addEventListener('click', () => {
        const text = $('#markdown-input')?.value || '';
        exportToWord(text, window.__state.currentMode);
    });

    // Custom template management
    $('#customize-template')?.addEventListener('click', () => {
        window.__state.currentTemplate = 'custom';
        openSettingsModal();
    });
    $('#add-custom-template-btn')?.addEventListener('click', saveAsCustomTemplate);

    // Clear button
    $('#clear-btn')?.addEventListener('click', () => {
        const ta = $('#markdown-input');
        if (ta) {
            ta.value = '';
            updatePreview('#markdown-input', '#preview-content');
            updateStats('#markdown-input', {
                words: $('#stat-words'),
                chars: $('#stat-chars'),
                headings: $('#stat-headings'),
                tables: $('#stat-tables')
            });
            showToast('已清除内容', 'info');
        }
    });

    // Paste button
    $('#paste-btn')?.addEventListener('click', async () => {
        const ta = $('#markdown-input');
        if (!ta) return;
        ta.focus();
        try {
            const text = await navigator.clipboard.readText();
            document.execCommand('insertText', false, text);
        } catch {
            const text = prompt('从剪贴板粘贴，请粘贴你的文案：');
            if (text) {
                ta.value = text;
                ta.dispatchEvent(new Event('input'));
            }
        }
    });

    // File input — 限制文件类型
    const fileInput = $('#file-input');
    if (fileInput) fileInput.setAttribute('accept', '.txt,.md,.docx,.html');

    // Divider drag
    setupDividerDrag();

    // Keyboard shortcuts
    document.addEventListener('keydown', (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key === 's') {
            e.preventDefault();
            const text = $('#markdown-input')?.value || '';
            exportToWord(text, window.__state.currentMode);
        }
        if ((e.ctrlKey || e.metaKey) && e.key === '1') {
            window.__state.currentTheme = 'light'; saveState(); applyTheme();
        }
        if ((e.ctrlKey || e.metaKey) && e.key === '2') {
            window.__state.currentTheme = 'dark'; saveState(); applyTheme();
        }
        if ((e.ctrlKey || e.metaKey) && e.key === '3') {
            window.__state.currentTheme = 'warm'; saveState(); applyTheme();
        }
    });
}

// ===== 文件处理 =====
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    processFile(file);
    e.target.value = '';
}

/**
 * 处理批量文件上传
 * @param {Event} e - change 事件
 */
function handleBatchFileUpload(e) {
    const files = Array.from(e.target.files);
    if (!files || files.length === 0) return;

    // 读取所有文件内容
    const filePromises = files.map(file => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                resolve({
                    name: file.name,
                    content: e.target.result
                });
            };
            reader.onerror = () => reject(new Error(`Failed to read ${file.name}`));
            reader.readAsText(file);
        });
    });

    Promise.all(filePromises)
        .then(filesData => {
            // 显示文件数量提示
            if (filesData.length > 0) {
                showToast(`已读取 ${filesData.length} 个文件，开始批量导出...`, 'info');
                batchExport(filesData, window.__state.currentMode);
            }
        })
        .catch(err => {
            showToast(`文件读取失败: ${err.message}`, 'error');
        });

    e.target.value = '';
}

// ===== 模板选择 =====
function selectTemplate(name) {
    window.__state.currentTemplate = name;
    if (name !== 'custom') {
        window.__state.customSettings = null;
    }
    saveState();
    updateTemplateSelection();
    updatePreview('#markdown-input', '#preview-content');
    updateTemplatePreview();
    updateStats('#markdown-input', {
        words: $('#stat-words'),
        chars: $('#stat-chars'),
        headings: $('#stat-headings'),
        tables: $('#stat-tables')
    });
}

function updateTemplateSelection() {
    $$('.template-card').forEach(card => {
        card.classList.toggle('active', card.dataset.template === window.__state.currentTemplate);
    });
}

function renderTemplateGrid() {
    const grid = $('#template-grid');
    if (!grid) return;

    let html = '';

    Object.entries(DEFAULT_TEMPLATES).forEach(([key, tpl]) => {
        html += `
            <div class="template-card ${window.__state.currentTemplate === key ? 'active' : ''}" data-template="${key}">
                <div class="template-icon">${tpl.icon}</div>
                <div class="template-name">${tpl.name}</div>
                <div class="template-desc">${tpl.desc}</div>
                <div class="template-card-actions">
                    <button class="card-action-btn detail-btn" onclick="showTemplateDetail('${key}')" title="查看详情">🔍</button>
                </div>
            </div>`;
    });

    window.__state.customTemplates.forEach((tpl, idx) => {
        html += `
            <div class="template-card ${window.__state.currentTemplate === 'custom_' + idx ? 'active' : ''}" data-template="custom_${idx}">
                <div class="template-icon">⭐</div>
                <div class="template-name">${tpl.name}</div>
                <div class="template-desc">自定义模板</div>
                <div class="template-custom-badge">已保存</div>
                <div class="template-card-actions">
                    <button class="card-action-btn edit-btn" onclick="editCustomTemplate(${idx})" title="编辑">✏️</button>
                    <button class="card-action-btn delete-btn" onclick="deleteCustomTemplate(${idx})" title="删除">🗑️</button>
                </div>
            </div>`;
    });

    html += `
        <div class="template-card" data-template="custom" id="add-custom-card" style="border-style: dashed; opacity: 0.7;">
            <div class="template-icon">➕</div>
            <div class="template-name">新建自定义</div>
            <div class="template-desc">保存当前设置为新模板</div>
        </div>`;

    grid.innerHTML = html;

    const addCard = $('#add-custom-card');
    if (addCard) {
        addCard.addEventListener('click', () => openSettingsModal());
    }
}

// ===== 模式选择 =====
function updateModeSelection() {
    $$('input[name="mode"]').forEach(radio => {
        radio.checked = radio.value === window.__state.currentMode;
    });
}

// ===== 设置弹窗 =====
function openSettingsModal() {
    $('#template-modal')?.classList.add('show');
    populateSettingsForm();
}

function closeSettingsModal() {
    $('#template-modal')?.classList.remove('show');
}

function getEffectiveTemplate() {
    const { currentTemplate, customTemplates, customSettings } = window.__state;
    
    if (currentTemplate.startsWith('custom_') && customTemplates.length) {
        const idx = parseInt(currentTemplate.split('_')[1]);
        if (!isNaN(idx) && customTemplates[idx]) {
            return customTemplates[idx];
        }
    }
    if (currentTemplate === 'custom' && customSettings) {
        return customSettings;
    }
    return DEFAULT_TEMPLATES[currentTemplate] || DEFAULT_TEMPLATES.official;
}

function populateSettingsForm() {
    const tpl = getEffectiveTemplate();
    if (!tpl) return;

    const els = ['headingFont', 'bodyFont', 'headingSize', 'bodySize', 'heading2Font', 'heading2Size', 'lineSpacing', 'lineRule', 'alignment', 'firstLineIndent'];
    els.forEach(id => {
        const el = $(`#${id}`);
        if (el) {
            if (id === 'useIndent') {
                el.checked = (tpl.firstLineIndent > 0);
            } else if (['headingSize', 'bodySize', 'heading2Size', 'firstLineIndent'].includes(id)) {
                el.value = String(tpl[id] || (id === 'firstLineIndent' ? 0 : 12));
            } else if (id === 'lineSpacing') {
                el.value = String(tpl.lineHeight || 1.5);
            } else {
                el.value = tpl[id] || getDefault(id);
            }
        }
    });

    const useIndent = $('#use-indent');
    if (useIndent) useIndent.checked = (tpl.firstLineIndent > 0);
}

function getDefault(id) {
    const defaults = {
        headingFont: 'SimSun',
        bodyFont: 'SimSun',
        headingSize: 16,
        bodySize: 12,
        heading2Font: 'SimSun',
        heading2Size: 14,
        lineRule: 'auto',
        alignment: 'left'
    };
    return defaults[id] || '';
}

function resetSettingsForm() {
    const tpl = DEFAULT_TEMPLATES[window.__state.currentTemplate] || DEFAULT_TEMPLATES.official;

    const els = ['headingFont', 'bodyFont', 'headingSize', 'bodySize', 'heading2Font', 'heading2Size', 'lineSpacing', 'lineRule', 'alignment'];
    els.forEach(id => {
        const el = $(`#${id}`);
        if (el) {
            if (id === 'lineSpacing') {
                el.value = String(tpl.lineHeight);
            } else {
                el.value = tpl[id] || getDefault(id);
            }
        }
    });

    const useIndent = $('#use-indent');
    const firstLineIndent = $('#first-line-indent');
    if (useIndent) useIndent.checked = tpl.firstLineIndent > 0;
    if (firstLineIndent) firstLineIndent.value = String(tpl.firstLineIndent);
}

function saveCurrentSettings() {
    const { currentTemplate, customTemplates } = window.__state;

    const settings = {
        headingFont: $('#heading-font')?.value || 'SimSun',
        headingSize: parseInt($('#heading-size')?.value) || 12,
        heading2Font: $('#heading2-font')?.value || 'SimSun',
        heading2Size: parseInt($('#heading2-size')?.value) || 14,
        bodyFont: $('#body-font')?.value || 'SimSun',
        bodySize: parseInt($('#body-size')?.value) || 12,
        lineHeight: parseFloat($('#line-spacing')?.value) || 1.5,
        lineRule: $('#line-rule')?.value || 'auto',
        alignment: $('#alignment')?.value || 'left',
        firstLineIndent: parseInt($('#first-line-indent')?.value) || 0,
        indentTwips: 0,
        margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
    };

    if (settings.firstLineIndent > 0) {
        settings.indentTwips = settings.firstLineIndent * 160;
    }

    if (currentTemplate.startsWith('custom_')) {
        const idx = parseInt(currentTemplate.split('_')[1]);
        if (!isNaN(idx) && customTemplates[idx]) {
            Object.assign(customTemplates[idx], settings);
        }
    } else if (currentTemplate === 'custom') {
        window.__state.customSettings = settings;
    } else if (DEFAULT_TEMPLATES[currentTemplate]) {
        Object.assign(DEFAULT_TEMPLATES[currentTemplate], {
            headingFont: settings.headingFont,
            headingSize: settings.headingSize,
            heading2Font: settings.heading2Font,
            heading2Size: settings.heading2Size,
            bodyFont: settings.bodyFont,
            bodySize: settings.bodySize,
            lineHeight: settings.lineHeight,
            lineRule: settings.lineRule,
            alignment: settings.alignment,
            firstLineIndent: settings.firstLineIndent,
            indentTwips: settings.indentTwips
        });
    }

    window.__state.customSettings = settings;
    closeSettingsModal();
    saveState();
    renderTemplateGrid();
    updateTemplateSelection();
    updatePreview('#markdown-input', '#preview-content');
    showToast('设置已保存', 'success');
}

// ===== 自定义模板管理 =====
function saveAsCustomTemplate() {
    const name = $('#custom-template-name')?.value.trim();
    if (!name) {
        showToast('请输入模板名称', 'error');
        return;
    }

    const tpl = getEffectiveTemplate();
    const newTpl = {
        name,
        ...JSON.parse(JSON.stringify(tpl))
    };

    window.__state.customTemplates.push(newTpl);
    if (window.__state.customTemplates.length > 10) {
        window.__state.customTemplates.splice(0, 1);
    }

    window.__state.currentTemplate = 'custom_' + (window.__state.customTemplates.length - 1);
    saveState();
    renderTemplateGrid();
    renderCustomTemplateList();
    updateTemplateSelection();
    closeSettingsModal();
    showToast(`模板 "${name}" 已保存`, 'success');
}

function renderCustomTemplateList() {
    const list = $('#template-list');
    if (!list) return;

    if (window.__state.customTemplates.length === 0) {
        list.innerHTML = '<p style="color: var(--text-muted); font-size: 0.85rem; text-align: center; padding: 1rem;">暂无保存的自定义模板</p>';
        return;
    }

    let html = '';
    window.__state.customTemplates.forEach((tpl, idx) => {
        html += `
            <div class="template-list-item">
                <span class="template-list-name">${tpl.name}</span>
                <div class="template-list-actions">
                    <button onclick="loadCustomTemplate(${idx})" title="使用此模板">📋</button>
                    <button class="delete-btn" onclick="deleteCustomTemplate(${idx})" title="删除">🗑️</button>
                </div>
            </div>`;
    });
    list.innerHTML = html;
}

function loadCustomTemplate(idx) {
    window.__state.currentTemplate = 'custom_' + idx;
    saveState();
    renderTemplateGrid();
    updateTemplateSelection();
    updatePreview('#markdown-input', '#preview-content');
    showToast(`已加载模板: ${window.__state.customTemplates[idx].name}`, 'info');
}

function deleteCustomTemplate(idx) {
    const { currentTemplate, customTemplates } = window.__state;
    const name = customTemplates[idx].name;
    if (!confirm(`确定要删除模板 "${name}" 吗？`)) return;

    customTemplates.splice(idx, 1);
    saveState();
    renderTemplateGrid();
    renderCustomTemplateList();

    if (currentTemplate === 'custom_' + idx) {
        window.__state.currentTemplate = 'official';
        updateTemplateSelection();
    }

    if (currentTemplate.startsWith('custom_')) {
        const num = parseInt(currentTemplate.split('_')[1]);
        if (num > idx) {
            window.__state.currentTemplate = 'custom_' + (num - 1);
        } else if (num === idx) {
            window.__state.currentTemplate = 'official';
        }
    }
    saveState();
    updateTemplateSelection();
    updatePreview('#markdown-input', '#preview-content');
    updateTemplatePreview();
    showToast(`模板 "${name}" 已删除`, 'info');
}

// ===== 模板详情弹窗 =====
function showTemplateDetail(key) {
    const tpl = DEFAULT_TEMPLATES[key];
    if (!tpl) return;

    const info = `
📄 模板：${tpl.name}<br>
━━━━━━━━━━━━━━━<br>
<b>一级标题：</b>${tpl.headingFont} · ${tpl.headingSize}pt<br>
<b>二级标题：</b>${tpl.heading2Font} · ${tpl.heading2Size}pt<br>
<b>正文：</b>${tpl.bodyFont} · ${tpl.bodySize}pt<br>
<b>行距：</b>${tpl.lineHeight}${tpl.lineRule === 'exact' ? ' (固定磅值)' : ' (倍数)'}<br>
<b>对齐：</b>${tpl.alignment}<br>
<b>首行缩进：</b>${tpl.firstLineIndent} 字符
    `;

    const modal = $('#template-modal');
    if (!modal) return;

    const header = modal.querySelector('.modal-header h3');
    if (header) header.textContent = `🔍 ${tpl.name} — 模板详情`;

    let body = modal.querySelector('.modal-body-temp');
    if (!body) {
        body = document.createElement('div');
        body.className = 'modal-body-temp';
        modal.querySelector('.modal-content')?.appendChild(body);
    }
    body.innerHTML = `<div style="padding: 1rem; white-space: pre-wrap; line-height: 1.8; font-size: 0.9rem; color: var(--text-primary);">${info}</div>`;

    modal.querySelectorAll('.setting-group, .setting-row, .checkbox-group, .modal-footer, [style*="border-top"]').forEach(el => {
        el.style.display = 'none';
    });

    const footer = modal.querySelector('[style*="border-top"]');
    if (footer) footer.style.display = 'block';

    const saveBtn = $('#save-template-modal');
    if (saveBtn) {
        saveBtn.textContent = '✓ 关闭';
        saveBtn.onclick = () => {
            closeSettingsModal();
            restoreModalFromDetail();
        };
    }

    modal.classList.add('show');
}

function restoreModalFromDetail() {
    const modal = $('#template-modal');
    if (!modal) return;

    modal.querySelectorAll('.setting-group, .setting-row, .checkbox-group').forEach(el => {
        el.style.display = '';
    });
    const footer = modal.querySelector('[style*="border-top"]');
    if (footer) footer.style.display = '';

    const tempBody = modal.querySelector('.modal-body-temp');
    if (tempBody) tempBody.remove();

    const header = modal.querySelector('.modal-header h3');
    if (header) header.textContent = '⚙️ 模板设置';

    const saveBtn = $('#save-template-modal');
    if (saveBtn) {
        saveBtn.textContent = '✅ 应用设置';
        saveBtn.onclick = () => saveCurrentSettings();
    }
}

// ===== 编辑自定义模板 =====
function editCustomTemplate(idx) {
    window.__state.currentTemplate = 'custom_' + idx;
    openSettingsModal();

    const tpl = window.__state.customTemplates[idx];
    if (!tpl) return;

    const setVal = (id, val) => {
        const el = $(`#${id}`);
        if (el) el.value = val;
    };

    setVal('heading-font', tpl.headingFont || 'SimSun');
    setVal('body-font', tpl.bodyFont || 'SimSun');
    setVal('heading-size', tpl.headingSize || 12);
    setVal('body-size', tpl.bodySize || 12);
    setVal('heading2-font', tpl.heading2Font || tpl.headingFont || 'SimSun');
    setVal('heading2-size', tpl.heading2Size || tpl.headingSize || 14);
    setVal('line-spacing', String(tpl.lineHeight || 1.5));
    setVal('line-rule', tpl.lineRule || 'auto');
    setVal('alignment', tpl.alignment || 'left');

    const useIndent = $('#use-indent');
    const firstLineIndent = $('#first-line-indent');
    if (useIndent) useIndent.checked = (tpl.firstLineIndent > 0);
    if (firstLineIndent) firstLineIndent.value = String(tpl.firstLineIndent || 0);
}

// ===== 模板预览效果 =====
function updateTemplatePreview() {
    const previewEl = $('#preview-content');
    if (!previewEl || !previewEl.querySelector('h1')) return;

    const tpl = getEffectiveTemplate();

    const headingFont = tpl.headingFont || 'SimSun';
    const bodyFont = tpl.bodyFont || 'SimSun';
    const headingSize = (tpl.headingSize || 16) + 'pt';
    const bodySize = (tpl.bodySize || 12) + 'pt';

    previewEl.style.fontFamily = `var(--font-serif), ${bodyFont}`;
    previewEl.querySelectorAll('h1, h2, h3').forEach(h => {
        h.style.fontFamily = `var(--font-sans), ${headingFont}`;
        h.style.fontSize = h.tagName === 'H1' ? headingSize :
                          h.tagName === 'H2' ? (tpl.heading2Size || 14) + 'pt' : headingSize;
    });
    previewEl.querySelectorAll('p').forEach(p => {
        p.style.fontSize = bodySize;
        p.style.lineHeight = tpl.lineHeight || 1.5;
        if (tpl.firstLineIndent > 0) {
            p.style.textIndent = `${tpl.firstLineIndent}em`;
        }
    });
}

// ===== 启动 =====
document.addEventListener('DOMContentLoaded', init);

// 暴露全局函数供 HTML onclick 调用
window.showTemplateDetail = showTemplateDetail;
window.editCustomTemplate = editCustomTemplate;
window.loadCustomTemplate = loadCustomTemplate;
window.deleteCustomTemplate = deleteCustomTemplate;
