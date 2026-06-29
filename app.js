// ===== AI转Word助手 — 核心逻辑 (v3.0) =====

const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
    BorderStyle, WidthType, Table, TableRow, TableCell, TableLayoutType,
    UnderlineType, TabStopType, TabStopPosition, SectionType, PageNumber,
    VerticalAlign, PageOrientation } = docx;

// ===== 中文传统字号 =====
const CHINESE_SIZES = [
    { n: '初号', p: 42 }, { n: '小初', p: 36 }, { n: '一号', p: 26 },
    { n: '小一', p: 24 }, { n: '二号', p: 22 }, { n: '小二', p: 18 },
    { n: '三号', p: 16 }, { n: '小三', p: 15 }, { n: '四号', p: 14 },
    { n: '小四', p: 12 }, { n: '五号', p: 10.5 }, { n: '小五', p: 9 }
];

function ptToChinese(pt) {
    for (const s of CHINESE_SIZES) {
        if (Math.abs(s.p - pt) < 0.8) return s.n;
    }
    return `${pt}pt`;
}

function chineseToPt(val) {
    if (!val) return 12;
    const num = parseFloat(val);
    if (!isNaN(num)) return num;
    for (const s of CHINESE_SIZES) {
        if (s.n === val) return s.p;
    }
    return 12;
}

// ===== 内置模板 =====
const DEFAULT_TEMPLATES = {
    official: {
        name: '公文模板',
        icon: '📄',
        desc: '党政机关公文格式 (GB/T 9704)',
        headingFont: 'SimHei',
        headingSize: 16,      // 三号
        bodyFont: 'FangSong',
        bodySize: 16,          // 三号
        heading2Font: 'KaiTi',
        heading2Size: 16,
        lineHeight: 28,        // 28磅固定值
        lineRule: 'exact',     // exact = 磅值, auto = 倍数
        alignment: 'justified',
        firstLineIndent: 2,    // 字符数
        indentTwips: 320,      // 2字符 ≈ 320twips
        margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        spacingBefore: 0,
        spacingAfter: 0
    },
    workplace: {
        name: '职场模板',
        icon: '💼',
        desc: '商务办公文档',
        headingFont: 'MicrosoftYaHei',
        headingSize: 15,       // 小三
        bodyFont: 'MicrosoftYaHei',
        bodySize: 12,          // 小四
        heading2Font: 'MicrosoftYaHei',
        heading2Size: 14,
        lineHeight: 1.5,
        lineRule: 'auto',
        alignment: 'left',
        firstLineIndent: 0,
        indentTwips: 0,
        margins: { top: 1270, right: 1270, bottom: 1270, left: 1270 },
        spacingBefore: 6,
        spacingAfter: 6
    },
    report: {
        name: '报告模板',
        icon: '📊',
        desc: '数据分析报告',
        headingFont: 'SimHei',
        headingSize: 15,       // 小三
        bodyFont: 'SimSun',
        bodySize: 12,          // 小四
        heading2Font: 'SimHei',
        heading2Size: 14,
        lineHeight: 1.5,
        lineRule: 'auto',
        alignment: 'left',
        firstLineIndent: 2,
        indentTwips: 320,
        margins: { top: 1270, right: 1270, bottom: 1270, left: 1270 },
        spacingBefore: 6,
        spacingAfter: 6
    }
};

const MODE_NAMES = {
    purify: '净化',
    rearrange: '重排',
    general: '通用'
};

// ===== 全局状态 =====
let currentTemplate = 'official';
let currentMode = 'purify';
let currentTheme = 'light';
let customSettings = null;
let customTemplates = []; // localStorage 保存的自定义模板
let splitRatio = 0.5; // 左右面板分割比例

// ===== DOM =====
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

const el = {
    themeToggle: $('#theme-toggle'),
    settingsBtn: $('#template-settings'),
    templateModal: $('#template-modal'),
    markdownInput: $('#markdown-input'),
    clearBtn: $('#clear-btn'),
    pasteBtn: $('#paste-btn'),
    uploadBtn: $('#upload-btn'),
    fileInput: $('#file-input'),
    dropZone: $('#drop-zone'),
    exportBtn: $('#export-btn'),
    previewContent: $('#preview-content'),
    loadingOverlay: $('#loading-overlay'),
    customizeBtn: $('#customize-template'),
    statsWords: $('#stat-words'),
    statsChars: $('#stat-chars'),
    statsHeadings: $('#stat-headings'),
    statsTables: $('#stat-tables'),
    templateGrid: $('#template-grid'),
    saveTemplateBtn: $('#save-template'),
    saveTemplateModal: $('#save-template-modal'),
    resetTemplateBtn: $('#reset-template'),
    customTemplateName: $('#custom-template-name'),
    closeSettingsBtn: $('#close-settings'),
    closeTemplatesBtn: $('#close-templates'),
    templateList: $('#template-list'),
    addCustomTemplateBtn: $('#add-custom-template-btn'),
    headingFont: $('#heading-font'),
    bodyFont: $('#body-font'),
    headingSize: $('#heading-size'),
    bodySize: $('#body-size'),
    heading2Font: $('#heading2-font'),
    heading2Size: $('#heading2-size'),
    lineSpacing: $('#line-spacing'),
    lineRule: $('#line-rule'),
    alignment: $('#alignment'),
    firstLineIndent: $('#first-line-indent'),
    useIndent: $('#use-indent'),
    themeIcon: $('#theme-icon'),
    themeCycleBtn: $('#theme-cycle'),
    splitPanel: $('#split-panel'),
    divider: $('#divider')
};

// ===== 初始化 =====
function init() {
    loadState();
    renderTemplateGrid();
    renderCustomTemplateList();
    setupEventListeners();
    applyTheme();
    updateTemplateSelection();
    updateModeSelection();
    updatePreview();
    updateStats();
}

function loadState() {
    try {
        const saved = localStorage.getItem('ai-word-converter-v3');
        if (saved) {
            const s = JSON.parse(saved);
            currentTemplate = s.template || 'official';
            currentMode = s.mode || 'purify';
            currentTheme = s.theme || 'light';
            customSettings = s.customSettings || null;
            customTemplates = s.customTemplates || [];
            splitRatio = s.splitRatio ?? 0.5;
        }
    } catch(e) { console.warn('Load state failed:', e); }
}

function saveState() {
    try {
        localStorage.setItem('ai-word-converter-v3', JSON.stringify({
            template: currentTemplate,
            mode: currentMode,
            theme: currentTheme,
            customSettings,
            customTemplates,
            splitRatio
        }));
    } catch(e) { console.warn('Save state failed:', e); }
}

// ===== 事件绑定 =====
function setupEventListeners() {
    // Theme
    el.themeCycleBtn.addEventListener('click', cycleTheme);

    // Template settings modal
    el.settingsBtn.addEventListener('click', () => openSettingsModal());
    el.closeSettingsBtn.addEventListener('click', () => closeSettingsModal());
    el.templateModal.addEventListener('click', (e) => {
        if (e.target === el.templateModal) closeSettingsModal();
    });

    // Save/Reset settings
    el.saveTemplateBtn?.addEventListener('click', saveCurrentSettings);
    el.saveTemplateModal?.addEventListener('click', saveCurrentSettings);
    el.resetTemplateBtn.addEventListener('click', resetSettingsForm);

    // Template cards
    el.templateGrid.addEventListener('click', (e) => {
        const card = e.target.closest('.template-card');
        if (!card) return;
        selectTemplate(card.dataset.template);
    });

    // Mode selection
    el.modeOptions?.querySelectorAll('input[type="radio"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            currentMode = e.target.value;
            saveState();
            updatePreview();
            updateStats();
        });
    });

    // File input
    el.uploadBtn.addEventListener('click', () => el.fileInput.click());
    el.fileInput.addEventListener('change', handleFileUpload);

    // Drag & Drop
    setupDragDrop();

    // Textarea input
    el.markdownInput.addEventListener('input', () => {
        updatePreview();
        updateStats();
    });

    // Paste interception for purify mode
    el.markdownInput.addEventListener('paste', handlePaste);

    // Export
    el.exportBtn.addEventListener('click', exportToWord);

    // Custom template management
    el.customizeBtn.addEventListener('click', () => {
        currentTemplate = 'custom';
        openSettingsModal();
    });
    el.addCustomTemplateBtn.addEventListener('click', saveAsCustomTemplate);

    // Paste button — 移动端兼容：优先 navigator.clipboard，降级为 prompt
    el.clearBtn?.addEventListener('click', () => {
        el.markdownInput.value = '';
        updatePreview();
        updateStats();
        showToast('已清除内容', 'info');
    });

    el.pasteBtn?.addEventListener('click', async () => {
        const ta = el.markdownInput;
        ta.focus();
        try {
            const text = await navigator.clipboard.readText();
            document.execCommand('insertText', false, text);
        } catch {
            // 降级：弹出输入框让用户粘贴
            const text = prompt('从剪贴板粘贴，请粘贴你的文案：');
            if (text) {
                ta.value = text;
                ta.dispatchEvent(new Event('input'));
            }
        }
    });

    // File input — 移动端放宽 accept，允许所有文件类型
    el.fileInput.removeAttribute('accept');

    // Divider drag
    setupDividerDrag();

    // Keyboard shortcuts
    document.addEventListener('keydown', (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key === 's') {
            e.preventDefault();
            exportToWord();
        }
        if ((e.ctrlKey || e.metaKey) && e.key === '1') {
            currentTheme = 'light'; saveState(); applyTheme();
        }
        if ((e.ctrlKey || e.metaKey) && e.key === '2') {
            currentTheme = 'dark'; saveState(); applyTheme();
        }
        if ((e.ctrlKey || e.metaKey) && e.key === '3') {
            currentTheme = 'warm'; saveState(); applyTheme();
        }
    });
}

// ===== 拖拽上传 =====
function setupDragDrop() {
    const dz = el.dropZone;
    if (!dz) return;

    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(evt => {
        dz.addEventListener(evt, (e) => { e.preventDefault(); e.stopPropagation(); });
    });

    dz.addEventListener('dragover', () => dz.classList.add('drag-over'));
    dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
    dz.addEventListener('drop', (e) => {
        dz.classList.remove('drag-over');
        const files = e.dataTransfer.files;
        if (files.length > 0) processFile(files[0]);
    });

    // Click to upload
    dz.addEventListener('click', () => el.fileInput.click());
}

function handlePaste(e) {
    if (currentMode === 'purify') {
        const clipboardData = e.clipboardData || window.clipboardData;
        const html = clipboardData.getData('text/html');
        const text = clipboardData.getData('text/plain');

        if (html && html.length > text.length) {
            // Parse HTML to clean markdown
            const parsed = parseHtmlToMarkdown(html);
            e.preventDefault();
            document.execCommand('insertText', false, parsed);
            setTimeout(() => {
                updatePreview();
                updateStats();
            }, 50);
        }
    }
}

function parseHtmlToMarkdown(html) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    return htmlToMarkdown(doc.body);
}

function htmlToMarkdown(node) {
    let md = '';
    for (const child of node.childNodes) {
        if (child.nodeType === Node.TEXT_NODE) {
            md += child.textContent;
        } else if (child.tagName === 'STRONG' || child.tagName === 'B') {
            md += `**${htmlToMarkdown(child)}**`;
        } else if (child.tagName === 'EM' || child.tagName === 'I') {
            md += `*${htmlToMarkdown(child)}*`;
        } else if (child.tagName === 'DEL' || child.tagName === 'STRIKE') {
            md += `~~${htmlToMarkdown(child)}~~`;
        } else if (child.tagName === 'CODE') {
            md += `\`${htmlToMarkdown(child)}\``;
        } else if (child.tagName === 'BR') {
            md += '\n';
        } else if (child.tagName === 'P') {
            md += '\n' + htmlToMarkdown(child) + '\n';
        } else if (child.tagName === 'H1') {
            md += `\n# ${htmlToMarkdown(child).trim()}\n`;
        } else if (child.tagName === 'H2') {
            md += `\n## ${htmlToMarkdown(child).trim()}\n`;
        } else if (child.tagName === 'H3') {
            md += `\n### ${htmlToMarkdown(child).trim()}\n`;
        } else if (child.tagName === 'UL') {
            for (const li of child.querySelectorAll(':scope > li')) {
                md += `- ${htmlToMarkdown(li)}\n`;
            }
        } else if (child.tagName === 'OL') {
            let i = 1;
            for (const li of child.querySelectorAll(':scope > li')) {
                md += `${i}. ${htmlToMarkdown(li)}\n`;
                i++;
            }
        } else if (child.tagName === 'TABLE') {
            md += htmlTableToMarkdown(child);
        }
    }
    return md.trim();
}

function htmlTableToMarkdown(table) {
    let md = '\n';
    const rows = table.querySelectorAll('tr');
    rows.forEach((row, i) => {
        const cells = row.querySelectorAll('th, td');
        const cellTexts = Array.from(cells).map(c => c.textContent.trim());
        md += '| ' + cellTexts.join(' | ') + ' |\n';
        if (i === 0) {
            md += '| ' + cellTexts.map(() => '---').join('|') + ' |\n';
        }
    });
    return md;
}

function processFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    const reader = new FileReader();
    reader.onload = (e) => {
        if (ext === 'docx') {
            // Handle docx upload — extract text
            extractDocxText(e.target.result);
        } else {
            // Text/Markdown file
            el.markdownInput.value = e.target.result;
            updatePreview();
            updateStats();
            showToast('文件已加载', 'success');
        }
    };
    if (ext === 'docx') {
        reader.readAsArrayBuffer(file);
    } else {
        reader.readAsText(file);
    }
}

// ===== 主题切换 =====
function cycleTheme() {
    const themes = ['light', 'dark', 'warm'];
    const icons = { light: '🌙', dark: '☀️', warm: '📖' };
    const idx = themes.indexOf(currentTheme);
    currentTheme = themes[(idx + 1) % themes.length];
    saveState();
    applyTheme();
}

function applyTheme() {
    const icons = { light: '🌙', dark: '☀️', warm: '📖' };
    document.documentElement.setAttribute('data-theme', currentTheme);
    el.themeIcon.textContent = icons[currentTheme] || icons.light;
}

// ===== 模板选择 =====
function selectTemplate(name) {
    currentTemplate = name;
    if (name !== 'custom') {
        customSettings = null;
    }
    saveState();
    updateTemplateSelection();
    updatePreview();
    updateStats();
}

function updateTemplateSelection() {
    $$('.template-card').forEach(card => {
        card.classList.toggle('active', card.dataset.template === currentTemplate);
    });
}

function renderTemplateGrid() {
    const grid = el.templateGrid;
    if (!grid) return;

    let html = '';

    // Built-in templates
    Object.entries(DEFAULT_TEMPLATES).forEach(([key, tpl]) => {
        html += `
            <div class="template-card ${currentTemplate === key ? 'active' : ''}" data-template="${key}">
                <div class="template-icon">${tpl.icon}</div>
                <div class="template-name">${tpl.name}</div>
                <div class="template-desc">${tpl.desc}</div>
            </div>`;
    });

    // Custom templates
    customTemplates.forEach((tpl, idx) => {
        html += `
            <div class="template-card ${currentTemplate === 'custom_' + idx ? 'active' : ''}" data-template="custom_${idx}">
                <div class="template-icon">⭐</div>
                <div class="template-name">${tpl.name}</div>
                <div class="template-desc">自定义模板</div>
                <div class="template-custom-badge">已保存</div>
            </div>`;
    });

    // Add new custom template card
    html += `
        <div class="template-card" data-template="custom" id="add-custom-card" style="border-style: dashed; opacity: 0.7;">
            <div class="template-icon">➕</div>
            <div class="template-name">新建自定义</div>
            <div class="template-desc">保存当前设置为新模板</div>
        </div>`;

    grid.innerHTML = html;

    // Re-bind click for the add card
    const addCard = document.getElementById('add-custom-card');
    if (addCard) {
        addCard.addEventListener('click', () => openSettingsModal());
    }
}

// ===== 模式选择 =====
function updateModeSelection() {
    $$('input[name="mode"]').forEach(radio => {
        radio.checked = radio.value === currentMode;
    });
}

// ===== 设置弹窗 =====
function openSettingsModal() {
    el.templateModal.classList.add('show');
    populateSettingsForm();
}

function closeSettingsModal() {
    el.templateModal.classList.remove('show');
}

function populateSettingsForm() {
    const tpl = getEffectiveTemplate();
    if (!tpl) return;

    el.headingFont.value = tpl.headingFont || 'SimSun';
    el.bodyFont.value = tpl.bodyFont || 'SimSun';
    el.headingSize.value = tpl.headingSize || 12;
    el.bodySize.value = tpl.bodySize || 12;
    el.heading2Font.value = tpl.heading2Font || tpl.headingFont || 'SimSun';
    el.heading2Size.value = tpl.heading2Size || tpl.headingSize || 14;
    el.lineSpacing.value = String(tpl.lineHeight || 1.5);
    el.lineRule.value = tpl.lineRule || 'auto';
    el.alignment.value = tpl.alignment || 'left';
    el.useIndent.checked = (tpl.firstLineIndent > 0);
    el.firstLineIndent.value = String(tpl.firstLineIndent || 0);
}

function resetSettingsForm() {
    const tpl = DEFAULT_TEMPLATES[currentTemplate] || DEFAULT_TEMPLATES.official;
    el.headingFont.value = tpl.headingFont;
    el.bodyFont.value = tpl.bodyFont;
    el.headingSize.value = tpl.headingSize;
    el.bodySize.value = tpl.bodySize;
    el.heading2Font.value = tpl.heading2Font;
    el.heading2Size.value = tpl.heading2Size;
    el.lineSpacing.value = String(tpl.lineHeight);
    el.lineRule.value = tpl.lineRule;
    el.alignment.value = tpl.alignment;
    el.useIndent.checked = tpl.firstLineIndent > 0;
    el.firstLineIndent.value = String(tpl.firstLineIndent);
}

function saveCurrentSettings() {
    const settings = {
        headingFont: el.headingFont.value,
        headingSize: parseInt(el.headingSize.value) || 12,
        heading2Font: el.heading2Font.value,
        heading2Size: parseInt(el.heading2Size.value) || 14,
        bodyFont: el.bodyFont.value,
        bodySize: parseInt(el.bodySize.value) || 12,
        lineHeight: parseFloat(el.lineSpacing.value) || 1.5,
        lineRule: el.lineRule.value,
        alignment: el.alignment.value,
        firstLineIndent: parseInt(el.firstLineIndent.value) || 0,
        indentTwips: 0,
        margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
    };

    // Calculate indent twips
    if (settings.firstLineIndent > 0) {
        settings.indentTwips = settings.firstLineIndent * 160;
    }

    // Determine where to save
    if (currentTemplate.startsWith('custom_')) {
        const idx = parseInt(currentTemplate.split('_')[1]);
        if (!isNaN(idx) && customTemplates[idx]) {
            Object.assign(customTemplates[idx], settings);
        }
    } else if (currentTemplate === 'custom') {
        customSettings = settings;
    } else if (DEFAULT_TEMPLATES[currentTemplate]) {
        // Override built-in template
        DEFAULT_TEMPLATES[currentTemplate] = {
            ...DEFAULT_TEMPLATES[currentTemplate],
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
        };
    }

    customSettings = settings;
    closeSettingsModal();
    saveState();
    renderTemplateGrid();
    updateTemplateSelection();
    updatePreview();
    showToast('设置已保存', 'success');
}

function saveAsCustomTemplate() {
    const name = el.customTemplateName.value.trim();
    if (!name) {
        showToast('请输入模板名称', 'error');
        return;
    }

    const tpl = getEffectiveTemplate();
    const newTpl = {
        name,
        ...JSON.parse(JSON.stringify(tpl))
    };

    customTemplates.push(newTpl);
    if (customTemplates.length > 10) customTemplates.splice(0, 1); // Max 10

    currentTemplate = 'custom_' + (customTemplates.length - 1);
    saveState();
    renderTemplateGrid();
    renderCustomTemplateList();
    updateTemplateSelection();
    closeSettingsModal();
    showToast(`模板 "${name}" 已保存`, 'success');
}

function renderCustomTemplateList() {
    if (!el.templateList) return;
    if (customTemplates.length === 0) {
        el.templateList.innerHTML = '<p style="color: var(--text-muted); font-size: 0.85rem; text-align: center; padding: 1rem;">暂无保存的自定义模板</p>';
        return;
    }

    let html = '';
    customTemplates.forEach((tpl, idx) => {
        html += `
            <div class="template-list-item">
                <span class="template-list-name">${tpl.name}</span>
                <div class="template-list-actions">
                    <button onclick="loadCustomTemplate(${idx})" title="使用此模板">📋</button>
                    <button class="delete-btn" onclick="deleteCustomTemplate(${idx})" title="删除">🗑️</button>
                </div>
            </div>`;
    });
    el.templateList.innerHTML = html;
}

function loadCustomTemplate(idx) {
    currentTemplate = 'custom_' + idx;
    saveState();
    renderTemplateGrid();
    updateTemplateSelection();
    updatePreview();
    showToast(`已加载模板: ${customTemplates[idx].name}`, 'info');
}

function deleteCustomTemplate(idx) {
    const name = customTemplates[idx].name;
    customTemplates.splice(idx, 1);
    saveState();
    renderTemplateGrid();
    renderCustomTemplateList();
    if (currentTemplate === 'custom_' + idx) {
        currentTemplate = 'official';
        updateTemplateSelection();
    }
    // Renumber
    if (currentTemplate.startsWith('custom_')) {
        const num = parseInt(currentTemplate.split('_')[1]);
        if (num > idx) {
            currentTemplate = 'custom_' + (num - 1);
        } else if (num === idx) {
            currentTemplate = 'official';
        }
    }
    saveState();
    showToast(`模板 "${name}" 已删除`, 'info');
}

// ===== 文件上传 =====
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    processFile(file);
    e.target.value = '';
}

function extractDocxText(arrayBuffer) {
    // Simple docx text extraction using JSZip-like approach
    // For MVP, we'll use a basic approach: try to read as zip and extract text
    showToast('正在解析文档...', 'info');

    // Use the built-in approach: treat docx as zip, extract word/text.xml
    // Since we can't easily unzip in browser without a library,
    // fall back to reading as text for now
    const reader = new FileReader();
    reader.onload = (e) => {
        // Try to find text content in the ArrayBuffer
        const bytes = new Uint8Array(e.target.result);
        let text = '';

        // Simple heuristic: look for <w:t> tags in the binary
        // This is a basic approach; for production, use a proper ZIP library
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

        if (text) {
            el.markdownInput.value = text;
            updatePreview();
            updateStats();
            showToast('文档已加载', 'success');
        } else {
            showToast('无法读取文档内容，请尝试上传为 .txt 或 .md 文件', 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

// ===== 文本处理 =====
function processMarkdown(text) {
    switch (currentMode) {
        case 'purify': return purifyText(text);
        case 'rearrange': return rearrangeText(text);
        default: return text;
    }
}

function purifyText(text) {
    return text
        .replace(/\n{4,}/g, '\n\n\n')
        .replace(/^#{1,6}\s*/gm, (match) => {
            // Keep heading markers but normalize
            return match;
        })
        .replace(/(?<![#])\n[#!~-]+\s/g, '\n')
        .replace(/\n{3}/g, '\n\n')
        .trim();
}

function rearrangeText(text) {
    return text
        .replace(/\n{4,}/g, '\n\n\n')
        .replace(/•/g, '-')
        .replace(/–/g, '-')
        .replace(/^[\s·\-•]+/gm, '')
        .replace(/\n\s*\n/g, '\n\n')
        .trim();
}

// ===== 预览更新 =====
function updatePreview() {
    const text = el.markdownInput.value.trim();
    if (!text) {
        el.previewContent.innerHTML = `
            <div class="empty-preview">
                <div class="icon">📝</div>
                <p>粘贴AI文案后点击"导出Word"查看效果</p>
                <p style="font-size: 0.8rem; margin-top: 0.5rem; opacity: 0.7;">支持 Markdown 格式：标题、加粗、列表、表格等</p>
            </div>`;
        return;
    }
    const processed = processMarkdown(text);
    try {
        el.previewContent.innerHTML = marked.parse(processed);
    } catch(e) {
        el.previewContent.textContent = processed;
    }
}

// ===== 字数统计 =====
function updateStats() {
    const text = el.markdownInput.value.trim();
    if (!text) {
        el.statsWords.textContent = '0';
        el.statsChars.textContent = '0';
        el.statsHeadings.textContent = '0';
        el.statsTables.textContent = '0';
        return;
    }

    const wordCount = text.replace(/\s/g, '').length;
    const charCount = text.length;
    const headingCount = (text.match(/^#{1,6}\s/gm) || []).length;
    const tableCount = (text.match(/^\|/gm) || []).length;

    el.statsWords.textContent = wordCount;
    el.statsChars.textContent = charCount;
    el.statsHeadings.textContent = headingCount;
    el.statsTables.textContent = Math.floor(tableCount / 2);
}

// ===== 获取有效模板 =====
function getEffectiveTemplate() {
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

// ===== 解析Markdown为docx段落 =====
function parseMarkdownToParagraphs(markdown, template) {
    const paragraphs = [];
    const lines = markdown.split('\n');
    let i = 0;

    while (i < lines.length) {
        const line = lines[i];

        // Empty line
        if (line.trim() === '') {
            i++;
            continue;
        }

        // Horizontal rule
        if (/^(-{3,}|_{3,}|\*{3,})$/.test(line.trim())) {
            paragraphs.push(new Paragraph({
                border: {
                    bottom: { style: BorderStyle.SINGLE, size: 6, color: 'auto' }
                },
                spacing: { before: 120, after: 120 }
            }));
            i++;
            continue;
        }

        // Tables
        if (line.startsWith('|')) {
            const tableResult = parseTable(lines, i, template);
            paragraphs.push(...tableResult.paragraphs);
            i = tableResult.nextIndex;
            continue;
        }

        // Headings
        if (line.startsWith('### ')) {
            paragraphs.push(createHeading(line.slice(4), 3, template));
            i++;
            continue;
        }
        if (line.startsWith('## ')) {
            paragraphs.push(createHeading(line.slice(3), 2, template));
            i++;
            continue;
        }
        if (line.startsWith('# ')) {
            paragraphs.push(createHeading(line.slice(2), 1, template));
            i++;
            continue;
        }

        // Unordered list
        if (/^[-*+]\s/.test(line)) {
            const items = [];
            while (i < lines.length && /^[-*+]\s/.test(lines[i])) {
                items.push(lines[i].replace(/^[-*+]\s/, ''));
                i++;
            }
            paragraphs.push(...createBullets(items, template));
            continue;
        }

        // Ordered list
        if (/^\d+\.\s/.test(line)) {
            const items = [];
            while (i < lines.length && /^\d+\.\s/.test(lines[i])) {
                items.push(lines[i].replace(/^\d+\.\s/, ''));
                i++;
            }
            paragraphs.push(...createNumberedItems(items, template));
            continue;
        }

        // Code block
        if (line.startsWith('```')) {
            const codeLines = [];
            i++; // skip opening ```
            while (i < lines.length && !lines[i].startsWith('```')) {
                codeLines.push(lines[i]);
                i++;
            }
            i++; // skip closing ```
            paragraphs.push(createCodeBlock(codeLines.join('\n'), template));
            continue;
        }

        // Blockquote
        if (line.startsWith('> ')) {
            const quoteLines = [];
            while (i < lines.length && lines[i].startsWith('> ')) {
                quoteLines.push(lines[i].slice(2));
                i++;
            }
            paragraphs.push(createQuote(quoteLines.join('\n'), template));
            continue;
        }

        // Regular paragraph
        const paraLines = [];
        while (i < lines.length &&
               lines[i].trim() !== '' &&
               !lines[i].startsWith('#') &&
               !lines[i].startsWith('|') &&
               !/^[-*+]\s/.test(lines[i]) &&
               !/^\d+\.\s/.test(lines[i]) &&
               !lines[i].startsWith('```') &&
               !lines[i].startsWith('> ')) {
            paraLines.push(lines[i]);
            i++;
        }
        if (paraLines.length > 0) {
            paragraphs.push(createParagraph(paraLines.join('\n'), template));
        }
    }

    return paragraphs;
}

function createHeading(text, level, template) {
    const sizes = {
        1: template.headingSize + 4,
        2: template.heading2Size || template.headingSize + 2,
        3: template.headingSize
    };
    const fonts = {
        1: template.headingFont,
        2: template.heading2Font || template.headingFont,
        3: template.headingFont
    };

    return new Paragraph({
        children: [new TextRun({
            text: text,
            bold: true,
            font: fonts[level],
            size: sizes[level] * 2
        })],
        heading: level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3,
        alignment: level === 1 ? AlignmentType.CENTER : AlignmentType.LEFT,
        spacing: {
            before: level === 1 ? 400 : 300,
            after: 200
        }
    });
}

function createParagraph(text, template) {
    const runs = processInlineFormat(text, template);

    const spacingOpts = {};
    if (template.lineRule === 'exact') {
        // 磅值固定行距
        spacingOpts.line = Math.round(template.lineHeight * 20);
        spacingOpts.lineRule = 'exact';
    } else {
        // 倍数行距
        spacingOpts.line = Math.round(template.lineHeight * 240);
        spacingOpts.lineRule = 'auto';
    }

    if (template.spacingBefore) spacingOpts.before = template.spacingBefore * 60;
    if (template.spacingAfter) spacingOpts.after = template.spacingAfter * 60;

    return new Paragraph({
        children: runs,
        alignment: template.alignment === 'justified' ? AlignmentType.JUSTIFIED :
                   template.alignment === 'center' ? AlignmentType.CENTER :
                   AlignmentType.LEFT,
        spacing: spacingOpts,
        indent: template.firstLineIndent > 0 ? { firstLine: template.indentTwips || 320 } : undefined
    });
}

function processInlineFormat(text, template) {
    const runs = [];

    // Handle bold, italic, strikethrough, code
    const parts = text.split(/(\*\*[^*]+\*\*|__[^_]+__|\*[^*]+\*|_[^_]+_|~~[^~]+~~|`[^`]+`)/g);

    parts.forEach(part => {
        if (!part) return;

        if (/^\*\*.*\*\*$/.test(part) || /^__.*__$/.test(part)) {
            runs.push(new TextRun({
                text: part.replace(/^\*\*|\*\*$|^__|__$/g, ''),
                bold: true,
                font: template.bodyFont,
                size: template.bodySize * 2
            }));
        } else if (/^\*.*\*$/.test(part) || /^_.*_$/.test(part)) {
            runs.push(new TextRun({
                text: part.replace(/^\*|\*|^_|_$/g, ''),
                italic: true,
                font: template.bodyFont,
                size: template.bodySize * 2
            }));
        } else if (/^~~.*~~/.test(part)) {
            runs.push(new TextRun({
                text: part.replace(/^~~|~~$/g, ''),
                strike: true,
                font: template.bodyFont,
                size: template.bodySize * 2
            }));
        } else if (/^`.*`/.test(part)) {
            runs.push(new TextRun({
                text: part.replace(/^`|`$/g, ''),
                font: 'Consolas',
                size: template.bodySize * 2,
                color: 'c7254e'
            }));
        } else {
            runs.push(new TextRun({
                text: part,
                font: template.bodyFont,
                size: template.bodySize * 2
            }));
        }
    });

    return runs.length > 0 ? runs : [new TextRun({ text, font: template.bodyFont, size: template.bodySize * 2 })];
}

function createBullets(items, template) {
    return items.map(item => new Paragraph({
        children: [new TextRun({
            text: item,
            font: template.bodyFont,
            size: template.bodySize * 2
        })],
        bullet: { level: 0 },
        spacing: { after: 60 }
    }));
}

function createNumberedItems(items, template) {
    return items.map((item, idx) => new Paragraph({
        children: [new TextRun({
            text: item,
            font: template.bodyFont,
            size: template.bodySize * 2
        })],
        number: { reference: 'default-numbering', level: 0 },
        spacing: { after: 60 }
    }));
}

function createCodeBlock(code, template) {
    return new Paragraph({
        children: code.split('\n').map(line =>
            new TextRun({
                text: line,
                font: 'Consolas',
                size: (template.bodySize - 2) * 2,
                break: line !== code.split('\n').pop() ? 1 : 0
            })
        ),
        spacing: { before: 120, after: 120 },
        border: {
            top: { style: BorderStyle.SINGLE, size: 1, color: 'cccccc' },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: 'cccccc' },
            left: { style: BorderStyle.SINGLE, size: 1, color: 'cccccc' },
            right: { style: BorderStyle.SINGLE, size: 1, color: 'cccccc' }
        },
        shading: { type: 'clear', fill: 'f5f5f5', color: 'f5f5f5' }
    });
}

function createQuote(text, template) {
    return new Paragraph({
        children: [new TextRun({
            text: '「' + text + '」',
            italic: true,
            font: template.bodyFont,
            size: template.bodySize * 2,
            color: '666666'
        })],
        spacing: { before: 60, after: 60 },
        indent: { left: 360 }
    });
}

function parseTable(lines, startIndex, template) {
    const tableRows = [];
    let i = startIndex;

    while (i < lines.length && lines[i].startsWith('|')) {
        const cells = lines[i].split('|').filter((c, idx, arr) => idx > 0 && idx < arr.length - 1);
        tableRows.push(cells.map(c => c.trim()));
        i++;
    }

    if (tableRows.length < 2) {
        return { paragraphs: [], nextIndex: i };
    }

    const nc = Math.max(...tableRows.map(r => r.length));
    const cw = Math.floor(9000 / nc);

    const tableCells = tableRows.map((row, rowIdx) => {
        const cells = row.map(cellText => {
            const runs = [new TextRun({
                text: cellText || '',
                font: rowIdx === 0 ? template.headingFont : template.bodyFont,
                size: rowIdx === 0 ? template.headingSize * 2 : template.bodySize * 2,
                bold: rowIdx === 0
            })];

            return new TableCell({
                children: [new Paragraph({ children: runs })],
                width: { size: cw, type: WidthType.DXA },
                shading: {
                    type: 'clear',
                    fill: rowIdx === 0 ? 'e8e8e8' : 'ffffff',
                    color: rowIdx === 0 ? 'e8e8e8' : 'ffffff'
                }
            });
        });

        // Pad missing cells
        while (cells.length < nc) {
            cells.push(new TableCell({
                children: [new Paragraph({ children: [new TextRun({ text: '', font: template.bodyFont, size: template.bodySize * 2 })] })],
                width: { size: cw, type: WidthType.DXA }
            }));
        }

        return new TableRow({ children: cells });
    });

    const table = new Table({
        children: tableCells,
        width: { size: 9000, type: WidthType.DXA },
        layout: TableLayoutType.FIXED,
        tableVerticalAlign: 'top'
    });

    return { paragraphs: [table], nextIndex: i };
}

// ===== 导出Word =====
async function exportToWord() {
    const text = el.markdownInput.value.trim();
    if (!text) {
        showToast('请先粘贴或上传AI文案', 'error');
        return;
    }

    el.loadingOverlay.classList.add('show');

    try {
        const processed = processMarkdown(text);
        const template = getEffectiveTemplate();
        const children = parseMarkdownToParagraphs(processed, template);

        if (children.length === 0) {
            showToast('没有可导出的内容', 'error');
            return;
        }

        const modeLabel = MODE_NAMES[currentMode] || currentMode;
        const tplName = template.name || '自定义';
        const dateStr = new Date().toLocaleDateString('zh-CN').replace(/\//g, '');
        const fileName = `AI转Word_${modeLabel}模式_${tplName}_${dateStr}.docx`;

        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        width: 11906,
                        height: 16838,
                        margin: template.margins || { top: 1440, right: 1440, bottom: 1440, left: 1440 }
                    }
                },
                children: children
            }]
        });

        const blob = await Packer.toBlob(doc);
        downloadBlob(blob, fileName);
        showToast(`✅ 导出成功！文件: ${fileName}`, 'success');
    } catch (error) {
        console.error('Export failed:', error);
        showToast(`❌ 导出失败: ${error.message}`, 'error');
    } finally {
        el.loadingOverlay.classList.remove('show');
    }
}

function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// ===== 分割面板拖拽 =====
function setupDividerDrag() {
    if (!el.divider) return;

    let isDragging = false;

    el.divider.addEventListener('mousedown', (e) => {
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
        splitRatio = Math.max(0.25, Math.min(0.75, ratio));
        container.style.gridTemplateColumns = `${splitRatio * 100}% ${((1 - splitRatio) * 100).toFixed(2)}%`;
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

// ===== Toast 通知 =====
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

// ===== 启动 =====
init();
