// ===== AI转Word助手 - 主应用逻辑 =====

const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
    BorderStyle, WidthType, Table, TableRow, TableCell, TableLayoutType,
    UnderlineType, TabStopType, TabStopPosition, SectionType, PageNumber } = docx;

// ===== 模板配置 =====
const TEMPLATES = {
    official: {
        name: '公文模板',
        headingFont: 'SimHei',
        headingSize: 22,
        bodyFont: 'FangSong',
        bodySize: 16,
        lineHeight: 560,
        alignment: AlignmentType.JUSTIFIED,
        firstLineIndent: 432,
        margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
    },
    workplace: {
        name: '职场模板',
        headingFont: 'MicrosoftYaHei',
        headingSize: 18,
        bodyFont: 'MicrosoftYaHei',
        bodySize: 14,
        lineHeight: 420,
        alignment: AlignmentType.LEFT,
        firstLineIndent: 0,
        margins: { top: 1270, right: 1270, bottom: 1270, left: 1270 }
    },
    report: {
        name: '报告模板',
        headingFont: 'SimHei',
        headingSize: 20,
        bodyFont: 'SimSun',
        bodySize: 14,
        lineHeight: 480,
        alignment: AlignmentType.LEFT,
        firstLineIndent: 432,
        margins: { top: 1270, right: 1270, bottom: 1270, left: 1270 }
    },
    custom: null
};

const MODE_NAMES = {
    purify: '净化',
    rearrange: '重排',
    general: '通用'
};

let currentTemplate = 'official';
let currentMode = 'purify';
let currentTheme = 'light';
let customSettings = null;

// ===== DOM元素 =====
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

const elements = {
    themeToggle: $('#theme-toggle'),
    templateSettings: $('#template-settings'),
    templateModal: $('#template-modal'),
    markdownInput: $('#markdown-input'),
    pasteBtn: $('#paste-btn'),
    uploadBtn: $('#upload-btn'),
    fileInput: $('#file-input'),
    exportBtn: $('#export-btn'),
    previewContent: $('#preview-content'),
    loadingOverlay: $('#loading-overlay'),
    customizeTemplate: $('#customize-template'),
    resetTemplate: $('#reset-template'),
    saveTemplate: $('#save-template'),
    templateCards: $$('.template-card'),
    modeOptions: $$('.mode-option input'),
    headingFont: $('#heading-font'),
    bodyFont: $('#body-font'),
    headingSize: $('#heading-size'),
    bodySize: $('#body-size'),
    lineSpacing: $('#line-spacing')
};

// ===== 初始化 =====
function init() {
    loadFromLocalStorage();
    setupEventListeners();
    updateTheme();
    updateTemplateSelection();
    updateModeSelection();
    updatePreview();
}

function loadFromLocalStorage() {
    try {
        const saved = localStorage.getItem('ai-word-converter-settings');
        if (saved) {
            const settings = JSON.parse(saved);
            currentTemplate = settings.template || 'official';
            currentMode = settings.mode || 'purify';
            currentTheme = settings.theme || 'light';
            customSettings = settings.customSettings || null;
        }
    } catch(e) {}
}

function saveToLocalStorage() {
    localStorage.setItem('ai-word-converter-settings', JSON.stringify({
        template: currentTemplate,
        mode: currentMode,
        theme: currentTheme,
        customSettings: customSettings
    }));
}

// ===== 事件监听 =====
function setupEventListeners() {
    elements.themeToggle.addEventListener('click', toggleTheme);
    elements.templateSettings.addEventListener('click', showModal);
    $('.close-btn').addEventListener('click', hideModal);
    elements.templateModal.addEventListener('click', (e) => {
        if (e.target === elements.templateModal) hideModal();
    });

    elements.saveTemplate.addEventListener('click', saveTemplateSettings);
    elements.resetTemplate.addEventListener('click', resetTemplateSettings);

    elements.templateCards.forEach(card => {
        card.addEventListener('click', () => {
            elements.templateCards.forEach(c => c.classList.remove('active'));
            card.classList.add('active');
            currentTemplate = card.dataset.template;
            if (currentTemplate !== 'custom') customSettings = null;
            saveToLocalStorage();
            updatePreview();
        });
    });

    elements.modeOptions.forEach(radio => {
        radio.addEventListener('change', (e) => {
            currentMode = e.target.value;
            saveToLocalStorage();
            updatePreview();
        });
    });

    elements.pasteBtn.addEventListener('click', async () => {
        try {
            const text = await navigator.clipboard.readText();
            elements.markdownInput.value = text;
            updatePreview();
        } catch (err) {
            alert('无法读取剪贴板，请手动粘贴');
        }
    });

    elements.uploadBtn.addEventListener('click', () => elements.fileInput.click());
    elements.fileInput.addEventListener('change', handleFileUpload);
    elements.markdownInput.addEventListener('input', updatePreview);
    elements.exportBtn.addEventListener('click', exportToWord);
    elements.customizeTemplate.addEventListener('click', () => {
        currentTemplate = 'custom';
        elements.templateCards.forEach(c => c.classList.remove('active'));
        $('[data-template="custom"]').classList.add('active');
        showModal();
    });
}

// ===== 主题切换 =====
function toggleTheme() {
    const themes = ['light', 'dark', 'warm'];
    const icons = { light: '🌙', dark: '☀️', warm: '📖' };
    const idx = themes.indexOf(currentTheme);
    currentTheme = themes[(idx + 1) % themes.length];
    applyTheme();
    saveToLocalStorage();
}

function applyTheme() {
    const icons = { light: '🌙', dark: '☀️', warm: '📖' };
    document.documentElement.removeAttribute('data-theme');
    if (currentTheme !== 'light') {
        document.documentElement.setAttribute('data-theme', currentTheme);
    }
    elements.themeToggle.textContent = icons[currentTheme] || icons.light;
}

function updateTheme() { applyTheme(); }

// ===== 弹窗控制 =====
function showModal() {
    elements.templateModal.classList.add('show');
    const tpl = currentTemplate === 'custom' && customSettings
        ? customSettings : TEMPLATES[currentTemplate];
    if (tpl) {
        elements.headingFont.value = tpl.headingFont || 'SimSun';
        elements.bodyFont.value = tpl.bodyFont || 'SimSun';
        elements.headingSize.value = tpl.headingSize || 16;
        elements.bodySize.value = tpl.bodySize || 12;
        elements.lineSpacing.value = tpl.lineHeight ? (tpl.lineHeight / 240).toFixed(1) : '2.0';
    }
}

function hideModal() { elements.templateModal.classList.remove('show'); }

function saveTemplateSettings() {
    customSettings = {
        headingFont: elements.headingFont.value,
        bodyFont: elements.bodyFont.value,
        headingSize: parseInt(elements.headingSize.value),
        bodySize: parseInt(elements.bodySize.value),
        lineHeight: parseFloat(elements.lineSpacing.value) * 240
    };
    currentTemplate = 'custom';
    elements.templateCards.forEach(c => c.classList.remove('active'));
    $('[data-template="custom"]').classList.add('active');
    hideModal();
    saveToLocalStorage();
    updatePreview();
}

function resetTemplateSettings() {
    elements.headingFont.value = 'SimSun';
    elements.bodyFont.value = 'SimSun';
    elements.headingSize.value = 16;
    elements.bodySize.value = 12;
    elements.lineSpacing.value = '2.0';
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
        .replace(/\n{3,}/g, '\n\n')
        .replace(/^[#!~-]+ /gm, '')
        .replace(/\*\*/g, '')
        .replace(/`/g, '')
        .trim();
}

function rearrangeText(text) {
    return text
        .replace(/\n{3,}/g, '\n\n')
        .replace(/•/g, '-')
        .replace(/^\s+|\s+$/gm, '')
        .trim();
}

// ===== 预览 =====
function updatePreview() {
    const text = elements.markdownInput.value.trim();
    if (!text) {
        elements.previewContent.innerHTML = '<div class="empty-preview"><p>粘贴AI文案后点击"导出Word"查看效果</p></div>';
        return;
    }
    const processed = processMarkdown(text);
    try {
        elements.previewContent.innerHTML = marked.parse(processed);
    } catch(e) {
        elements.previewContent.textContent = processed;
    }
}

// ===== 文件上传 =====
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
        elements.markdownInput.value = e.target.result;
        updatePreview();
    };
    reader.readAsText(file);
    event.target.value = '';
}

// ===== 获取模板设置 =====
function getTemplateSettings() {
    if (currentTemplate === 'custom' && customSettings) {
        return {
            headingFont: customSettings.headingFont,
            headingSize: customSettings.headingSize,
            bodyFont: customSettings.bodyFont,
            bodySize: customSettings.bodySize,
            lineHeight: customSettings.lineHeight || 480,
            alignment: AlignmentType.LEFT,
            firstLineIndent: 0,
            margins: { top: 1270, right: 1270, bottom: 1270, left: 1270 }
        };
    }
    return TEMPLATES[currentTemplate] || TEMPLATES.official;
}

// ===== 解析Markdown为docx段落 =====
function parseMarkdownToParagraphs(markdown, template) {
    const paragraphs = [];
    const lines = markdown.split('\n');
    let i = 0;

    while (i < lines.length) {
        const line = lines[i];

        if (line.startsWith('### ')) {
            paragraphs.push(createHeading(line.slice(4), 3, template));
        } else if (line.startsWith('## ')) {
            paragraphs.push(createHeading(line.slice(3), 2, template));
        } else if (line.startsWith('# ')) {
            paragraphs.push(createHeading(line.slice(2), 1, template));
        } else if (line.match(/^[-*+] /)) {
            const items = [];
            while (i < lines.length && lines[i].match(/^[-*+] /)) {
                items.push(lines[i].replace(/^[-*+] /, ''));
                i++;
            }
            paragraphs.push(...createBullets(items, template));
            continue;
        } else if (line.match(/^\d+\.\s/)) {
            const items = [];
            while (i < lines.length && lines[i].match(/^\d+\.\s/)) {
                items.push(lines[i].replace(/^\d+\.\s/, ''));
                i++;
            }
            paragraphs.push(...createNumberedItems(items, template));
            continue;
        } else if (line.trim() === '---' || line.trim() === '***') {
            paragraphs.push(new Paragraph({
                borderBottom: { style: BorderStyle.SINGLE, size: 6, color: 'auto' },
                spacing: { before: 120, after: 120 }
            }));
            i++;
            continue;
        } else if (line.trim() === '') {
            i++;
            continue;
        } else {
            const paraLines = [];
            while (i < lines.length && lines[i].trim() !== '' &&
                   !lines[i].startsWith('#') && !lines[i].match(/^[-*+] /) &&
                   !lines[i].match(/^\d+\.\s/) && lines[i].trim() !== '---') {
                paraLines.push(lines[i]);
                i++;
            }
            if (paraLines.length > 0) {
                paragraphs.push(createParagraph(paraLines.join('\n'), template));
            }
            continue;
        }
        i++;
    }
    return paragraphs;
}

function createHeading(text, level, template) {
    const sizes = { 1: template.headingSize + 6, 2: template.headingSize + 2, 3: template.headingSize };
    return new Paragraph({
        children: [new TextRun({
            text: text,
            bold: true,
            font: template.headingFont,
            size: sizes[level] * 2,
            color: '000000'
        })],
        heading: level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3,
        alignment: AlignmentType.CENTER,
        spacing: { before: 300, after: 200 }
    });
}

function createParagraph(text, template) {
    const runs = processInlineFormat(text, template);
    return new Paragraph({
        children: runs,
        alignment: template.alignment,
        spacing: { line: template.lineHeight, after: 100 },
        indent: template.firstLineIndent > 0 ? { firstLine: template.firstLineIndent } : undefined
    });
}

function processInlineFormat(text, template) {
    const runs = [];
    const parts = text.split(/(\*\*[^*]+\*\*)/g);
    parts.forEach(part => {
        if (part.startsWith('**') && part.endsWith('**')) {
            runs.push(new TextRun({
                text: part.slice(2, -2),
                bold: true,
                font: template.bodyFont,
                size: template.bodySize * 2
            }));
        } else {
            runs.push(new TextRun({
                text: part,
                font: template.bodyFont,
                size: template.bodySize * 2
            }));
        }
    });
    return runs;
}

function createBullets(items, template) {
    return items.map(item => new Paragraph({
        children: [new TextRun({
            text: item,
            font: template.bodyFont,
            size: template.bodySize * 2
        })],
        spacing: { after: 80 },
        bullet: { level: 0 }
    }));
}

function createNumberedItems(items, template) {
    return items.map(item => new Paragraph({
        children: [new TextRun({
            text: item,
            font: template.bodyFont,
            size: template.bodySize * 2
        })],
        spacing: { after: 80 },
        number: { reference: 1, level: 0 }
    }));
}

// ===== 导出Word =====
async function exportToWord() {
    const text = elements.markdownInput.value.trim();
    if (!text) {
        alert('请先粘贴或上传AI文案');
        return;
    }

    elements.loadingOverlay.style.display = 'flex';

    try {
        const processed = processMarkdown(text);
        const template = getTemplateSettings();
        const children = parseMarkdownToParagraphs(processed, template);

        const modeLabel = MODE_NAMES[currentMode] || currentMode;
        const templateName = template.name || '自定义';
        const fileName = `AI转Word_${modeLabel}模式_${templateName}_${new Date().toLocaleDateString('zh-CN')}.docx`;

        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        width: 11906,
                        height: 16838,
                        margin: template.margins
                    }
                },
                children: children
            }]
        });

        const blob = await Packer.toBlob(doc);
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        alert('✅ Word文档导出成功！\n文件名：' + fileName);
    } catch (error) {
        console.error('导出失败:', error);
        alert('❌ 导出失败: ' + error.message);
    } finally {
        elements.loadingOverlay.style.display = 'none';
    }
}

// ===== 辅助函数 =====
function updateTemplateSelection() {
    elements.templateCards.forEach(card => {
        card.classList.remove('active');
        if (card.dataset.template === currentTemplate) card.classList.add('active');
    });
}

function updateModeSelection() {
    elements.modeOptions.forEach(radio => {
        if (radio.value === currentMode) radio.checked = true;
    });
}

// ===== 启动 =====
init();
