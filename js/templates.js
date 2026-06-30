// ===== AI转Word助手 — 模板管理 =====

import { DEFAULT_TEMPLATES } from './constants.js';
import { state, getState, setCurrentTemplateAndSave, saveState } from './state.js';
import { showToast } from './ui.js';

/**
 * 获取有效模板
 * @returns {Object} 当前有效的模板配置
 */
function getEffectiveTemplate() {
    const { currentTemplate, customTemplates, customSettings } = getState();
    
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

// ===== 模板选择 =====
function selectTemplate(name) {
    setCurrentTemplateAndSave(name);
    if (name !== 'custom') {
        getState().customSettings = null;
    }
    saveState();
}

function updateTemplateSelection() {
    const currentTemplate = getState().currentTemplate;
    document.querySelectorAll('.template-card').forEach(card => {
        card.classList.toggle('active', card.dataset.template === currentTemplate);
    });
}

// ===== 模板渲染 =====
function renderTemplateGrid() {
    const grid = document.querySelector('#template-grid');
    if (!grid) return;

    const { currentTemplate, customTemplates } = getState();
    let html = '';

    // Built-in templates (view-only, click to see detail)
    Object.entries(DEFAULT_TEMPLATES).forEach(([key, tpl]) => {
        html += `
            <div class="template-card ${currentTemplate === key ? 'active' : ''}" data-template="${key}">
                <div class="template-icon">${tpl.icon}</div>
                <div class="template-name">${tpl.name}</div>
                <div class="template-desc">${tpl.desc}</div>
                <div class="template-card-actions">
                    <button class="card-action-btn detail-btn" onclick="event.stopPropagation(); showTemplateDetail('${key}')" title="查看详情">🔍</button>
                </div>
            </div>`;
    });

    // Custom templates (with edit/delete)
    customTemplates.forEach((tpl, idx) => {
        html += `
            <div class="template-card ${currentTemplate === 'custom_' + idx ? 'active' : ''}" data-template="custom_${idx}">
                <div class="template-icon">⭐</div>
                <div class="template-name">${tpl.name}</div>
                <div class="template-desc">自定义模板</div>
                <div class="template-custom-badge">已保存</div>
                <div class="template-card-actions">
                    <button class="card-action-btn edit-btn" onclick="event.stopPropagation(); editCustomTemplate(${idx})" title="编辑">✏️</button>
                    <button class="card-action-btn delete-btn" onclick="event.stopPropagation(); deleteCustomTemplate(${idx})" title="删除">🗑️</button>
                </div>
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
    const { currentMode } = getState();
    document.querySelectorAll('input[name="mode"]').forEach(radio => {
        radio.checked = radio.value === currentMode;
    });
}

// ===== 设置弹窗 =====
function openSettingsModal() {
    const modal = document.querySelector('#template-modal');
    if (modal) {
        modal.classList.add('show');
        populateSettingsForm();
    }
}

function closeSettingsModal() {
    const modal = document.querySelector('#template-modal');
    if (modal) {
        modal.classList.remove('show');
    }
}

function populateSettingsForm() {
    const tpl = getEffectiveTemplate();
    if (!tpl) return;

    const els = {
        headingFont: document.querySelector('#heading-font'),
        bodyFont: document.querySelector('#body-font'),
        headingSize: document.querySelector('#heading-size'),
        bodySize: document.querySelector('#body-size'),
        heading2Font: document.querySelector('#heading2-font'),
        heading2Size: document.querySelector('#heading2-size'),
        lineSpacing: document.querySelector('#line-spacing'),
        lineRule: document.querySelector('#line-rule'),
        alignment: document.querySelector('#alignment'),
        useIndent: document.querySelector('#use-indent'),
        firstLineIndent: document.querySelector('#first-line-indent')
    };

    if (els.headingFont) els.headingFont.value = tpl.headingFont || 'SimSun';
    if (els.bodyFont) els.bodyFont.value = tpl.bodyFont || 'SimSun';
    if (els.headingSize) els.headingSize.value = tpl.headingSize || 12;
    if (els.bodySize) els.bodySize.value = tpl.bodySize || 12;
    if (els.heading2Font) els.heading2Font.value = tpl.heading2Font || tpl.headingFont || 'SimSun';
    if (els.heading2Size) els.heading2Size.value = tpl.heading2Size || tpl.headingSize || 14;
    if (els.lineSpacing) els.lineSpacing.value = String(tpl.lineHeight || 1.5);
    if (els.lineRule) els.lineRule.value = tpl.lineRule || 'auto';
    if (els.alignment) els.alignment.value = tpl.alignment || 'left';
    if (els.useIndent) els.useIndent.checked = (tpl.firstLineIndent > 0);
    if (els.firstLineIndent) els.firstLineIndent.value = String(tpl.firstLineIndent || 0);
}

function resetSettingsForm() {
    const { currentTemplate } = getState();
    const tpl = DEFAULT_TEMPLATES[currentTemplate] || DEFAULT_TEMPLATES.official;

    const els = {
        headingFont: document.querySelector('#heading-font'),
        bodyFont: document.querySelector('#body-font'),
        headingSize: document.querySelector('#heading-size'),
        bodySize: document.querySelector('#body-size'),
        heading2Font: document.querySelector('#heading2-font'),
        heading2Size: document.querySelector('#heading2-size'),
        lineSpacing: document.querySelector('#line-spacing'),
        lineRule: document.querySelector('#line-rule'),
        alignment: document.querySelector('#alignment'),
        useIndent: document.querySelector('#use-indent'),
        firstLineIndent: document.querySelector('#first-line-indent')
    };

    if (els.headingFont) els.headingFont.value = tpl.headingFont;
    if (els.bodyFont) els.bodyFont.value = tpl.bodyFont;
    if (els.headingSize) els.headingSize.value = tpl.headingSize;
    if (els.bodySize) els.bodySize.value = tpl.bodySize;
    if (els.heading2Font) els.heading2Font.value = tpl.heading2Font;
    if (els.heading2Size) els.heading2Size.value = tpl.heading2Size;
    if (els.lineSpacing) els.lineSpacing.value = String(tpl.lineHeight);
    if (els.lineRule) els.lineRule.value = tpl.lineRule;
    if (els.alignment) els.alignment.value = tpl.alignment;
    if (els.useIndent) els.useIndent.checked = tpl.firstLineIndent > 0;
    if (els.firstLineIndent) els.firstLineIndent.value = String(tpl.firstLineIndent);
}

function saveCurrentSettings() {
    const { currentTemplate, customTemplates } = getState();
    
    const els = {
        headingFont: document.querySelector('#heading-font'),
        headingSize: document.querySelector('#heading-size'),
        heading2Font: document.querySelector('#heading2-font'),
        heading2Size: document.querySelector('#heading2-size'),
        bodyFont: document.querySelector('#body-font'),
        bodySize: document.querySelector('#body-size'),
        lineSpacing: document.querySelector('#line-spacing'),
        lineRule: document.querySelector('#line-rule'),
        alignment: document.querySelector('#alignment'),
        firstLineIndent: document.querySelector('#first-line-indent')
    };

    const settings = {
        headingFont: els.headingFont?.value || 'SimSun',
        headingSize: parseInt(els.headingSize?.value) || 12,
        heading2Font: els.heading2Font?.value || 'SimSun',
        heading2Size: parseInt(els.heading2Size?.value) || 14,
        bodyFont: els.bodyFont?.value || 'SimSun',
        bodySize: parseInt(els.bodySize?.value) || 12,
        lineHeight: parseFloat(els.lineSpacing?.value) || 1.5,
        lineRule: els.lineRule?.value || 'auto',
        alignment: els.alignment?.value || 'left',
        firstLineIndent: parseInt(els.firstLineIndent?.value) || 0,
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
        setState.customSettings = settings;
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

    setState.customSettings = settings;
    closeSettingsModal();
    saveState();
    renderTemplateGrid();
    updateTemplateSelection();
    showToast('设置已保存', 'success');
}

// ===== 自定义模板管理 =====
function saveAsCustomTemplate() {
    const nameInput = document.querySelector('#custom-template-name');
    const name = nameInput?.value.trim();
    if (!name) {
        showToast('请输入模板名称', 'error');
        return;
    }

    const tpl = getEffectiveTemplate();
    const { customTemplates } = getState();
    
    const newTpl = {
        name,
        ...JSON.parse(JSON.stringify(tpl))
    };

    customTemplates.push(newTpl);
    if (customTemplates.length > 10) customTemplates.splice(0, 1); // Max 10

    setCurrentTemplateAndSave('custom_' + (customTemplates.length - 1));
    saveState();
    renderTemplateGrid();
    renderCustomTemplateList();
    updateTemplateSelection();
    closeSettingsModal();
    showToast(`模板 "${name}" 已保存`, 'success');
}

function renderCustomTemplateList() {
    const list = document.querySelector('#template-list');
    if (!list) return;

    const { customTemplates } = getState();
    if (customTemplates.length === 0) {
        list.innerHTML = '<p style="color: var(--text-muted); font-size: 0.85rem; text-align: center; padding: 1rem;">暂无保存的自定义模板</p>';
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
    list.innerHTML = html;
}

function loadCustomTemplate(idx) {
    setCurrentTemplateAndSave('custom_' + idx);
    saveState();
    renderTemplateGrid();
    updateTemplateSelection();
    showToast(`已加载模板: ${getState().customTemplates[idx].name}`, 'info');
}

function deleteCustomTemplate(idx) {
    const { customTemplates, currentTemplate } = getState();
    const name = customTemplates[idx].name;
    if (!confirm(`确定要删除模板 "${name}" 吗？`)) return;
    
    customTemplates.splice(idx, 1);
    saveState();
    renderTemplateGrid();
    renderCustomTemplateList();
    
    if (currentTemplate === 'custom_' + idx) {
        setCurrentTemplateAndSave('official');
        updateTemplateSelection();
    }
    
    // Renumber
    if (currentTemplate.startsWith('custom_')) {
        const num = parseInt(currentTemplate.split('_')[1]);
        if (num > idx) {
            setCurrentTemplateAndSave('custom_' + (num - 1));
        } else if (num === idx) {
            setCurrentTemplateAndSave('official');
        }
    }
    saveState();
    updateTemplateSelection();
    showToast(`模板 "${name}" 已删除`, 'info');
}

// ===== 模板详情弹窗 =====
function showTemplateDetail(key) {
    const tpl = DEFAULT_TEMPLATES[key];
    if (!tpl) return;
    
    const cfg = tpl.config || {};
    const info = `
📄 模板：${tpl.name}<br>
━━━━━━━━━━━━━━━<br>
<b>一级标题：</b>${cfg.headingFont || tpl.bodyFont} · ${cfg.headingSize || 16}pt<br>
<b>二级标题：</b>${cfg.heading2Font || cfg.headingFont || tpl.bodyFont} · ${cfg.heading2Size || 14}pt<br>
<b>正文：</b>${cfg.bodyFont || tpl.bodyFont} · ${cfg.bodySize || 12}pt<br>
<b>行距：</b>${cfg.lineHeight || 1.5}${cfg.lineRule === 'exact' ? ' (固定磅值)' : ' (倍数)'}<br>
<b>对齐：</b>${cfg.alignment || 'left'}<br>
<b>首行缩进：</b>${(cfg.firstLineIndent || 0)} 字符<br>
<b>间距：</b>段前${cfg.spacingBefore || 0}pt / 段后${cfg.spacingAfter || 0}pt
    `;
    
    const modal = document.querySelector('#template-modal');
    if (modal) {
        const header = modal.querySelector('.modal-header h3');
        if (header) header.textContent = `🔍 ${tpl.name} — 模板详情`;
        
        const body = modal.querySelector('.modal-body-temp') || document.createElement('div');
        body.className = 'modal-body-temp';
        body.innerHTML = `<div style="padding: 1rem; white-space: pre-wrap; line-height: 1.8; font-size: 0.9rem; color: var(--text-primary);">${info}</div>`;
        
        modal.querySelectorAll('.setting-group, .setting-row, .checkbox-group, .modal-footer, [style*="border-top"]').forEach(el => {
            el.style.display = 'none';
        });
        
        if (!modal.querySelector('.modal-body-temp')) {
            modal.querySelector('.modal-content').appendChild(body);
        }
        
        const footer = modal.querySelector('[style*="border-top"]');
        if (footer) footer.style.display = 'block';
        
        const saveBtn = document.querySelector('#save-template-modal');
        if (saveBtn) {
            saveBtn.textContent = '✓ 关闭';
            saveBtn.onclick = () => {
                closeSettingsModal();
                restoreModalFromDetail();
            };
        }
        
        modal.classList.add('show');
    }
}

function restoreModalFromDetail() {
    const modal = document.querySelector('#template-modal');
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
    
    const saveBtn = document.querySelector('#save-template-modal');
    if (saveBtn) {
        saveBtn.textContent = '✅ 应用设置';
        saveBtn.onclick = () => saveCurrentSettings();
    }
}

// ===== 编辑自定义模板 =====
function editCustomTemplate(idx) {
    const { currentTemplate, customTemplates } = getState();
    setCurrentTemplateAndSave('custom_' + idx);
    openSettingsModal();
    
    const tpl = customTemplates[idx];
    if (tpl) {
        const els = {
            headingFont: document.querySelector('#heading-font'),
            bodyFont: document.querySelector('#body-font'),
            headingSize: document.querySelector('#heading-size'),
            bodySize: document.querySelector('#body-size'),
            heading2Font: document.querySelector('#heading2-font'),
            heading2Size: document.querySelector('#heading2-size'),
            lineSpacing: document.querySelector('#line-spacing'),
            lineRule: document.querySelector('#line-rule'),
            alignment: document.querySelector('#alignment'),
            useIndent: document.querySelector('#use-indent'),
            firstLineIndent: document.querySelector('#first-line-indent')
        };

        if (els.headingFont) els.headingFont.value = tpl.headingFont || 'SimSun';
        if (els.bodyFont) els.bodyFont.value = tpl.bodyFont || 'SimSun';
        if (els.headingSize) els.headingSize.value = tpl.headingSize || 12;
        if (els.bodySize) els.bodySize.value = tpl.bodySize || 12;
        if (els.heading2Font) els.heading2Font.value = tpl.heading2Font || tpl.headingFont || 'SimSun';
        if (els.heading2Size) els.heading2Size.value = tpl.heading2Size || tpl.headingSize || 14;
        if (els.lineSpacing) els.lineSpacing.value = String(tpl.lineHeight || 1.5);
        if (els.lineRule) els.lineRule.value = tpl.lineRule || 'auto';
        if (els.alignment) els.alignment.value = tpl.alignment || 'left';
        if (els.useIndent) els.useIndent.checked = (tpl.firstLineIndent > 0);
        if (els.firstLineIndent) els.firstLineIndent.value = String(tpl.firstLineIndent || 0);
    }
}

// ===== 模板预览效果 =====
function updateTemplatePreview() {
    const previewEl = document.querySelector('#preview-content');
    if (!previewEl || !previewEl.querySelector('h1')) return;
    
    const tpl = getEffectiveTemplate();
    const cfg = tpl.config || tpl;
    
    const headingFont = cfg.headingFont || 'SimSun';
    const bodyFont = cfg.bodyFont || 'SimSun';
    const headingSize = (cfg.headingSize || 16) + 'pt';
    const bodySize = (cfg.bodySize || 12) + 'pt';
    
    previewEl.style.fontFamily = `var(--font-serif), ${bodyFont}`;
    previewEl.querySelectorAll('h1, h2, h3').forEach(h => {
        h.style.fontFamily = `var(--font-sans), ${headingFont}`;
        h.style.fontSize = h.tagName === 'H1' ? headingSize : 
                           h.tagName === 'H2' ? (cfg.heading2Size || 14) + 'pt' : headingSize;
    });
    previewEl.querySelectorAll('p').forEach(p => {
        p.style.fontSize = bodySize;
        p.style.lineHeight = cfg.lineHeight || 1.5;
        if (cfg.firstLineIndent > 0) {
            p.style.textIndent = `${cfg.firstLineIndent}em`;
        }
    });
}

// 辅助函数：获取 state 对象（避免循环导入）
function getState() {
    return {
        currentTemplate: state_currentTemplate,
        currentMode: state_currentMode,
        currentTheme: state_currentTheme,
        customSettings: state_customSettings,
        customTemplates: state_customTemplates,
        splitRatio: state_splitRatio
    };
}

// 从 state.js 导入实际的状态变量（通过 window 暂时解决）
let state_currentTemplate, state_currentMode, state_currentTheme, state_customSettings, state_customTemplates, state_splitRatio;

function syncStateFromGlobal() {
    state_currentTemplate = window.__state?.currentTemplate;
    state_currentMode = window.__state?.currentMode;
    state_currentTheme = window.__state?.currentTheme;
    state_customSettings = window.__state?.customSettings;
    state_customTemplates = window.__state?.customTemplates;
    state_splitRatio = window.__state?.splitRatio;
}

export {
    getEffectiveTemplate,
    selectTemplate,
    updateTemplateSelection,
    renderTemplateGrid,
    updateModeSelection,
    openSettingsModal,
    closeSettingsModal,
    populateSettingsForm,
    resetSettingsForm,
    saveCurrentSettings,
    saveAsCustomTemplate,
    renderCustomTemplateList,
    loadCustomTemplate,
    deleteCustomTemplate,
    showTemplateDetail,
    restoreModalFromDetail,
    editCustomTemplate,
    updateTemplatePreview,
    syncStateFromGlobal
};
