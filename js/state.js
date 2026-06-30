// ===== AI转Word助手 — 状态管理 =====

import { DEFAULT_TEMPLATES } from './constants.js';

// ===== 全局状态 =====
let currentTemplate = 'official';
let currentMode = 'purify';
let currentTheme = 'light';
let customSettings = null;
let customTemplates = []; // localStorage 保存的自定义模板
let splitRatio = 0.5; // 左右面板分割比例

const STORAGE_KEY = 'ai-word-converter-v3';

/**
 * 加载保存的状态
 */
function loadState() {
    try {
        const saved = localStorage.getItem(STORAGE_KEY);
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

/**
 * 保存当前状态
 */
function saveState() {
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify({
            template: currentTemplate,
            mode: currentMode,
            theme: currentTheme,
            customSettings,
            customTemplates,
            splitRatio
        }));
    } catch(e) { console.warn('Save state failed:', e); }
}

// ===== 状态访问器 =====
function getState() {
    return {
        currentTemplate,
        currentMode,
        currentTheme,
        customSettings,
        customTemplates,
        splitRatio
    };
}

function setTemplate(template) {
    currentTemplate = template;
    saveState();
}

function setMode(mode) {
    currentMode = mode;
    saveState();
}

function setTheme(theme) {
    currentTheme = theme;
    saveState();
}

function setCustomSettings(settings) {
    customSettings = settings;
    saveState();
}

function setCustomTemplates(templates) {
    customTemplates = templates;
    saveState();
}

function setSplitRatio(ratio) {
    splitRatio = ratio;
    saveState();
}

function setCurrentTemplateAndSave(template) {
    currentTemplate = template;
    saveState();
}

// ===== 状态对象（包含getter/setter） =====
const state = {};

Object.defineProperties(state, {
    currentTemplate: {
        get() { return currentTemplate; },
        set(v) { currentTemplate = v; },
        enumerable: true
    },
    currentMode: {
        get() { return currentMode; },
        set(v) { currentMode = v; },
        enumerable: true
    },
    currentTheme: {
        get() { return currentTheme; },
        set(v) { currentTheme = v; },
        enumerable: true
    },
    customSettings: {
        get() { return customSettings; },
        set(v) { customSettings = v; },
        enumerable: true
    },
    customTemplates: {
        get() { return customTemplates; },
        set(v) { customTemplates = v; },
        enumerable: true
    },
    splitRatio: {
        get() { return splitRatio; },
        set(v) { splitRatio = v; },
        enumerable: true
    }
});

export {
    state,
    getState,
    setTemplate,
    setMode,
    setTheme,
    setCustomSettings,
    setCustomTemplates,
    setSplitRatio,
    setCurrentTemplateAndSave,
    loadState,
    saveState
};
