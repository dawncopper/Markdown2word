// ===== AI转Word助手 — Word 导出模块 =====

import { DEFAULT_TEMPLATES, MODE_NAMES } from './constants.js';
import { showToast } from './ui.js';
import { processMarkdown } from './textProcessor.js';
import { parseMarkdownToParagraphs, generateTableOfContents } from './parser.js';

/**
 * 获取有效的模板配置
 * @returns {Object} 模板配置
 */
function getEffectiveTemplate() {
    const state = window.__state || {};
    const { currentTemplate, customTemplates, customSettings } = state;
    
    if (currentTemplate?.startsWith('custom_') && customTemplates?.length) {
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

/**
 * 导出 Word 文档
 * @param {string} text - Markdown 文本
 * @param {string} mode - 处理模式
 * @returns {Promise<void>}
 */
async function exportToWord(text, mode) {
    if (!text || !text.trim()) {
        showToast('请先粘贴或上传AI文案', 'error');
        return;
    }

    const loadingOverlay = document.querySelector('#loading-overlay');
    if (loadingOverlay) loadingOverlay.classList.add('show');

    try {
        const processed = processMarkdown(text, mode);
        const template = getEffectiveTemplate();
        
        // Check if TOC is enabled
        const includeToc = document.querySelector('#include-toc')?.checked || false;
        
        let children = [];
        
        // Add TOC if enabled
        if (includeToc) {
            const tocParagraphs = generateTableOfContents(processed, template);
            children.push(...tocParagraphs);
        }
        
        // Add main content
        const contentParagraphs = parseMarkdownToParagraphs(processed, template);
        children.push(...contentParagraphs);

        if (children.length === 0) {
            showToast('没有可导出的内容', 'error');
            return;
        }

        if (!window.docx) {
            throw new Error('docx库未加载，请检查网络连接或刷新页面');
        }
        const { Document, Packer } = window.docx;
        const modeLabel = MODE_NAMES[mode] || mode;
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
        if (loadingOverlay) loadingOverlay.classList.remove('show');
    }
}

/**
 * 下载 Blob 文件
 * @param {Blob} blob - 文件 Blob
 * @param {string} filename - 文件名
 */
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

/**
 * 批量导出 Word 文档
 * @param {Array<{name: string, content: string}>} files - 文件数组
 * @param {string} mode - 处理模式
 * @returns {Promise<void>}
 */
async function batchExport(files, mode) {
    if (!files || files.length === 0) {
        showToast('请先上传文件', 'error');
        return;
    }

    const loadingOverlay = document.querySelector('#loading-overlay');
    const loadingText = loadingOverlay?.querySelector('p');
    if (loadingOverlay) {
        loadingOverlay.classList.add('show');
        if (loadingText) loadingText.textContent = '正在批量导出...';
    }

    try {
        const template = getEffectiveTemplate();
        if (!window.docx) {
            throw new Error('docx库未加载，请检查网络连接或刷新页面');
        }
        const { Document, Packer } = window.docx;
        const modeLabel = MODE_NAMES[mode] || mode;
        const tplName = template.name || '自定义';
        const dateStr = new Date().toLocaleDateString('zh-CN').replace(/\//g, '');

        let successCount = 0;
        let failCount = 0;

        for (const file of files) {
            try {
                const processed = processMarkdown(file.content, mode);
                
                // Check if TOC is enabled
                const includeToc = document.querySelector('#include-toc')?.checked || false;
                
                let children = [];
                
                // Add TOC if enabled
                if (includeToc) {
                    const tocParagraphs = generateTableOfContents(processed, template);
                    children.push(...tocParagraphs);
                }
                
                // Add main content
                const contentParagraphs = parseMarkdownToParagraphs(processed, template);
                children.push(...contentParagraphs);

                if (children.length === 0) {
                    failCount++;
                    continue;
                }

                const baseName = file.name.replace(/\.(txt|md|html)$/i, '');
                const fileName = `AI转Word_${baseName}_${modeLabel}模式_${tplName}_${dateStr}.docx`;

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
                successCount++;

                // Small delay between downloads to avoid browser blocking
                await new Promise(resolve => setTimeout(resolve, 300));
            } catch (err) {
                console.error(`Failed to export ${file.name}:`, err);
                failCount++;
            }
        }

        if (loadingText) loadingText.textContent = '正在生成Word文档...';

        if (successCount > 0) {
            showToast(`✅ 成功导出 ${successCount} 个文件${failCount > 0 ? `，${failCount} 个失败` : ''}`, 'success');
        } else {
            showToast('❌ 所有文件导出失败', 'error');
        }
    } catch (error) {
        console.error('Batch export failed:', error);
        showToast(`❌ 批量导出失败: ${error.message}`, 'error');
    } finally {
        if (loadingOverlay) {
            loadingOverlay.classList.remove('show');
            const loadingText = loadingOverlay.querySelector('p');
            if (loadingText) loadingText.textContent = '正在生成Word文档...';
        }
    }
}

export {
    exportToWord,
    batchExport,
    downloadBlob
};
