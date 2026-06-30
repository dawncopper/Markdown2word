// ===== AI转Word助手 — 文本处理 =====

/**
 * 根据当前模式处理 Markdown 文本
 * @param {string} text - 原始文本
 * @param {string} mode - 处理模式
 * @returns {string} 处理后的文本
 */
function processMarkdown(text, mode) {
    switch (mode) {
        case 'purify': return purifyText(text);
        case 'rearrange': return rearrangeText(text);
        default: return text;
    }
}

/**
 * 净化模式：清除特殊符号，规范化格式
 * @param {string} text - 原始文本
 * @returns {string} 净化后的文本
 */
function purifyText(text) {
    return text
        .replace(/\n{4,}/g, '\n\n\n')
        .replace(/^#{1,6}\s*/gm, (match) => {
            return match;
        })
        .replace(/(?<![#])\n[#!~-]+\s/g, '\n')
        .replace(/\n{3}/g, '\n\n')
        .trim();
}

/**
 * 重排模式：适配豆包AI等工具的输出格式
 * @param {string} text - 原始文本
 * @returns {string} 重排后的文本
 */
function rearrangeText(text) {
    return text
        .replace(/\n{4,}/g, '\n\n\n')
        .replace(/•/g, '-')
        .replace(/–/g, '-')
        .replace(/^[\s·\-•]+/gm, '')
        .replace(/\n\s*\n/g, '\n\n')
        .trim();
}

/**
 * 处理粘贴事件（净化模式）
 * @param {ClipboardEvent} e - 粘贴事件对象
 * @returns {string|null} 处理后的纯文本，null表示不拦截
 */
function handlePaste(e) {
    if (window.__state?.currentMode !== 'purify') return null;
    
    const clipboardData = e.clipboardData || window.clipboardData;
    const html = clipboardData.getData('text/html');
    const text = clipboardData.getData('text/plain');

    if (html && html.length > text.length) {
        const parsed = parseHtmlToMarkdown(html);
        e.preventDefault();
        document.execCommand('insertText', false, parsed);
        return parsed;
    }
    return null;
}

/**
 * 解析 HTML 为 Markdown
 * @param {string} html - HTML 字符串
 * @returns {string} Markdown 字符串
 */
function parseHtmlToMarkdown(html) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    return htmlToMarkdown(doc.body);
}

/**
 * 递归将 DOM 节点转换为 Markdown
 * @param {Node} node - DOM 节点
 * @returns {string} Markdown 文本
 */
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

/**
 * 将 HTML 表格转换为 Markdown 表格
 * @param {HTMLTableElement} table - 表格元素
 * @returns {string} Markdown 表格
 */
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

export {
    processMarkdown,
    purifyText,
    rearrangeText,
    handlePaste,
    parseHtmlToMarkdown,
    htmlToMarkdown,
    htmlTableToMarkdown
};
