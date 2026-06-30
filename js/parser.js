// ===== AI转Word助手 — Markdown 解析核心 =====
// docx 通过 UMD 全局引入，直接从 window.docx 获取
const docx = window.docx;
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
    BorderStyle, WidthType, Table, TableRow, TableCell, TableLayoutType,
    ImageRun, ExternalHyperlink, FootnoteReferenceRun } = docx || {};

/**
 * 解析 Markdown 文本为 docx 段落数组
 * @param {string} markdown - Markdown 文本
 * @param {Object} template - 模板配置
 * @returns {Paragraph[]} 段落数组
 */
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
            paragraphs.push(createHorizontalRule());
            i++;
            continue;
        }

        // Tables
        if (line.startsWith('|')) {
            const tableResult = parseTable(lines, i, template);
            if (tableResult.paragraphs.length > 0) {
                paragraphs.push(...tableResult.paragraphs);
            }
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

        // Standalone image line: ![alt](url)
        if (/^!\[([^\]]*)\]\(([^)]+)\)$/.test(line)) {
            paragraphs.push(createImageParagraph(line, template));
            i++;
            continue;
        }

        // Block math formula: $$...$$ or \[...\]
        if (/^\$\$$/.test(line) || /^\[/.test(line)) {
            const mathLines = [];
            const isDoubleDollar = line === '$$';
            const closingPattern = isDoubleDollar ? '$$' : '\\]';
            
            mathLines.push(line.replace(/^\$\$|^\\\[/, ''));
            i++;
            while (i < lines.length && 
                   !lines[i].trim().endsWith('$$') && 
                   !lines[i].trim().endsWith('\\]')) {
                mathLines.push(lines[i]);
                i++;
            }
            if (i < lines.length) {
                mathLines.push(lines[i].replace(/\$\$$|\\\]$/, ''));
                i++;
            }
            
            paragraphs.push(createMathBlock(mathLines.join(' '), template));
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

/**
 * 创建水平分隔线
 * @returns {Paragraph}
 */
function createHorizontalRule() {
    return new Paragraph({
        border: {
            bottom: { style: BorderStyle.SINGLE, size: 6, color: 'auto' }
        },
        spacing: { before: 120, after: 120 }
    });
}

/**
 * 创建标题段落
 * @param {string} text - 标题文本
 * @param {number} level - 标题级别 (1-3)
 * @param {Object} template - 模板配置
 * @returns {Paragraph}
 */
function createHeading(text, level, template) {
    // 获取各级标题字号
    const sizes = {
        1: template.headingSize || 22,
        2: template.heading2Size || template.headingSize || 16,
        3: template.heading3Size || template.heading2Size || template.headingSize || 16
    };
    
    // 获取各级标题字体（含降级）
    const fonts = {
        1: template.headingFont,
        2: template.heading2Font || template.headingFont,
        3: template.heading3Font || template.heading2Font || template.headingFont
    };

    // 一级标题居中，二级三级左对齐（左空2字）
    const alignment = level === 1 ? AlignmentType.CENTER : AlignmentType.LEFT;

    return new Paragraph({
        children: [new TextRun({
            text: text,
            bold: false,
            font: fonts[level],
            size: sizes[level] * 2
        })],
        heading: level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3,
        alignment: alignment,
        indent: level > 1 ? { firstLine: template.indentTwips || 320 } : undefined,
        spacing: {
            before: level === 1 ? 400 : 200,
            after: 200
        }
    });
}

/**
 * 创建普通段落
 * @param {string} text - 段落文本
 * @param {Object} template - 模板配置
 * @returns {Paragraph}
 */
function createParagraph(text, template) {
    const runs = processInlineFormat(text, template);

    const spacingOpts = {};
    if (template.lineRule === 'exact') {
        spacingOpts.line = Math.round(template.lineHeight * 20);
        spacingOpts.lineRule = 'exact';
    } else {
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

/**
 * 处理行内格式（加粗、斜体、删除线、代码、图片、脚注、数学公式）
 * @param {string} text - 原始文本
 * @param {Object} template - 模板配置
 * @returns {(TextRun|Paragraph)[]}
 */
function processInlineFormat(text, template) {
    const runs = [];
    const bodyFont = template.bodyFont || 'SimSun';
    const bodySize = template.bodySize || 12;

    // Handle bold, italic, strikethrough, code, image, footnote, inline math
    // Math patterns: $...$, \(...\)，footnote: [^...]
    const parts = text.split(/(\*\*[^*]+\*\*|__[^_]+__|\*[^*]+\*|_[^_]+_|~~[^~]+~~|`[^`]+`|!\[([^\]]*)\]\(([^)]+)\)|\[\^[^\]]+\]|\$[^$]+\$|\\\([^)]+\\\))/g);

    parts.forEach(part => {
        if (!part) return;

        // Footnote reference: [^1]
        const footnoteMatch = part.match(/^\[\^([^\]]+)\]$/);
        if (footnoteMatch) {
            runs.push(new TextRun({
                text: `[${footnoteMatch[1]}]`,
                font: bodyFont,
                size: bodySize * 2,
                superscript: true,
                color: '3b82f6'
            }));
            return;
        }

        // Inline math: $...$ or \(...\)
        const inlineMathMatch = part.match(/^\$([^$]+)\$$/) || part.match(/^\\\(([^)]+)\\\)$/);
        if (inlineMathMatch) {
            runs.push(new TextRun({
                text: ` ${inlineMathMatch[1]} `,
                font: 'Cambria Math',
                size: bodySize * 2,
                color: '666666',
                italics: true
            }));
            return;
        }

        // Image: ![alt](url)
        const imageMatch = part.match(/^!\[([^\]]*)\]\(([^)]+)\)$/);
        if (imageMatch) {
            const [, alt, url] = imageMatch;
            // For images, we add alt text as a hyperlink
            if (url.startsWith('data:')) {
                // Base64 image - try to embed
                runs.push(new TextRun({
                    text: `[图片: ${alt || 'embedded image'}]`,
                    font: bodyFont,
                    size: bodySize * 2,
                    color: '666666',
                    italics: true
                }));
            } else {
                // External URL - show as link text
                runs.push(new ExternalHyperlink({
                    children: [new TextRun({
                        text: `[图片: ${alt || '点击查看'}]`,
                        font: bodyFont,
                        size: bodySize * 2,
                        color: '3b82f6',
                        underline: { type: 'single' }
                    })],
                    link: url
                }));
            }
            return;
        }

        if (/^\*\*.*\*\*$/.test(part) || /^__.*__$/.test(part)) {
            runs.push(new TextRun({
                text: part.replace(/^\*\*|\*\*$|^__|__$/g, ''),
                bold: true,
                font: bodyFont,
                size: bodySize * 2
            }));
        } else if (/^\*.*\*$/.test(part) || /^_.*_$/.test(part)) {
            runs.push(new TextRun({
                text: part.replace(/^\*|\*|^_|_$/g, ''),
                italic: true,
                font: bodyFont,
                size: bodySize * 2
            }));
        } else if (/^~~.*~~/.test(part)) {
            runs.push(new TextRun({
                text: part.replace(/^~~|~~$/g, ''),
                strike: true,
                font: bodyFont,
                size: bodySize * 2
            }));
        } else if (/^`.*`/.test(part)) {
            runs.push(new TextRun({
                text: part.replace(/^`|`$/g, ''),
                font: 'Consolas',
                size: bodySize * 2,
                color: 'c7254e'
            }));
        } else {
            runs.push(new TextRun({
                text: part,
                font: bodyFont,
                size: bodySize * 2
            }));
        }
    });

    return runs.length > 0 ? runs : [new TextRun({ text, font: bodyFont, size: bodySize * 2 })];
}

/**
 * 创建无序列表项
 * @param {string[]} items - 列表项文本数组
 * @param {Object} template - 模板配置
 * @returns {Paragraph[]}
 */
function createBullets(items, template) {
    const bodyFont = template.bodyFont || 'SimSun';
    const bodySize = template.bodySize || 12;
    
    return items.map(item => new Paragraph({
        children: [new TextRun({
            text: item,
            font: bodyFont,
            size: bodySize * 2
        })],
        bullet: { level: 0 },
        spacing: { after: 60 }
    }));
}

/**
 * 创建有序列表项
 * @param {string[]} items - 列表项文本数组
 * @param {Object} template - 模板配置
 * @returns {Paragraph[]}
 */
function createNumberedItems(items, template) {
    const bodyFont = template.bodyFont || 'SimSun';
    const bodySize = template.bodySize || 12;
    
    return items.map((item, idx) => new Paragraph({
        children: [new TextRun({
            text: item,
            font: bodyFont,
            size: bodySize * 2
        })],
        number: { reference: 'default-numbering', level: 0 },
        spacing: { after: 60 }
    }));
}

/**
 * 创建代码块
 * @param {string} code - 代码文本
 * @param {Object} template - 模板配置
 * @returns {Paragraph}
 */
function createCodeBlock(code, template) {
    const bodyFont = template.bodyFont || 'SimSun';
    const bodySize = template.bodySize || 12;
    const lines = code.split('\n');
    
    return new Paragraph({
        children: lines.map((line, idx) =>
            new TextRun({
                text: line,
                font: 'Consolas',
                size: (bodySize - 2) * 2,
                break: idx < lines.length - 1 ? 1 : 0
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

/**
 * 创建引用块
 * @param {string} text - 引用文本
 * @param {Object} template - 模板配置
 * @returns {Paragraph}
 */
function createQuote(text, template) {
    const bodyFont = template.bodyFont || 'SimSun';
    const bodySize = template.bodySize || 12;
    
    return new Paragraph({
        children: [new TextRun({
            text: '「' + text + '」',
            italic: true,
            font: bodyFont,
            size: bodySize * 2,
            color: '666666'
        })],
        spacing: { before: 60, after: 60 },
        indent: { left: 360 }
    });
}

/**
 * 创建图片段落
 * @param {string} line - 包含图片语法的完整行
 * @param {Object} template - 模板配置
 * @returns {Paragraph}
 */
function createImageParagraph(line, template) {
    const match = line.match(/^!\[([^\]]*)\]\(([^)]+)\)(?:\s*"([^"]*)")?$/);
    if (!match) {
        return new Paragraph({ children: [] });
    }
    
    const [, alt, url, title] = match;
    const bodyFont = template.bodyFont || 'SimSun';
    const bodySize = template.bodySize || 12;
    
    // Display image as a centered paragraph with link
    return new Paragraph({
        children: [
            new ExternalHyperlink({
                children: [
                    new TextRun({
                        text: `[📷 图片: ${alt || '点击查看'}${title ? ' - ' + title : ''}]`,
                        font: bodyFont,
                        size: bodySize * 2,
                        color: '3b82f6',
                        underline: { type: 'single' }
                    })
                ],
                link: url
            })
        ],
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 120 }
    });
}

/**
 * 创建数学公式块
 * @param {string} formula - 公式内容
 * @param {Object} template - 模板配置
 * @returns {Paragraph}
 */
function createMathBlock(formula, template) {
    const bodyFont = template.bodyFont || 'SimSun';
    const bodySize = template.bodySize || 12;
    
    // Clean up formula - remove $ delimiters
    const cleanFormula = formula.replace(/^\$|\$$/g, '').trim();
    
    return new Paragraph({
        children: [
            new TextRun({
                text: cleanFormula,
                font: 'Cambria Math',
                size: bodySize * 2,
                color: '333333'
            })
        ],
        alignment: AlignmentType.CENTER,
        spacing: { before: 160, after: 160 },
        indent: { left: 360, right: 360 },
        border: {
            top: { style: BorderStyle.SINGLE, size: 1, color: 'dddddd' },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: 'dddddd' }
        },
        shading: { type: 'clear', fill: 'f9f9f9', color: 'f9f9f9' }
    });
}

/**
 * 解析 Markdown 表格
 * @param {string[]} lines - 所有行
 * @param {number} startIndex - 表格起始索引
 * @param {Object} template - 模板配置
 * @returns {{paragraphs: (Table|PParagraph)[], nextIndex: number}}
 */
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

    const headingFont = template.headingFont || 'SimSun';
    const bodyFont = template.bodyFont || 'SimSun';
    const headingSize = template.headingSize || 16;
    const bodySize = template.bodySize || 12;

    const nc = Math.max(...tableRows.map(r => r.length));
    const cw = Math.floor(9000 / nc);

    const tableCells = tableRows.map((row, rowIdx) => {
        const cells = row.map(cellText => {
            const runs = [new TextRun({
                text: cellText || '',
                font: rowIdx === 0 ? headingFont : bodyFont,
                size: rowIdx === 0 ? headingSize * 2 : bodySize * 2,
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
                children: [new Paragraph({ children: [new TextRun({ text: '', font: bodyFont, size: bodySize * 2 })] })],
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

/**
 * 从 Markdown 提取标题生成目录
 * @param {string} markdown - Markdown 文本
 * @param {Object} template - 模板配置
 * @returns {Paragraph[]} 目录段落数组
 */
function generateTableOfContents(markdown, template) {
    const headings = [];
    const lines = markdown.split('\n');

    for (const line of lines) {
        const h1Match = line.match(/^# (.+)$/);
        const h2Match = line.match(/^## (.+)$/);
        const h3Match = line.match(/^### (.+)$/);

        if (h1Match) {
            headings.push({ level: 1, text: h1Match[1], indent: 0 });
        } else if (h2Match) {
            headings.push({ level: 2, text: h2Match[1], indent: 1 });
        } else if (h3Match) {
            headings.push({ level: 3, text: h3Match[1], indent: 2 });
        }
    }

    if (headings.length === 0) {
        return [];
    }

    const bodyFont = template.bodyFont || 'SimSun';
    const bodySize = template.bodySize || 12;
    const paragraphs = [];

    // Title
    paragraphs.push(new Paragraph({
        children: [new TextRun({
            text: '目  录',
            bold: true,
            font: template.headingFont || bodyFont,
            size: (template.headingSize || 16) * 2
        })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 400 }
    }));

    // TOC entries
    for (const heading of headings) {
        const indentTwips = heading.indent * 240;
        const fontSize = heading.level === 1 ? (template.headingSize || 16) * 2 :
                        heading.level === 2 ? (template.heading2Size || 14) * 2 :
                        bodySize * 2;

        paragraphs.push(new Paragraph({
            children: [
                new TextRun({
                    text: heading.text,
                    font: bodyFont,
                    size: fontSize,
                    bold: heading.level === 1
                }),
                new TextRun({
                    text: '  ' + '．'.repeat(Math.max(1, 30 - heading.text.length - heading.indent * 2)),
                    font: bodyFont,
                    size: fontSize,
                    color: '888888'
                })
            ],
            alignment: AlignmentType.LEFT,
            indent: { left: indentTwips },
            spacing: { before: 60, after: 60 }
        }));
    }

    // Page break after TOC
    paragraphs.push(new Paragraph({
        children: [new TextRun({ text: '' })],
        pageBreakBefore: true
    }));

    return paragraphs;
}

/**
 * 解析 Markdown 脚注定义
 * 格式: [^1]: 脚注内容
 * @param {string} markdown - Markdown 文本
 * @returns {Map<string, string>} 脚注映射
 */
function parseFootnoteDefinitions(markdown) {
    const footnotes = new Map();
    const regex = /^\[\^([^\]]+)\]:\s*(.+)$/gm;
    let match;
    
    while ((match = regex.exec(markdown)) !== null) {
        footnotes.set(match[1], match[2]);
    }
    
    return footnotes;
}

/**
 * 处理行内脚注引用
 * 格式: [^1]
 * @param {string} text - 文本
 * @param {Map<string, string>} footnotes - 脚注定义
 * @param {number} currentIndex - 当前脚注序号
 * @param {Object} template - 模板配置
 * @returns {{runs: Array, footnotes: Map}}
 */
function processInlineFootnotes(text, footnotes, currentIndex, template) {
    const runs = [];
    const newFootnotes = new Map(footnotes);
    const bodyFont = template.bodyFont || 'SimSun';
    const bodySize = template.bodySize || 12;
    
    // Split by footnote reference pattern
    const parts = text.split(/(\[\^[^\]]+\])/g);
    
    parts.forEach(part => {
        if (!part) return;
        
        const footnoteMatch = part.match(/^\[\^([^\]]+)\]$/);
        if (footnoteMatch) {
            const key = footnoteMatch[1];
            if (newFootnotes.has(key)) {
                // Add footnote reference
                runs.push(new TextRun({
                    text: `[${currentIndex}]`,
                    font: bodyFont,
                    size: bodySize * 2,
                    superscript: true,
                    color: '3b82f6'
                }));
                newFootnotes.set(`ref_${currentIndex}`, newFootnotes.get(key));
                newFootnotes.delete(key);
            }
        } else {
            runs.push(new TextRun({
                text: part,
                font: bodyFont,
                size: bodySize * 2
            }));
        }
    });
    
    return { runs, footnotes: newFootnotes, nextIndex: currentIndex + 1 };
}

/**
 * 生成脚注段落
 * @param {Map<string, string>} footnotes - 脚注内容
 * @param {Object} template - 模板配置
 * @returns {Paragraph[]}
 */
function generateFootnotes(footnotes, template) {
    if (footnotes.size === 0) return [];
    
    const bodyFont = template.bodyFont || 'SimSun';
    const bodySize = template.bodySize || 12;
    const paragraphs = [];
    
    // Separator line
    paragraphs.push(new Paragraph({
        border: {
            top: { style: BorderStyle.SINGLE, size: 6, color: 'cccccc' }
        },
        spacing: { before: 400, after: 200 }
    }));
    
    // Footnote heading
    paragraphs.push(new Paragraph({
        children: [new TextRun({
            text: '脚注',
            bold: true,
            font: template.headingFont || bodyFont,
            size: bodySize * 2
        })],
        spacing: { before: 200, after: 200 }
    }));
    
    // Footnote entries
    let index = 1;
    for (const [key, content] of footnotes) {
        if (key.startsWith('ref_')) {
            paragraphs.push(new Paragraph({
                children: [
                    new TextRun({
                        text: `[${index}] `,
                        font: bodyFont,
                        size: bodySize * 2,
                        bold: true,
                        color: '3b82f6'
                    }),
                    new TextRun({
                        text: content,
                        font: bodyFont,
                        size: bodySize * 2
                    })
                ],
                spacing: { before: 60, after: 60 },
                indent: { left: 240 }
            }));
            index++;
        }
    }
    
    return paragraphs;
}

export {
    parseMarkdownToParagraphs,
    createHeading,
    createParagraph,
    createBullets,
    createNumberedItems,
    createCodeBlock,
    createQuote,
    createImageParagraph,
    createMathBlock,
    parseTable,
    generateTableOfContents,
    parseFootnoteDefinitions,
    generateFootnotes
};
