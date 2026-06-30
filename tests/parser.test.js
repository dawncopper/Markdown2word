/**
 * @vitest-environment jsdom
 */
import { describe, it, expect, beforeAll } from 'vitest';
import { ptToChinese, chineseToPt } from '../js/constants.js';

// Mock docx module
const mockDocx = {
    Document: class {},
    Packer: class {},
    Paragraph: class {},
    TextRun: class {},
    HeadingLevel: { HEADING_1: 1, HEADING_2: 2, HEADING_3: 3 },
    AlignmentType: { JUSTIFIED: 'justify', LEFT: 'left', CENTER: 'center' },
    BorderStyle: { SINGLE: 'single' },
    WidthType: { DXA: 'dxa' },
    Table: class {},
    TableRow: class {},
    TableCell: class {},
    TableLayoutType: { FIXED: 'fixed' },
    ImageRun: class {},
    ExternalHyperlink: class {}
};

global.window = { docx: mockDocx };

// Import after mock
const { parseMarkdownToParagraphs } = await import('../js/parser.js');

// Default template for tests
const defaultTemplate = {
    name: '测试模板',
    headingFont: 'SimHei',
    headingSize: 16,
    bodyFont: 'SimSun',
    bodySize: 12,
    heading2Font: 'KaiTi',
    heading2Size: 14,
    lineHeight: 1.5,
    lineRule: 'auto',
    alignment: 'left',
    firstLineIndent: 2,
    indentTwips: 320,
    margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
    spacingBefore: 0,
    spacingAfter: 0
};

describe('常量函数', () => {
    describe('ptToChinese', () => {
        it('应正确转换标准字号', () => {
            expect(ptToChinese(42)).toBe('初号');
            expect(ptToChinese(36)).toBe('小初');
            expect(ptToChinese(26)).toBe('一号');
            expect(ptToChinese(16)).toBe('三号');
            expect(ptToChinese(12)).toBe('小四');
        });

        it('应返回原始值当无法匹配时', () => {
            expect(ptToChinese(10)).toBe('10pt');
            expect(ptToChinese(99)).toBe('99pt');
        });
    });

    describe('chineseToPt', () => {
        it('应正确转换中文字号', () => {
            expect(chineseToPt('初号')).toBe(42);
            expect(chineseToPt('一号')).toBe(26);
            expect(chineseToPt('三号')).toBe(16);
            expect(chineseToPt('小四')).toBe(12);
        });

        it('应正确处理数字输入', () => {
            expect(chineseToPt(16)).toBe(16);
            expect(chineseToPt(12.5)).toBe(12.5);
        });

        it('应返回默认值当无法匹配时', () => {
            expect(chineseToPt('不存在')).toBe(12);
            expect(chineseToPt(null)).toBe(12);
        });
    });
});

describe('Markdown解析器', () => {
    describe('parseMarkdownToParagraphs', () => {
        it('应解析空字符串', () => {
            const result = parseMarkdownToParagraphs('', defaultTemplate);
            expect(result).toEqual([]);
        });

        it('应解析纯文本', () => {
            const result = parseMarkdownToParagraphs('Hello World', defaultTemplate);
            expect(result.length).toBeGreaterThan(0);
        });

        it('应解析h1标题', () => {
            const result = parseMarkdownToParagraphs('# 主标题', defaultTemplate);
            expect(result.length).toBe(1);
        });

        it('应解析h2标题', () => {
            const result = parseMarkdownToParagraphs('## 二级标题', defaultTemplate);
            expect(result.length).toBe(1);
        });

        it('应解析h3标题', () => {
            const result = parseMarkdownToParagraphs('### 三级标题', defaultTemplate);
            expect(result.length).toBe(1);
        });

        it('应解析多个标题', () => {
            const md = '# 标题1\n\n## 标题2\n\n### 标题3';
            const result = parseMarkdownToParagraphs(md, defaultTemplate);
            expect(result.length).toBe(3);
        });

        it('应解析无序列表', () => {
            const md = '- 项目1\n- 项目2\n- 项目3';
            const result = parseMarkdownToParagraphs(md, defaultTemplate);
            expect(result.length).toBe(3);
        });

        it('应解析有序列表', () => {
            const md = '1. 第一项\n2. 第二项\n3. 第三项';
            const result = parseMarkdownToParagraphs(md, defaultTemplate);
            expect(result.length).toBe(3);
        });

        it('应解析分隔线', () => {
            const md = '上文\n\n---\n\n下文';
            const result = parseMarkdownToParagraphs(md, defaultTemplate);
            expect(result.length).toBe(3);
        });

        it('应解析引用块', () => {
            const md = '> 这是一段引用';
            const result = parseMarkdownToParagraphs(md, defaultTemplate);
            expect(result.length).toBe(1);
        });

        it('应解析代码块', () => {
            const md = '```\ncode line 1\ncode line 2\n```';
            const result = parseMarkdownToParagraphs(md, defaultTemplate);
            expect(result.length).toBe(1);
        });

        it('应跳过连续空行', () => {
            const md = '段落1\n\n\n\n\n段落2';
            const result = parseMarkdownToParagraphs(md, defaultTemplate);
            expect(result.length).toBe(2);
        });
    });
});

describe('行内格式处理', () => {
    it('应处理加粗文本', () => {
        const md = '这是**加粗**文本';
        const result = parseMarkdownToParagraphs(md, defaultTemplate);
        expect(result.length).toBe(1);
    });

    it('应处理斜体文本', () => {
        const md = '这是*斜体*文本';
        const result = parseMarkdownToParagraphs(md, defaultTemplate);
        expect(result.length).toBe(1);
    });

    it('应处理删除线文本', () => {
        const md = '这是~~删除线~~文本';
        const result = parseMarkdownToParagraphs(md, defaultTemplate);
        expect(result.length).toBe(1);
    });

    it('应处理行内代码', () => {
        const md = '这是`代码`文本';
        const result = parseMarkdownToParagraphs(md, defaultTemplate);
        expect(result.length).toBe(1);
    });
});

describe('图片语法处理', () => {
    it('应解析单独一行的图片', () => {
        const md = '![alt text](https://example.com/image.png)';
        const result = parseMarkdownToParagraphs(md, defaultTemplate);
        expect(result.length).toBe(1);
    });

    it('应解析带标题的图片', () => {
        const md = '![alt text](https://example.com/image.png "图片标题")';
        const result = parseMarkdownToParagraphs(md, defaultTemplate);
        expect(result.length).toBe(1);
    });

    it('应解析base64图片', () => {
        const md = '![logo](data:image/png;base64,iVBORw0KG...)';
        const result = parseMarkdownToParagraphs(md, defaultTemplate);
        expect(result.length).toBe(1);
    });

    it('应处理行内图片', () => {
        const md = '这是一段文字包含![图片](url.png)行内图片';
        const result = parseMarkdownToParagraphs(md, defaultTemplate);
        expect(result.length).toBe(1);
    });
});

describe('表格处理', () => {
    it('应解析简单表格', () => {
        const md = '| 列1 | 列2 |\n| --- | --- |\n| A1 | A2 |';
        const result = parseMarkdownToParagraphs(md, defaultTemplate);
        expect(result.length).toBe(1);
    });

    it('应解析表头加粗的多列表格', () => {
        const md = '| 标题1 | 标题2 | 标题3 |\n| --- | --- | --- |\n| A1 | A2 | A3 |\n| B1 | B2 | B3 |';
        const result = parseMarkdownToParagraphs(md, defaultTemplate);
        expect(result.length).toBe(1);
    });
});
