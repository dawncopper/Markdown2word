/**
 * @vitest-environment jsdom
 */
import { describe, it, expect } from 'vitest';
import { processMarkdown, purifyText, rearrangeText, htmlToMarkdown } from '../js/textProcessor.js';

describe('文本处理器', () => {
    describe('purifyText (净化模式)', () => {
        it('应移除多余空行', () => {
            const input = '段落1\n\n\n\n\n段落2';
            const result = purifyText(input);
            expect(result.split('\n').filter(l => l.trim()).length).toBe(2);
        });

        it('应规范化标题标记', () => {
            const input = '# 标题1\n## 标题2';
            const result = purifyText(input);
            expect(result).toContain('# 标题1');
            expect(result).toContain('## 标题2');
        });

        it('应去除多余符号', () => {
            const input = '上文\n#!~- 符号\n下文';
            const result = purifyText(input);
            expect(result).not.toContain('#!~-');
        });

        it('应去除首尾空白', () => {
            const input = '   内容   ';
            const result = purifyText(input);
            expect(result.trim()).toBe('内容');
        });
    });

    describe('rearrangeText (重排模式)', () => {
        it('应统一列表符号', () => {
            const input = '项目1\n• 项目2\n– 项目3';
            const result = rearrangeText(input);
            expect(result).toContain('- 项目1');
            expect(result).toContain('- 项目2');
            expect(result).toContain('- 项目3');
        });

        it('应去除多余空行', () => {
            const input = '段落1\n\n\n\n\n段落2';
            const result = rearrangeText(input);
            expect(result.split('\n').filter(l => l.trim()).length).toBe(2);
        });

        it('应去除列表前缀空白', () => {
            const input = '   - 项目1\n  - 项目2';
            const result = rearrangeText(input);
            expect(result).toContain('- 项目1');
        });
    });

    describe('processMarkdown', () => {
        it('应调用净化模式', () => {
            const input = '测试\n\n\n\n测试';
            const result = processMarkdown(input, 'purify');
            expect(result.split('\n\n').filter(l => l.trim()).length).toBe(2);
        });

        it('应调用重排模式', () => {
            const input = '• 项目';
            const result = processMarkdown(input, 'rearrange');
            expect(result).toContain('- 项目');
        });

        it('应返回原始文本当模式未知', () => {
            const input = '原始文本';
            const result = processMarkdown(input, 'unknown');
            expect(result).toBe(input);
        });
    });
});

describe('HTML转Markdown', () => {
    describe('htmlToMarkdown', () => {
        it('应转换加粗标签', () => {
            const div = document.createElement('div');
            div.innerHTML = '<strong>加粗</strong>';
            const result = htmlToMarkdown(div);
            expect(result).toBe('**加粗**');
        });

        it('应转换斜体标签', () => {
            const div = document.createElement('div');
            div.innerHTML = '<em>斜体</em>';
            const result = htmlToMarkdown(div);
            expect(result).toBe('*斜体*');
        });

        it('应转换删除线标签', () => {
            const div = document.createElement('div');
            div.innerHTML = '<del>删除</del>';
            const result = htmlToMarkdown(div);
            expect(result).toBe('~~删除~~');
        });

        it('应转换代码标签', () => {
            const div = document.createElement('div');
            div.innerHTML = '<code>代码</code>';
            const result = htmlToMarkdown(div);
            expect(result).toBe('`代码`');
        });

        it('应转换h1标题', () => {
            const div = document.createElement('div');
            div.innerHTML = '<h1>标题</h1>';
            const result = htmlToMarkdown(div);
            expect(result).toContain('# 标题');
        });

        it('应转换h2标题', () => {
            const div = document.createElement('div');
            div.innerHTML = '<h2>二级标题</h2>';
            const result = htmlToMarkdown(div);
            expect(result).toContain('## 二级标题');
        });

        it('应转换无序列表', () => {
            const div = document.createElement('div');
            div.innerHTML = '<ul><li>项目1</li><li>项目2</li></ul>';
            const result = htmlToMarkdown(div);
            expect(result).toContain('- 项目1');
            expect(result).toContain('- 项目2');
        });

        it('应转换有序列表', () => {
            const div = document.createElement('div');
            div.innerHTML = '<ol><li>第一</li><li>第二</li></ol>';
            const result = htmlToMarkdown(div);
            expect(result).toContain('1. 第一');
            expect(result).toContain('2. 第二');
        });

        it('应保留文本节点', () => {
            const div = document.createElement('div');
            div.textContent = '纯文本';
            const result = htmlToMarkdown(div);
            expect(result).toBe('纯文本');
        });
    });
});
