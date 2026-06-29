# AI转Word助手

一个纯前端在线工具，专门解决 AI 生成 Markdown 文案粘贴到 Word 时排版混乱的问题。

## 功能特点

- **一键导出**：粘贴 AI 文案，选择模板和模式，即可导出格式规范的 `.docx` 文件
- **兼容 Office/WPS**：生成标准 Office Open XML 格式
- **三种模板**：公文、职场、报告，支持自定义模板
- **三种模式**：净化模式（清除特殊符号）、重排模式（适配豆包AI）、通用模式
- **三种主题**：浅色、深色、护眼黄
- **隐私安全**：全程本地运行，数据不上传服务器，断网可用
- **跨平台**：支持电脑和手机浏览器使用

## 使用方法

1. 用浏览器打开 `index.html`
2. 粘贴 AI 生成的 Markdown 文案（或上传 `.txt`/`.md` 文件）
3. 选择转换模式和模板
4. 点击「导出 Word」按钮下载 `.docx` 文件

## 技术栈

- **前端**：纯 HTML + CSS + JavaScript（零后端）
- **Markdown 解析**：marked.js
- **Word 生成**：docx.js
- **存储**：localStorage（保存用户设置）

## 文件结构

```
ai-word-converter/
├── index.html      # 主页面
├── styles.css      # 样式表（含三色主题）
├── app.js          # 应用逻辑
├── test-sample.md  # 测试用示例文件
└── README.md       # 本文件
```

## 部署

直接双击 `index.html` 即可使用，无需任何服务器。也可部署到任意静态文件托管服务（GitHub Pages、Nginx 等）。

## 浏览器支持

- Chrome 90+
- Edge 90+
- Firefox 88+
- Safari 14+
- 移动端浏览器

## 许可证

MIT License
