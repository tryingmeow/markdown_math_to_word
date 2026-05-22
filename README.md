# Markdown转Word工具

一个强大的在线Markdown转Word工具，支持数学公式、代码高亮、表格等丰富格式。

## ✨ 功能特点

- **实时预览** - 左侧编辑，右侧即时预览
- **数学公式** - 支持LaTeX数学公式渲染（自动转换为图片，兼容Word）
- **代码高亮** - 支持多种编程语言语法高亮
- **丰富格式** - 支持表格、列表、链接等markdown格式
- **一键复制** - 直接复制到Word等文档编辑器
- **导出DOCX** - 生成标准 `.docx` Word 文档

## 🚀 使用方法

1. 在左侧编辑器中粘贴或编辑你的markdown内容
2. 右侧会实时预览渲染效果（数学公式自动转换为图片）
3. 满意后点击"复制内容"按钮复制到剪贴板，或点击"导出 DOCX"生成 `.docx` 文件
4. 然后可以粘贴到Word或其他文档编辑器中，或直接打开导出的 Word 文档

## 🧮 数学公式示例

### 行内公式
在长方体 $ABCD-A_1B_1C_1D_1$ 中，设 $\\overrightarrow{AB} = \\vec{a}$, $\\overrightarrow{AD} = \\vec{b}$, $\\overrightarrow{AA_1} = \\vec{c}$。

### 块级公式
$$
\\int_{-\\infty}^{\\infty} e^{-x^2} dx = \\sqrt{\\pi}
$$

## 💻 技术栈

- React + TypeScript
- Vite
- KaTeX (数学公式渲染)
- html-docx-js-typescript (DOCX导出)
- react-markdown
- highlight.js

## 🛠️ 本地开发

```bash
# 安装依赖
npm install

# 启动开发服务器
npm run dev

# 构建生产版本
npm run build

# 代码检查
npm run lint
```

## 📝 许可证

MIT License
