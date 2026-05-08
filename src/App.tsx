import React, { useState, useRef, useMemo } from 'react'
import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'
import remarkMath from 'remark-math'
import rehypeHighlight from 'rehype-highlight'
import rehypeRaw from 'rehype-raw'
import rehypeKatex from 'rehype-katex'
import { asBlob } from 'html-docx-js-typescript'
import { normalizeLatexDelimiters } from './normalizeLatexDelimiters'
import 'katex/dist/katex.min.css'
import 'highlight.js/styles/github.css'
import './App.css'

const defaultMarkdown = `# Markdown转Word工具

## 功能特点

这是一个**强大**的markdown转word工具，支持：

1. **实时预览** - 左侧编辑，右侧即时预览
2. **数学公式** - 支持LaTeX数学公式渲染
3. **代码高亮** - 支持多种编程语言语法高亮
4. **丰富格式** - 支持表格、列表、链接等

## 数学公式示例

行内公式：$E = mc^2$

块级公式：
$$
\\int_{-\\infty}^{\\infty} e^{-x^2} dx = \\sqrt{\\pi}
$$

## 向量数学示例

在长方体 $ABCD-A_1B_1C_1D_1$ 中，设 $\\overrightarrow{AB} = \\vec{a}$, $\\overrightarrow{AD} = \\vec{b}$, $\\overrightarrow{AA_1} = \\vec{c}$。

用基底 $\\{\\vec{a}, \\vec{b}, \\vec{c}\\}$ 表示 $\\overrightarrow{AC_1}$ 和 $\\overrightarrow{BD_1}$：

$$\\overrightarrow{AC_1} = \\overrightarrow{AB} + \\overrightarrow{BC} + \\overrightarrow{CC_1} = \\vec{a} + \\vec{b} + \\vec{c}$$

$$\\overrightarrow{BD_1} = \\overrightarrow{BA} + \\overrightarrow{AD} + \\overrightarrow{DD_1} = -\\vec{a} + \\vec{b} + \\vec{c}$$

**解题心法**: 路径分解法 - 利用空间几何体的棱，通过向量加减法的几何意义（首尾相接），将目标向量分解为基底向量的线性组合。

## 代码示例

\`\`\`javascript
function fibonacci(n) {
  if (n <= 1) return n;
  return fibonacci(n - 1) + fibonacci(n - 2);
}
\`\`\`

## 表格示例

| 功能 | 支持状态 | 说明 |
|------|----------|------|
| 标题 | ✅ | 支持1-6级标题 |
| 列表 | ✅ | 有序/无序列表 |
| 表格 | ✅ | 支持表格格式 |
| 图片 | ✅ | 支持图片显示 |

---

**使用说明：**
1. 在左侧编辑器中粘贴或编辑你的markdown内容
2. 右侧会实时预览渲染效果（数学公式自动转换为图片）
3. 满意后点击"复制内容"按钮复制到剪贴板
4. 然后可以粘贴到Word或其他文档编辑器中

*开始编辑吧！*`

/**
 * 仅给「行内 code」加样式。若对 pre>code 也加 white-space:nowrap，Word 会把整段代码挤成一行，
 * 版面右侧像被「截断」或只显示一截。
 */
function applyInlineCodeStylesForWord(htmlContent: string): string {
  const parts = htmlContent.split(/(<pre\b[\s\S]*?<\/pre>)/gi)
  return parts
    .map((segment) => {
      if (/^<pre\b/i.test(segment)) return segment
      return segment.replace(
        /<code([^>]*)>/g,
        '<code$1 style="' +
          'background-color: #f1f3f4; ' +
          'color: #c7254e; ' +
          'padding: 2px 6px; ' +
          'margin: 0 2px; ' +
          'font-family: Consolas, \'Courier New\', Monaco, monospace; ' +
          'font-size: 90%; ' +
          'font-weight: normal; ' +
          'border: 1px solid #d1d5db; ' +
          'border-radius: 3px; ' +
          'word-wrap: break-word; ' +
          'white-space: normal; ' +
          'vertical-align: baseline; ' +
          'mso-element: span;' +
          '">'
      )
    })
    .join('')
}

function App() {
  const [markdown, setMarkdown] = useState(defaultMarkdown)
  const previewRef = useRef<HTMLDivElement>(null)
  const markdownForPreview = useMemo(
    () => normalizeLatexDelimiters(markdown),
    [markdown]
  )

  const generateWordData = async () => {
    if (!previewRef.current) return null

    // 1. 克隆预览节点以进行无损操作
    const clonedPreview = previewRef.current.cloneNode(true) as HTMLDivElement

    // 2. 将 KaTeX 公式转换为 Word 原生支持的 MathML
    const katexDisplays = clonedPreview.querySelectorAll('.katex-display')
    katexDisplays.forEach(displayEl => {
      const mathmlNode = displayEl.querySelector('.katex-mathml math')
      if (mathmlNode) {
        mathmlNode.setAttribute('xmlns', 'http://www.w3.org/1998/Math/MathML')
        const p = document.createElement('p')
        p.style.textAlign = 'center'
        p.style.margin = '16px 0'
        p.appendChild(mathmlNode.cloneNode(true))
        displayEl.parentNode?.replaceChild(p, displayEl)
      }
    })

    const katexInlines = clonedPreview.querySelectorAll('.katex')
    katexInlines.forEach(katexEl => {
      if (!clonedPreview.contains(katexEl)) return // 已经被块级处理移除了

      const mathmlNode = katexEl.querySelector('.katex-mathml math')
      if (mathmlNode) {
        mathmlNode.setAttribute('xmlns', 'http://www.w3.org/1998/Math/MathML')
        katexEl.parentNode?.replaceChild(mathmlNode.cloneNode(true), katexEl)
      }
    })

    // 3. 遍历代码块，从原始DOM计算样式并应用到克隆DOM，同时修复换行
    const originalPres = previewRef.current.querySelectorAll('pre')
    const clonedPres = clonedPreview.querySelectorAll('pre')

    originalPres.forEach((originalPre, index) => {
      const clonedPre = clonedPres[index]
      if (!clonedPre) return

      // --- Step A: 内联 <pre> 容器的样式 ---
      const preStyle = window.getComputedStyle(originalPre)
      clonedPre.style.backgroundColor = preStyle.backgroundColor
      clonedPre.style.padding = '12px 16px' // 使用固定值避免计算差异
      clonedPre.style.margin = '16px 0'
      clonedPre.style.border = '1px solid #e9ecef'
      clonedPre.style.borderRadius = '6px'
      clonedPre.style.fontFamily = "Consolas, 'Courier New', Monaco, monospace"
      clonedPre.style.fontSize = '14px'
      clonedPre.style.lineHeight = '1.5'
      clonedPre.style.whiteSpace = 'pre-wrap' // 确保长代码在Word中能换行
      clonedPre.style.wordWrap = 'break-word'
      clonedPre.style.overflowX = 'auto'

      // --- Step B: 内联代码高亮 (<span>) 的颜色 ---
      const originalSpans = originalPre.querySelectorAll('span[class^="hljs-"]')
      const clonedSpans = clonedPre.querySelectorAll('span[class^="hljs-"]')

      originalSpans.forEach((originalSpan, spanIndex) => {
        const clonedSpan = clonedSpans[spanIndex] as HTMLElement
        if (clonedSpan) {
          const spanStyle = window.getComputedStyle(originalSpan)
          clonedSpan.style.color = spanStyle.color
          clonedSpan.removeAttribute('class') // 清理class，避免Word样式冲突
        }
      })

      // --- Step C: 修复Word中的换行问题 ---
      const codeElement = clonedPre.querySelector('code')
      if (codeElement) {
        // 将换行符 \n 替换为 <br> 标签
        codeElement.innerHTML = codeElement.innerHTML.replace(/\n/g, '<br>')
      }
    })

    // 获取处理后的HTML内容
    let htmlContent = clonedPreview.innerHTML

    // 为整个内容添加Word兼容的基础样式
    htmlContent = `<div style="
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Helvetica Neue', Arial, sans-serif;
      font-size: 16px;
      line-height: 1.6;
      color: #1f2937;
      max-width: none;
      margin: 0;
      padding: 0;
      background-color: #ffffff;
      word-wrap: break-word;
      mso-line-height-rule: exactly;
      mso-text-raise: 0;
    ">${htmlContent}</div>`

    // 处理剩余的行内代码 <code> 标签（不作用于 pre>code，避免 Word 中单行溢出「截断」）
    htmlContent = applyInlineCodeStylesForWord(htmlContent)

    // 为表格添加Word兼容的样式
    htmlContent = htmlContent.replace(
      /<table([^>]*)>/g,
      '<table$1 style="' +
      'border-collapse: collapse; ' +
      'width: 100%; ' +
      'margin: 16px 0; ' +
      'font-size: 14px; ' +
      'mso-table-lspace: 0pt; ' +
      'mso-table-rspace: 0pt; ' +
      'mso-table-anchor-vertical: paragraph; ' +
      'mso-table-anchor-horizontal: margin;' +
      '">'
    )

    // 为表头单元格添加样式
    htmlContent = htmlContent.replace(
      /<th([^>]*)>/g,
      '<th$1 style="' +
      'border: 1px solid #d0d7de; ' +
      'padding: 8px 12px; ' +
      'background-color: #f6f8fa; ' +
      'font-weight: 600; ' +
      'text-align: left; ' +
      'vertical-align: top; ' +
      'mso-element: th;' +
      '">'
    )

    // 为表格单元格添加样式
    htmlContent = htmlContent.replace(
      /<td([^>]*)>/g,
      '<td$1 style="' +
      'border: 1px solid #d0d7de; ' +
      'padding: 8px 12px; ' +
      'vertical-align: top; ' +
      'mso-element: td;' +
      '">'
    )

    // 为引用块添加样式
    htmlContent = htmlContent.replace(
      /<blockquote([^>]*)>/g,
      '<blockquote$1 style="' +
      'margin: 16px 0; ' +
      'padding: 0 16px; ' +
      'border-left: 4px solid #d1d5db; ' +
      'background-color: #f9fafb; ' +
      'font-style: italic; ' +
      'color: #6b7280; ' +
      'page-break-inside: avoid;' +
      '">'
    )

    // 为标题添加Word兼容的样式
    for (let i = 1; i <= 6; i++) {
      const fontSize = Math.max(24 - i * 2, 14) // h1=22px, h2=20px, ..., h6=14px
      htmlContent = htmlContent.replace(
        new RegExp(`<h${i}([^>]*)>`, 'g'),
        `<h${i}$1 style="` +
        `font-size: ${fontSize}px; ` +
        `font-weight: bold; ` +
        `margin: ${i === 1 ? '24px' : '20px'} 0 16px 0; ` +
        `color: #1f2937; ` +
        `line-height: 1.25; ` +
        `page-break-after: avoid; ` +
        `mso-element: h${i};` +
        `">`
      )
    }

    const plainText = previewRef.current.innerText ?? ''
    return { htmlContent, plainText }
  }

  const handleCopy = async () => {
    try {
      const data = await generateWordData()
      if (!data) return
      const { htmlContent, plainText } = data

      if (navigator.clipboard && typeof ClipboardItem !== 'undefined') {
        const htmlBlob = new Blob([htmlContent], { type: 'text/html' })
        const textBlob = new Blob([plainText], { type: 'text/plain' })
        try {
          await navigator.clipboard.write([
            new ClipboardItem({
              'text/html': htmlBlob,
              'text/plain': textBlob,
            }),
          ])
        } catch {
          await navigator.clipboard.write([new ClipboardItem({ 'text/html': htmlBlob })])
        }
      } else {
        const range = document.createRange()
        if (previewRef.current) {
          range.selectNode(previewRef.current)
          const selection = window.getSelection()

          if (selection) {
            selection.removeAllRanges()
            selection.addRange(range)

            const success = document.execCommand('copy')
            selection.removeAllRanges()

            if (!success) {
              throw new Error('传统复制方法失败')
            }
          }
        }
      }
    } catch (err) {
      console.error('复制失败:', err)

      // 自动选择内容供用户手动复制
      if (previewRef.current) {
        const range = document.createRange()
        range.selectNodeContents(previewRef.current)
        const selection = window.getSelection()
        if (selection) {
          selection.removeAllRanges()
          selection.addRange(range)
        }
      }
    }
  }

  const handleExportWord = async () => {
    try {
      const data = await generateWordData()
      if (!data) return
      const { htmlContent } = data

      // 包装成 Word 兼容的 HTML 结构，包含页面设置
      const wordHtml = `
        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
        <head>
          <meta charset='utf-8'>
          <title>Export</title>
          <style>
            @page WordSection1 {
              size: 21.0cm 29.7cm;
              margin: 2.54cm 2.0cm 2.54cm 2.0cm;
              mso-header-margin: 1.5cm;
              mso-footer-margin: 1.5cm;
              mso-paper-source: 0;
            }
            div.WordSection1 { page: WordSection1; }
            /* 确保图片和代码块不被截断 */
            img { max-width: 100%; height: auto; }
            pre { white-space: pre-wrap !important; word-wrap: break-word !important; word-break: break-all !important; }
          </style>
        </head>
        <body>
          <div class="WordSection1">
            ${htmlContent}
          </div>
        </body>
        </html>
      `

      // 使用 html-docx-js-typescript 生成真正的 .docx (OOXML) 文件
      const docxBlob = (await asBlob(wordHtml, {
        orientation: 'portrait',
        margins: { top: 1440, right: 1134, bottom: 1440, left: 1134 },
      })) as Blob

      const url = URL.createObjectURL(docxBlob)
      const link = document.createElement('a')
      link.href = url
      link.download = 'markdown_export.docx'
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
      URL.revokeObjectURL(url)
    } catch (err) {
      console.error('导出 Word 失败:', err)
      alert('导出 Word 失败，请重试。')
    }
  }

  const handleTextareaChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setMarkdown(e.target.value)
  }

  return (
    <div className="app">
      <header className="header">
        <h1>📝 Markdown转Word工具</h1>
        <p>实时预览 | 数学公式 | 格式丰富</p>
      </header>
      
      <div className="container">
        <div className="editor-panel">
          <div className="panel-header">
            <h3>📝 Markdown编辑器</h3>
            <span className="char-count">{markdown.length} 字符</span>
          </div>
          <textarea
            value={markdown}
            onChange={handleTextareaChange}
            placeholder="在这里粘贴或编辑你的markdown内容..."
            className="editor"
          />
        </div>
        
        <div className="preview-panel">
          <div className="panel-header">
            <h3>👁️ 实时预览</h3>
            <div style={{ display: 'flex', gap: '8px' }}>
              <button onClick={handleCopy} className="copy-button">
                📋 复制内容
              </button>
              <button onClick={handleExportWord} className="copy-button" style={{ background: 'linear-gradient(135deg, #10b981 0%, #059669 100%)' }}>
                💾 导出 Word
              </button>
            </div>
          </div>
          <div 
            ref={previewRef}
            className="preview"
          >
            <ReactMarkdown
              remarkPlugins={[remarkMath, remarkGfm]}
              rehypePlugins={[rehypeRaw, rehypeKatex, rehypeHighlight]}
              components={{
                // 自定义组件渲染
                img: ({ src, alt, ...props }) => (
                  <img 
                    src={src} 
                    alt={alt} 
                    {...props}
                    style={{ maxWidth: '100%', height: 'auto' }}
                  />
                ),
                // 表格样式优化
                table: ({ children, ...props }) => (
                  <table {...props} style={{ 
                    borderCollapse: 'collapse', 
                    width: '100%', 
                    marginBottom: '1rem' 
                  }}>
                    {children}
                  </table>
                ),
                th: ({ children, ...props }) => (
                  <th {...props} style={{ 
                    border: '1px solid #ddd', 
                    padding: '8px', 
                    backgroundColor: '#f5f5f5',
                    textAlign: 'left'
                  }}>
                    {children}
                  </th>
                ),
                td: ({ children, ...props }) => (
                  <td {...props} style={{ 
                    border: '1px solid #ddd', 
                    padding: '8px' 
                  }}>
                    {children}
                  </td>
                ),
              }}
            >
              {markdownForPreview}
            </ReactMarkdown>
          </div>
        </div>
      </div>
    </div>
  )
}

export default App