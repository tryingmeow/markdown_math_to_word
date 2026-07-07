import React, { useState, useRef, useMemo, useEffect } from 'react'
import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'
import remarkMath from 'remark-math'
import remarkBreaks from 'remark-breaks'
import rehypeHighlight from 'rehype-highlight'
import rehypeRaw from 'rehype-raw'
import rehypeKatex from 'rehype-katex'
import { asBlob } from 'html-docx-js-typescript'
import html2canvas from 'html2canvas'
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

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
  URL.revokeObjectURL(url)
}

function collectDocumentStyles(): string {
  return Array.from(document.styleSheets)
    .map((styleSheet) => {
      try {
        return Array.from(styleSheet.cssRules)
          .map((rule) => rule.cssText)
          .join('\n')
      } catch {
        return ''
      }
    })
    .filter(Boolean)
    .join('\n')
}

function bytesFromBase64(base64: string): Uint8Array {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i)
  }
  return bytes
}

function concatBytes(chunks: Uint8Array[]): Uint8Array {
  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0)
  const result = new Uint8Array(totalLength)
  let offset = 0

  chunks.forEach((chunk) => {
    result.set(chunk, offset)
    offset += chunk.length
  })

  return result
}

function createPdfFromJpegs(images: Array<{ data: Uint8Array; width: number; height: number }>): Blob {
  const encoder = new TextEncoder()
  const pageWidth = 595.28
  const pageHeight = 841.89
  const margin = 36
  const contentWidth = pageWidth - margin * 2
  const objects: Uint8Array[] = []

  const text = (value: string) => encoder.encode(value)
  const kids = images.map((_, index) => `${3 + index * 3} 0 R`).join(' ')

  objects.push(text(`<< /Type /Catalog /Pages 2 0 R >>`))
  objects.push(text(`<< /Type /Pages /Kids [${kids}] /Count ${images.length} >>`))

  images.forEach((image, index) => {
    const pageObjectNumber = 3 + index * 3
    const contentObjectNumber = pageObjectNumber + 1
    const imageObjectNumber = pageObjectNumber + 2
    const imageName = `Im${index + 1}`
    const contentHeight = Math.min((contentWidth * image.height) / image.width, pageHeight - margin * 2)
    const y = pageHeight - margin - contentHeight
    const drawCommand = `q\n${contentWidth.toFixed(2)} 0 0 ${contentHeight.toFixed(2)} ${margin.toFixed(2)} ${y.toFixed(2)} cm\n/${imageName} Do\nQ\n`

    objects.push(
      text(
        `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pageWidth} ${pageHeight}] /Resources << /XObject << /${imageName} ${imageObjectNumber} 0 R >> >> /Contents ${contentObjectNumber} 0 R >>`
      )
    )
    objects.push(text(`<< /Length ${encoder.encode(drawCommand).length} >>\nstream\n${drawCommand}endstream`))
    objects.push(
      concatBytes([
        text(
          `<< /Type /XObject /Subtype /Image /Width ${image.width} /Height ${image.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${image.data.length} >>\nstream\n`
        ),
        image.data,
        text('\nendstream'),
      ])
    )
  })

  const chunks: Uint8Array[] = [text('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n')]
  const offsets: number[] = [0]
  let length = chunks[0].length

  objects.forEach((object, index) => {
    offsets.push(length)
    const objectHeader = text(`${index + 1} 0 obj\n`)
    const objectFooter = text('\nendobj\n')
    chunks.push(objectHeader, object, objectFooter)
    length += objectHeader.length + object.length + objectFooter.length
  })

  const xrefOffset = length
  const xrefRows = offsets
    .map((offset, index) => (index === 0 ? '0000000000 65535 f ' : `${offset.toString().padStart(10, '0')} 00000 n `))
    .join('\n')
  chunks.push(
    text(
      `xref\n0 ${objects.length + 1}\n${xrefRows}\ntrailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`
    )
  )

  const pdfBytes = concatBytes(chunks)
  const pdfBuffer = pdfBytes.buffer.slice(pdfBytes.byteOffset, pdfBytes.byteOffset + pdfBytes.byteLength) as ArrayBuffer
  return new Blob([pdfBuffer], { type: 'application/pdf' })
}

function App() {
  const [markdown, setMarkdown] = useState(defaultMarkdown)
  const [isNight, setIsNight] = useState(false)

  useEffect(() => {
    document.documentElement.dataset.mode = isNight ? 'night' : 'day'
  }, [isNight])
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

    originalPres.forEach((originalPre: any, index: number) => {
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

      originalSpans.forEach((originalSpan: any, spanIndex: number) => {
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

      downloadBlob(docxBlob, 'markdown_export.docx')
    } catch (err) {
      console.error('导出 Word 失败:', err)
      alert('导出 Word 失败，请重试。')
    }
  }

  const handleExportHtml = () => {
    try {
      if (!previewRef.current) return

      const html = `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Markdown Export</title>
  <style>
${collectDocumentStyles()}
    body {
      margin: 0;
      background: #f1f5f9;
      color: #333;
    }
    .export-page {
      max-width: 960px;
      margin: 32px auto;
      padding: 40px;
      background: #fff;
      box-shadow: 0 2px 12px rgba(15, 23, 42, 0.08);
    }
    @media (max-width: 768px) {
      .export-page {
        margin: 0;
        padding: 24px;
        box-shadow: none;
      }
    }
  </style>
</head>
<body>
  <main class="preview export-page">
${previewRef.current.innerHTML}
  </main>
</body>
</html>`

      downloadBlob(new Blob([html], { type: 'text/html;charset=utf-8' }), 'markdown_export.html')
    } catch (err) {
      console.error('导出 HTML 失败:', err)
      alert('导出 HTML 失败，请重试。')
    }
  }

  const handleExportPdf = async () => {
    try {
      if (!previewRef.current) return

      const source = previewRef.current
      const captureOptions = {
        backgroundColor: '#ffffff',
        background: '#ffffff',
        width: source.scrollWidth,
        height: source.scrollHeight,
        scale: Math.min(window.devicePixelRatio || 1, 2),
        useCORS: true,
        logging: false,
        windowWidth: source.scrollWidth,
        windowHeight: source.scrollHeight,
      }
      const canvas = await html2canvas(source, captureOptions)
      const pageRatio = (841.89 - 72) / (595.28 - 72)
      const pageCanvasHeight = Math.floor(canvas.width * pageRatio)
      const images: Array<{ data: Uint8Array; width: number; height: number }> = []

      for (let y = 0; y < canvas.height; y += pageCanvasHeight) {
        const sliceHeight = Math.min(pageCanvasHeight, canvas.height - y)
        const pageCanvas = document.createElement('canvas')
        pageCanvas.width = canvas.width
        pageCanvas.height = sliceHeight
        const context = pageCanvas.getContext('2d')

        if (!context) {
          throw new Error('无法创建 PDF 画布')
        }

        context.fillStyle = '#ffffff'
        context.fillRect(0, 0, pageCanvas.width, pageCanvas.height)
        context.drawImage(canvas, 0, y, canvas.width, sliceHeight, 0, 0, canvas.width, sliceHeight)

        const imageData = pageCanvas.toDataURL('image/jpeg', 0.95).split(',')[1]
        images.push({
          data: bytesFromBase64(imageData),
          width: pageCanvas.width,
          height: pageCanvas.height,
        })
      }

      downloadBlob(createPdfFromJpegs(images), 'markdown_export.pdf')
    } catch (err) {
      console.error('导出 PDF 失败:', err)
      alert('导出 PDF 失败，请重试。')
    }
  }

  const handleTextareaChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setMarkdown(e.target.value)
  }

  return (
    <>
      <div className="wall"></div>
      <div className="wall-veil"></div>
      <div className="app fadein">
      <div className="wall"></div>
      <div className="wall-veil"></div>
      
      <header className="header glass-strong rise">
        <div className="header-content">
          <h1>📝 Markdown转Word工具</h1>
          <p>实时预览 | LaTeX公式 | 格式丰富</p>
        </div>
        <div 
          className="mode-toggle" 
          onClick={() => setIsNight(!isNight)}
          title="切换主题"
        >
          {isNight ? '🌙' : '☀️'}
        </div>
      </header>
      
      <div className="container">
        <div className="panel editor-panel glass rise" style={{ animationDelay: '0.1s' }}>
          <div className="panel-header">
            <div className="panel-title">
              <span>📝</span>
              <span>编辑器</span>
            </div>
            <span className="char-count num text-dim">{markdown.length} 字符</span>
          </div>
          <textarea
            value={markdown}
            onChange={handleTextareaChange}
            placeholder="在这里粘贴或编辑你的markdown内容..."
            className="editor"
          />
        </div>
        
        <div className="panel preview-panel glass rise" style={{ animationDelay: '0.2s' }}>
          <div className="panel-header">
            <div className="panel-title">
              <span>👁️</span>
              <span>实时预览</span>
            </div>
            <div className="export-actions">
              <button onClick={handleCopy} className="action-btn">
                <span>📋</span> 复制
              </button>
              <button onClick={handleExportWord} className="action-btn">
                <span>💾</span> 导出 DOCX
              </button>
              <button onClick={handleExportPdf} className="action-btn">
                <span>📄</span> 导出 PDF
              </button>
              <button onClick={handleExportHtml} className="action-btn">
                <span>🌐</span> 导出 HTML
              </button>
            </div>
          </div>
          <div 
            ref={previewRef}
            className="preview"
          >
            <ReactMarkdown
              remarkPlugins={[remarkMath, remarkGfm, remarkBreaks]}
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
                // 表格样式由 App.css 接管，移除硬编码的 inline style
                table: ({ children, ...props }) => (
                  <div className="table-wrapper">
                    <table {...props}>{children}</table>
                  </div>
                )
              }}
            >
              {markdownForPreview}
            </ReactMarkdown>
          </div>
        </div>
      </div>
      </div>
    </>
  )
}

export default App
