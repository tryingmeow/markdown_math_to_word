/**
 * remark-math 仅支持 $ / $$ 分隔符；将 LaTeX 常用的 \[ \]、\( \) 转为 $$ / $，
 * 以便公式被解析并由 MathToImage 转为图片。不处理围栏代码块内的内容。
 */
export function normalizeLatexDelimiters(markdown: string): string {
  const parts = markdown.split(/(```[\s\S]*?```)/g)
  return parts
    .map((part, i) => {
      if (i % 2 === 1) return part
      return part
        .replace(/\\\[/g, '$$')
        .replace(/\\\]/g, '$$')
        .replace(/\\\(/g, '$')
        .replace(/\\\)/g, '$')
    })
    .join('')
}
