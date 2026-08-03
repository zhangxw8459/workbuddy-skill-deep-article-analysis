# -*- coding: utf-8 -*-
"""
md2all.py v2 — 将 Markdown 分析报告转换为三格式：
  .md（已有） → .html（精美响应式网页）→ .pdf（打印版，方头方脑）

工作流：MD → HTML（完整CSS排版）→ PDF（Playwright浏览器渲染，与HTML完全一致）

用法: python3 md2all.py <输入.md> [输出目录]
作者: 小巷老张
"""

import sys
import os
import re
import markdown


def md_to_html(md_path: str) -> tuple[str, str]:
    """Convert Markdown to fully styled HTML. Returns (html_content, title)."""
    with open(md_path, 'r', encoding='utf-8') as f:
        md_content = f.read()

    base = os.path.splitext(os.path.basename(md_path))[0]
    title_display = base.replace('-', ' ')

    # Extract the first # heading as the title
    title_match = re.search(r'^# (.+)$', md_content, re.MULTILINE)
    if title_match:
        proper_title = title_match.group(1)
    else:
        proper_title = title_display

    # Remove the first # heading from body (it becomes the page title)
    md_body = re.sub(r'^# .+\n?', '', md_content, count=1)

    body_html = markdown.markdown(
        md_body,
        extensions=[
            'markdown.extensions.fenced_code',
            'markdown.extensions.tables',
            'markdown.extensions.nl2br',
        ],
        output_format='html5',
    )

    # Wrap tables with responsive container
    body_html = body_html.replace('<table>', '<div class="table-wrap"><table>')
    body_html = body_html.replace('</table>', '</table></div>')

    # Wrap images
    body_html = re.sub(r'(<img[^>]+>)', r'<div class="img-wrap">\1</div>', body_html)

    css = """
@page {
    size: A4;
    margin: 2cm 2cm 2cm 2cm;
    @bottom-center {
        content: "第 " counter(page) " 页";
        font-family: "Noto Sans SC", "Microsoft YaHei", sans-serif;
        font-size: 8pt;
        color: #888;
        border-top: 0.5pt solid #ddd;
        padding-top: 4mm;
    }
}
@page :first {
    @bottom-center { content: none; }
}
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
    font-family: "Noto Sans SC", "Microsoft YaHei", "PingFang SC", -apple-system, sans-serif;
    font-size: 10.5pt;
    line-height: 1.75;
    color: #2c3e50;
    text-align: justify;
    padding: 0;
    background: #fff;
}
/* ── Cover Page ── */
.cover-page {
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
    height: 100vh;
    text-align: center;
    page-break-after: always;
    padding: 2cm;
}
.cover-page h1 {
    font-size: 26pt;
    font-weight: 700;
    color: #1a1a2e;
    margin-bottom: 16pt;
    line-height: 1.35;
}
.cover-page .meta {
    font-size: 10pt;
    color: #7f8c8d;
    margin-bottom: 8pt;
}
.cover-page .deco {
    width: 60mm;
    height: 2pt;
    background: linear-gradient(90deg, transparent, #1a1a2e, transparent);
    margin: 20pt auto;
}
.cover-page .subtitle {
    font-size: 11pt;
    color: #555;
}
/* ── Content ── */
h1 { font-size: 18pt; color: #1a1a2e; margin: 24pt 0 12pt; border-bottom: 2pt solid #1a1a2e; padding-bottom: 6pt; page-break-after: avoid; }
h2 { font-size: 15pt; color: #2d2d44; margin: 20pt 0 10pt; border-bottom: 1pt solid #ddd; padding-bottom: 4pt; page-break-after: avoid; }
h3 { font-size: 13pt; color: #3d3d55; margin: 16pt 0 8pt; page-break-after: avoid; }
h4 { font-size: 11.5pt; color: #4a4a60; margin: 12pt 0 6pt; }
p { margin: 8pt 0; text-indent: 2em; }
p.no-indent { text-indent: 0; }
/* Lists */
ul, ol { margin: 8pt 0 8pt 2em; }
li { margin: 4pt 0; }
/* Code */
code {
    background: #f0f0f0;
    padding: 1pt 5pt;
    border-radius: 3pt;
    font-size: 9.5pt;
    font-family: "SF Mono", "Fira Code", "Consolas", monospace;
    word-break: break-all;
}
pre {
    background: #1e1e2e;
    color: #cdd6f4;
    padding: 12pt;
    border-radius: 6pt;
    overflow-x: auto;
    margin: 12pt 0;
    font-size: 9pt;
    page-break-inside: avoid;
}
pre code { background: transparent; padding: 0; color: inherit; }
/* Blockquote */
blockquote {
    border-left: 4pt solid #1a1a2e;
    background: #f5f5fa;
    padding: 10pt 16pt;
    margin: 12pt 0;
    border-radius: 0 6pt 6pt 0;
    color: #34495e;
    page-break-inside: avoid;
}
blockquote p { margin: 4pt 0; text-indent: 0; }
/* Tables */
.table-wrap { overflow-x: auto; margin: 12pt 0; }
table {
    width: 100%;
    border-collapse: collapse;
    font-size: 9.5pt;
    page-break-inside: avoid;
}
th, td {
    border: 1pt solid #ddd;
    padding: 8pt 10pt;
    text-align: left;
}
th {
    background: #1a1a2e;
    color: #fff;
    font-weight: 600;
}
tr:nth-child(even) { background: #f8f8fa; }
/* Images */
.img-wrap { text-align: center; margin: 12pt 0; }
.img-wrap img { max-width: 100%; height: auto; }
/* Strong */
strong { color: #1a1a2e; font-weight: 700; }
/* HR */
hr { border: none; border-top: 1pt solid #ddd; margin: 20pt 0; }
/* Star ratings */
.stars { color: #f1c40f; }
/* Tags */
.tag {
    display: inline-block;
    background: #eef;
    color: #1a1a2e;
    padding: 2pt 7pt;
    border-radius: 10pt;
    font-size: 8.5pt;
    margin: 2pt;
}
/* Two-column layout for comparison */
.two-col { display: flex; gap: 16pt; margin: 12pt 0; }
.two-col > div { flex: 1; }
/* Print: avoid page breaks inside blocks */
h1, h2, h3, h4 { page-break-after: avoid; }
@media print {
    body { font-size: 10pt; }
    .cover-page { height: 100vh; }
}
"""

    # Build cover page (extract from body the first set of metadata)
    meta_line = ""
    meta_match = re.search(r'\*分析日期[^*]+\*', md_content)
    if meta_match:
        meta_line = meta_match.group(0).strip('*')

    cover = f'''<div class="cover-page">
    <h1>{proper_title}</h1>
    <div class="deco"></div>
    <div class="meta">{meta_line}</div>
    <div class="subtitle">深度阅读分析报告</div>
</div>'''

    full_html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<style>{css}</style>
</head>
<body>
{cover}
<div class="content">
{body_html}
</div>
</body>
</html>'''

    return full_html, title_display


def convert(md_path: str, output_dir: str | None = None) -> dict[str, str]:
    if output_dir is None:
        output_dir = os.path.dirname(os.path.abspath(md_path))
    os.makedirs(output_dir, exist_ok=True)

    base = os.path.splitext(os.path.basename(md_path))[0]

    # Step 1: HTML
    full_html, title = md_to_html(md_path)
    html_path = os.path.join(output_dir, f"{base}.html")
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(full_html)
    print(f"HTML: {html_path}")

    # Step 2: PDF via Playwright (Chromium engine, full CSS support)
    pdf_path = os.path.join(output_dir, f"{base}.pdf")
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            page.set_content(full_html, wait_until='domcontentloaded')
            page.pdf(
                path=pdf_path,
                format='A4',
                print_background=True,
                margin={'top': '0mm', 'bottom': '0mm', 'left': '0mm', 'right': '0mm'},
            )
            browser.close()
        pdf_size = os.path.getsize(pdf_path)
        print(f"PDF:  {pdf_path} ({pdf_size/1024:.0f} KB)")
    except Exception as e:
        print(f"PDF_ERROR: {e}")
        import traceback
        print(traceback.format_exc())

    return {"html": html_path, "pdf": pdf_path if os.path.exists(pdf_path) else None}


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python3 md2all.py <input.md> [output_dir]")
        sys.exit(1)

    md_path = os.path.abspath(sys.argv[1])
    if not os.path.exists(md_path):
        print(f"File not found: {md_path}")
        sys.exit(1)

    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    convert(md_path, output_dir)
