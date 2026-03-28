# Changelog

All notable changes to **markdocx** are documented here.

## [1.0.0] — 2026-03-28

### Added
- **Files tab** — drag & drop or browse multiple `.md` / `.markdown` files, batch convert to `.docx`
- **Clipboard tab** — paste Markdown directly or type it; auto-suggests filename from first heading
- **Math support** — LaTeX equations (`$$...$$`, `$...$`, `\[...\]`, `\(...\)`) converted to native Word OMML equations via temml + custom MathML→OMML converter + JSZip post-processing
- **Rich Markdown** — headings H1–H6, bold, italic, strikethrough, inline code, fenced code blocks, blockquotes, bullet lists, ordered lists (nested), tables, horizontal rules, hyperlinks
- **Options** — page size (US Letter / A4), code block styling toggle, filename preservation toggle
- **No backend** — 100% in-browser, no data leaves your machine
