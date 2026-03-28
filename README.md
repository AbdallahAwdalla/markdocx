# markdocx

> A Chrome extension that converts Markdown to Word (`.docx`) — with real math, right from your browser.

![Chrome Extension](https://img.shields.io/badge/Chrome-Extension-4285F4?logo=googlechrome&logoColor=white)
![Manifest V3](https://img.shields.io/badge/Manifest-V3-green)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow)
![No Backend](https://img.shields.io/badge/No%20Backend-100%25%20Local-blueviolet)

---

## Features

- 📁 **File mode** — drag & drop or browse `.md` / `.markdown` files, convert in batch
- 📋 **Clipboard mode** — paste Markdown text directly (or type it), name your output file, download instantly
- ➗ **Real math** — `$$...$$` and `$...$` LaTeX equations are converted to native Word equations (OMML), not images or plain text
- 📝 **Rich formatting** — headings H1–H6, bold, italic, strikethrough, inline code, code blocks, blockquotes, ordered & unordered lists (nested), tables, horizontal rules, hyperlinks
- 📄 **Page size** — US Letter or A4
- 🔒 **100% local** — no servers, no uploads, everything runs in your browser

---

## What gets converted

| Markdown | Word output |
|---|---|
| `# H1` … `###### H6` | Styled headings with outline levels |
| `**bold**`, `*italic*`, `~~del~~` | Bold, italic, strikethrough runs |
| `` `code` `` | Inline code (Courier New, pink) |
| ` ```code block``` ` | Bordered monospace block |
| `> blockquote` | Left-border indented paragraph |
| `- item` / `1. item` | Bullet / numbered lists (up to 3 levels) |
| `| table |` | Styled table with alternating rows |
| `---` | Horizontal rule |
| `[text](url)` | Real hyperlink |
| `$$E=mc^2$$` | Display equation (OMML) |
| `$x^2$` | Inline equation (OMML) |
| `\[...\]` / `\(...\)` | Display / inline equation (OMML) |

---

## Installation

### From source (Developer Mode)

1. **Download or clone** this repository
   ```bash
   git clone https://github.com/AbdallahAwdalla/markdocx.git
   ```

2. Open Chrome and navigate to `chrome://extensions`

3. Enable **Developer mode** (toggle in the top-right corner)

4. Click **"Load unpacked"** and select the `markdocx` folder

5. The 📝 icon will appear in your toolbar — click it to start converting

> **Note:** The extension is entirely self-contained. All libraries (`docx.js`, `marked.js`, `temml`, `JSZip`) are bundled — no internet connection required after installation.

---

## Usage

### Files tab

1. Drag & drop one or more `.md` files onto the drop zone, or click **browse to upload**
2. Each file appears in the list with a status indicator
3. Choose options (page size, code styling, filename preservation)
4. Click **⚡ Convert to Word** — each file downloads as `.docx`

### Clipboard tab

1. Copy some Markdown text to your clipboard
2. Open the extension and switch to the **📋 Clipboard** tab
3. Click **Paste from clipboard** (or press Ctrl/Cmd+V in the textarea)
4. The output filename is auto-suggested from the first heading in your content
5. Adjust options and click **⚡ Convert to Word**

---

## Math support

markdocx uses a three-stage pipeline for equations:

```
LaTeX  →  MathML (temml)  →  OMML (custom converter)  →  patched into .docx (JSZip)
```

The resulting equations are **native Word equations** — you can click and edit them in Word just like any other equation. They are not images.

**Supported syntax:**

```markdown
Display math (centered, own line):
$$f(x) = \frac{x^2 + 1}{x - 3}$$

Inline math:
The solution is $x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$.

LaTeX delimiters also work:
\[ \sum_{i=1}^{n} x_i = 0 \]
```

---

## Tech stack

| Library | Version | Role |
|---|---|---|
| [docx.js](https://github.com/dolanmiu/docx) | 9.x | Build `.docx` structure |
| [marked.js](https://github.com/markedjs/marked) | latest | Parse Markdown tokens |
| [temml](https://github.com/ronkok/Temml) | latest | Render LaTeX → MathML |
| [JSZip](https://github.com/Stuk/jszip) | 3.x | Patch OMML into the docx ZIP |
| Custom `mathml-to-omml.js` | — | Convert MathML → OMML (Office Math) |

---

## Project structure

```
markdocx/
├── manifest.json        # Chrome extension manifest (MV3)
├── popup.html           # Extension UI (tabs, drop zone, textarea, options)
├── popup.js             # All UI logic + conversion pipeline
├── mathml-to-omml.js    # Custom MathML → OMML converter
├── docx.iife.js         # docx.js (bundled, IIFE build)
├── marked.umd.js        # marked.js (bundled, UMD build)
├── temml.min.js         # temml (bundled, minified)
├── jszip.min.js         # JSZip (bundled, minified)
└── icons/
    ├── icon16.png
    ├── icon48.png
    └── icon128.png
```

---

## Development

No build step required — the extension is plain HTML/CSS/JS.

To update a bundled library:

```bash
npm install docx marked temml jszip
cp node_modules/docx/dist/index.iife.js     docx.iife.js
cp node_modules/marked/lib/marked.umd.js    marked.umd.js
cp node_modules/temml/dist/temml.min.js     temml.min.js
cp node_modules/jszip/dist/jszip.min.js     jszip.min.js
```

Then reload the extension in `chrome://extensions`.

---

## Limitations

- Images in Markdown (`![alt](url)`) are not embedded (URLs preserved as text)
- Very large files (>1 MB of Markdown) may be slow due to in-browser processing
- Complex LaTeX macros unsupported by temml will fall back to placeholder text
- Chrome's clipboard API requires the extension popup to have focus when using the paste button

---

## License

MIT © 2026 — see [LICENSE](LICENSE)
