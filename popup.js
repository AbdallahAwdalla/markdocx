/**
 * MD to Word – popup.js
 * Two input modes: Files (drag & drop / browse) and Clipboard (paste / type).
 * Shared conversion pipeline: Markdown → docx via docx.js + JSZip OMML patching.
 */
(function () {
  "use strict";

  // ═══════════════════════════════════════════════════════════════════════════
  // TAB SWITCHING
  // ═══════════════════════════════════════════════════════════════════════════
  document.querySelectorAll(".tab-btn").forEach(function(btn) {
    btn.addEventListener("click", function() {
      document.querySelectorAll(".tab-btn").forEach(function(b) { b.classList.remove("active"); });
      document.querySelectorAll(".panel").forEach(function(p) { p.classList.remove("active"); });
      btn.classList.add("active");
      document.getElementById("panel-" + btn.dataset.tab).classList.add("active");
    });
  });

  // ═══════════════════════════════════════════════════════════════════════════
  // SHARED UTILITIES
  // ═══════════════════════════════════════════════════════════════════════════
  var toast = document.getElementById("toast");

  function showToast(msg, type) {
    toast.textContent = msg;
    toast.className = "toast show " + (type || "info");
    setTimeout(function() { toast.className = "toast"; }, 2800);
  }

  function uid() { return Math.random().toString(36).slice(2); }

  function fmtSz(b) {
    return b < 1024 ? b + " B"
      : b < 1048576 ? (b / 1024).toFixed(1) + " KB"
      : (b / 1048576).toFixed(1) + " MB";
  }

  function setProgress(fillEl, barEl, pct) {
    barEl.style.display = "block";
    fillEl.style.width = pct + "%";
    if (pct >= 100) setTimeout(function() { barEl.style.display = "none"; }, 600);
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // FILES PANEL
  // ═══════════════════════════════════════════════════════════════════════════
  var dropZone    = document.getElementById("drop-zone");
  var fileInput   = document.getElementById("file-input");
  var browseLink  = document.getElementById("browse-link");
  var fileList    = document.getElementById("file-list");
  var convertBtn  = document.getElementById("convert-btn");
  var clearBtn    = document.getElementById("clear-btn");
  var progressBar   = document.getElementById("progress-bar-files");
  var progressFill  = document.getElementById("progress-fill-files");
  var optCode     = document.getElementById("opt-code");
  var optFilename = document.getElementById("opt-filename");
  var pageSize    = document.getElementById("page-size");

  var files = [];

  function addFiles(newFiles) {
    for (var i = 0; i < newFiles.length; i++) {
      var f = newFiles[i];
      if (!f.name.match(/\.(md|markdown)$/i)) {
        showToast(f.name + " is not a Markdown file", "error"); continue;
      }
      if (files.find(function(x) { return x.file.name === f.name && x.file.size === f.size; })) continue;
      files.push({ file: f, id: uid(), status: "pending" });
    }
    renderFileList();
    convertBtn.disabled = files.length === 0;
  }

  function removeFile(id) {
    files = files.filter(function(f) { return f.id !== id; });
    renderFileList();
    convertBtn.disabled = files.length === 0;
  }

  var STATUS_LABELS = { pending: "Pending", converting: "Converting\u2026", done: "\u2713 Done", error: "\u2717 Error" };

  function renderFileList() {
    fileList.innerHTML = "";
    files.forEach(function(entry) {
      var item = document.createElement("div");
      item.className = "file-item";
      item.id = "file-" + entry.id;
      item.innerHTML =
        '<span class="file-icon">\uD83D\uDCCB</span>' +
        '<div class="file-info">' +
          '<div class="file-name" title="' + entry.file.name + '">' + entry.file.name + '</div>' +
          '<div class="file-size">' + fmtSz(entry.file.size) + '</div>' +
        '</div>' +
        '<span class="file-status ' + entry.status + '">' + STATUS_LABELS[entry.status] + '</span>' +
        '<button class="remove-btn" data-id="' + entry.id + '">\u2715</button>';
      fileList.appendChild(item);
    });
    fileList.querySelectorAll(".remove-btn").forEach(function(btn) {
      btn.addEventListener("click", function() { removeFile(btn.dataset.id); });
    });
  }

  function updateFileStatus(id, status) {
    var entry = files.find(function(f) { return f.id === id; });
    if (entry) entry.status = status;
    var el = document.querySelector("#file-" + id + " .file-status");
    if (!el) return;
    el.textContent = STATUS_LABELS[status] || status;
    el.className = "file-status " + status;
  }

  // Drag & drop
  dropZone.addEventListener("dragover", function(e) { e.preventDefault(); dropZone.classList.add("drag-over"); });
  dropZone.addEventListener("dragleave", function() { dropZone.classList.remove("drag-over"); });
  dropZone.addEventListener("drop", function(e) {
    e.preventDefault(); dropZone.classList.remove("drag-over");
    addFiles(Array.from(e.dataTransfer.files));
  });
  dropZone.addEventListener("click", function() { fileInput.click(); });
  browseLink.addEventListener("click", function(e) { e.stopPropagation(); fileInput.click(); });
  fileInput.addEventListener("change", function() { addFiles(Array.from(fileInput.files)); fileInput.value = ""; });
  clearBtn.addEventListener("click", function() { files = []; renderFileList(); convertBtn.disabled = true; });

  convertBtn.addEventListener("click", async function() {
    if (!files.length) return;
    convertBtn.disabled = clearBtn.disabled = true;
    setProgress(progressFill, progressBar, 0);
    var done = 0, errors = 0;
    for (var i = 0; i < files.length; i++) {
      var entry = files[i];
      updateFileStatus(entry.id, "converting");
      try {
        var rawMd = await entry.file.text();
        var blob  = await convertMarkdown(rawMd, {
          code:     optCode.checked,
          pageSize: pageSize.value,
        });
        var outName = optFilename.checked
          ? entry.file.name.replace(/\.(md|markdown)$/i, ".docx")
          : "document.docx";
        dlBlob(blob, outName);
        updateFileStatus(entry.id, "done");
      } catch(err) {
        console.error(err);
        updateFileStatus(entry.id, "error");
        errors++;
      }
      done++;
      setProgress(progressFill, progressBar, Math.round(done / files.length * 100));
    }
    convertBtn.disabled = clearBtn.disabled = false;
    showToast(
      errors === 0
        ? "\u2713 " + done + " file" + (done > 1 ? "s" : "") + " converted!"
        : (done - errors) + " ok, " + errors + " failed",
      errors === 0 ? "success" : "error"
    );
  });

  // ═══════════════════════════════════════════════════════════════════════════
  // CLIPBOARD PANEL
  // ═══════════════════════════════════════════════════════════════════════════
  var pasteBtn        = document.getElementById("paste-btn");
  var mdTextarea      = document.getElementById("md-textarea");
  var charCount       = document.getElementById("char-count");
  var clearTextBtn    = document.getElementById("clear-text-btn");
  var convertClipBtn  = document.getElementById("convert-clip-btn");
  var clearClipBtn    = document.getElementById("clear-clip-btn");
  var clipFilename    = document.getElementById("clip-filename");
  var progressBarClip = document.getElementById("progress-bar-clip");
  var progressFillClip= document.getElementById("progress-fill-clip");
  var optCodeClip     = document.getElementById("opt-code-clip");
  var pageSizeClip    = document.getElementById("page-size-clip");

  function updateTextareaState() {
    var len = mdTextarea.value.length;
    charCount.textContent = len.toLocaleString() + " char" + (len !== 1 ? "s" : "");
    charCount.className = "char-count" + (len > 0 ? " has-text" : "");
    mdTextarea.className = "md-textarea" + (len > 0 ? " has-content" : "");
    convertClipBtn.disabled = len === 0;
  }

  mdTextarea.addEventListener("input", updateTextareaState);
  updateTextareaState();

  // Paste from clipboard button
  pasteBtn.addEventListener("click", async function() {
    pasteBtn.classList.add("pasting");
    pasteBtn.textContent = "Reading\u2026";
    try {
      var text = await navigator.clipboard.readText();
      if (!text.trim()) {
        showToast("Clipboard is empty or has no text", "error");
      } else {
        mdTextarea.value = text;
        updateTextareaState();
        showToast("\u2713 Pasted " + text.length.toLocaleString() + " chars", "success");
        // Auto-suggest filename from first heading
        var headingMatch = text.match(/^#{1,3}\s+(.+)/m);
        if (headingMatch) {
          var suggested = headingMatch[1].trim()
            .toLowerCase()
            .replace(/[^a-z0-9]+/g, "-")
            .replace(/^-|-$/g, "")
            .slice(0, 40) + ".docx";
          clipFilename.value = suggested;
        }
      }
    } catch(err) {
      // Clipboard API requires user permission; may fail in some contexts
      showToast("Clipboard access denied — paste manually below", "error");
      mdTextarea.focus();
    }
    pasteBtn.classList.remove("pasting");
    pasteBtn.innerHTML = "\uD83D\uDCCB Paste from clipboard";
  });

  // Also support Ctrl+V / Cmd+V directly on the textarea (native paste)
  // and update state when content arrives
  mdTextarea.addEventListener("paste", function() {
    setTimeout(function() {
      updateTextareaState();
      var text = mdTextarea.value;
      var headingMatch = text.match(/^#{1,3}\s+(.+)/m);
      if (headingMatch && clipFilename.value === "clipboard.docx") {
        clipFilename.value = headingMatch[1].trim()
          .toLowerCase()
          .replace(/[^a-z0-9]+/g, "-")
          .replace(/^-|-$/g, "")
          .slice(0, 40) + ".docx";
      }
    }, 0);
  });

  clearTextBtn.addEventListener("click", function() {
    mdTextarea.value = "";
    clipFilename.value = "clipboard.docx";
    updateTextareaState();
  });

  clearClipBtn.addEventListener("click", function() {
    mdTextarea.value = "";
    clipFilename.value = "clipboard.docx";
    updateTextareaState();
  });

  convertClipBtn.addEventListener("click", async function() {
    var rawMd = mdTextarea.value.trim();
    if (!rawMd) return;
    convertClipBtn.disabled = clearClipBtn.disabled = true;
    setProgress(progressFillClip, progressBarClip, 30);
    try {
      var blob = await convertMarkdown(rawMd, {
        code:     optCodeClip.checked,
        pageSize: pageSizeClip.value,
      });
      var name = clipFilename.value.trim() || "clipboard.docx";
      if (!name.endsWith(".docx")) name = name.replace(/\.[^.]+$/, "") + ".docx";
      dlBlob(blob, name);
      setProgress(progressFillClip, progressBarClip, 100);
      showToast("\u2713 Converted!", "success");
    } catch(err) {
      console.error(err);
      setProgress(progressFillClip, progressBarClip, 100);
      showToast("Conversion failed: " + err.message, "error");
    }
    convertClipBtn.disabled = clearClipBtn.disabled = false;
  });

  // ═══════════════════════════════════════════════════════════════════════════
  // MATH PIPELINE
  // ═══════════════════════════════════════════════════════════════════════════

  function extractMath(md) {
    var mathMap = new Map();
    var n = 0;
    function ph(latex, display) {
      var id = "MATHPH" + (n++);
      mathMap.set(id, { latex: latex.trim(), display: display });
      return display ? "\n\nMATHBLOCK_" + id + "\n\n" : "MATHINLINE_" + id;
    }
    var out = md
      .replace(/\$\$([\s\S]+?)\$\$/g,  function(_, l) { return ph(l, true);  })
      .replace(/\\\[([\s\S]+?)\\\]/g,   function(_, l) { return ph(l, true);  })
      .replace(/(?<!\$)\$(?!\$)([^$\n]+?)\$(?!\$)/g, function(_, l) { return ph(l, false); })
      .replace(/\\\((.+?)\\\)/gs,       function(_, l) { return ph(l, false); });
    return { processed: out, mathMap: mathMap };
  }

  function latexToOmml(latex, display) {
    try {
      var mathml = temml.renderToString(latex, { displayMode: display, throwOnError: false, xml: true });
      var clean  = mathml.replace(/\s+class="[^"]*"/g, "").replace(/\s+style="[^"]*"/g, "");
      return { omml: mathmlToOmml(clean), ok: true };
    } catch(e) {
      console.warn("Math conversion failed:", latex, e.message);
      return { ok: false };
    }
  }

  async function patchMathInDocx(blob, mathMap) {
    if (mathMap.size === 0) return blob;
    var zip = await JSZip.loadAsync(await blob.arrayBuffer());
    var xml = await zip.file("word/document.xml").async("string");

    for (var entry of mathMap) {
      var id   = entry[0];
      var info = entry[1];
      var res  = latexToOmml(info.latex, info.display);
      if (!res.ok) continue;

      if (info.display) {
        var ph = "MATHBLOCK_" + id;
        var tIdx = xml.indexOf(ph);
        if (tIdx === -1) continue;
        var pStart = xml.lastIndexOf("<w:p>", tIdx);
        if (pStart === -1) pStart = xml.lastIndexOf("<w:p ", tIdx);
        var pEnd = xml.indexOf("</w:p>", tIdx) + 6;
        if (pStart === -1 || pEnd < 6) continue;
        xml = xml.slice(0, pStart) + "<w:p>" + res.omml + "</w:p>" + xml.slice(pEnd);
      } else {
        var iph = "MATHINLINE_" + id;
        var rIdx = xml.indexOf(iph);
        if (rIdx === -1) continue;
        var rStart = xml.lastIndexOf("<w:r>", rIdx);
        if (rStart === -1) rStart = xml.lastIndexOf("<w:r ", rIdx);
        var rEnd = xml.indexOf("</w:r>", rIdx) + 6;
        if (rStart === -1 || rEnd < 6) continue;
        var oMathStart = res.omml.indexOf("<m:oMath>");
        var oMathEnd   = res.omml.lastIndexOf("</m:oMath>") + 10;
        var inlineOmml = oMathStart >= 0 ? res.omml.slice(oMathStart, oMathEnd) : res.omml;
        xml = xml.slice(0, rStart) + inlineOmml + xml.slice(rEnd);
      }
    }

    zip.file("word/document.xml", xml);
    return zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // DOCX BUILDING
  // ═══════════════════════════════════════════════════════════════════════════

  var Document      = docx.Document;
  var Packer        = docx.Packer;
  var Paragraph     = docx.Paragraph;
  var TextRun       = docx.TextRun;
  var Table         = docx.Table;
  var TableRow      = docx.TableRow;
  var TableCell     = docx.TableCell;
  var HeadingLevel  = docx.HeadingLevel;
  var AlignmentType = docx.AlignmentType;
  var LevelFormat   = docx.LevelFormat;
  var BorderStyle   = docx.BorderStyle;
  var WidthType     = docx.WidthType;
  var ShadingType   = docx.ShadingType;
  var VerticalAlign = docx.VerticalAlign;
  var ExternalHyperlink = docx.ExternalHyperlink;

  var PAGE_SIZES = {
    letter: { width: 12240, height: 15840 },
    a4:     { width: 11906, height: 16838 },
  };

  function buildNumbering() {
    return { config: [
      { reference: "bullet-list", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720,  hanging: 360 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1080, hanging: 360 } } } },
        { level: 2, format: LevelFormat.BULLET, text: "\u25AA", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
      ]},
      { reference: "ordered-list", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720,  hanging: 360 } } } },
        { level: 1, format: LevelFormat.DECIMAL, text: "%2.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1080, hanging: 360 } } } },
      ]},
    ]};
  }

  function buildStyles() {
    return {
      default: { document: { run: { font: "Calibri", size: 24 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 40, bold: true, font: "Calibri", color: "1F3864" },
          paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0,
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 4 } } } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 32, bold: true, font: "Calibri", color: "2E4057" },
          paragraph: { spacing: { before: 280, after: 100 }, outlineLevel: 1 } },
        { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 28, bold: true, font: "Calibri", color: "374151" },
          paragraph: { spacing: { before: 240, after: 80  }, outlineLevel: 2 } },
        { id: "Heading4", name: "Heading 4", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 26, bold: true, italics: true, font: "Calibri", color: "4B5563" },
          paragraph: { spacing: { before: 200, after: 60  }, outlineLevel: 3 } },
        { id: "Heading5", name: "Heading 5", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 24, bold: true, font: "Calibri", color: "6B7280" },
          paragraph: { spacing: { before: 160, after: 40  }, outlineLevel: 4 } },
        { id: "Heading6", name: "Heading 6", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 22, bold: true, italics: true, font: "Calibri", color: "9CA3AF" },
          paragraph: { spacing: { before: 120, after: 40  }, outlineLevel: 5 } },
      ],
    };
  }

  function inlineToRuns(tokens, base) {
    base = base || {};
    var runs = [];
    for (var i = 0; i < tokens.length; i++) {
      var tok = tokens[i];
      switch (tok.type) {
        case "text":
          if (tok.tokens) { runs = runs.concat(inlineToRuns(tok.tokens, base)); }
          else { runs.push(new TextRun(Object.assign({ text: tok.text || "" }, base))); }
          break;
        case "strong":
          runs = runs.concat(inlineToRuns(
            tok.tokens || [{ type: "text", text: tok.text }],
            Object.assign({}, base, { bold: true })));
          break;
        case "em":
          runs = runs.concat(inlineToRuns(
            tok.tokens || [{ type: "text", text: tok.text }],
            Object.assign({}, base, { italics: true })));
          break;
        case "codespan":
          runs.push(new TextRun({ text: tok.text, font: "Courier New", size: 20,
            color: "D63384", shading: { fill: "F3F4F6", type: ShadingType.CLEAR } }));
          break;
        case "link":
          runs.push(new ExternalHyperlink({ link: tok.href, children: [
            new TextRun(Object.assign({ text: tok.text || tok.href,
              color: "2563EB", underline: { type: "single" } }, base))
          ]}));
          break;
        case "br":
          runs.push(new TextRun({ break: 1 }));
          break;
        case "del":
          runs = runs.concat(inlineToRuns(
            tok.tokens || [{ type: "text", text: tok.text }],
            Object.assign({}, base, { strike: true })));
          break;
        default:
          if (tok.text) runs.push(new TextRun(Object.assign({ text: tok.text }, base)));
      }
    }
    return runs;
  }

  function listItems(items, ordered, level) {
    level = level || 0;
    var out = [];
    items.forEach(function(item) {
      var inlineToks = (item.tokens || []).filter(function(t) { return t.type !== "list"; });
      var runs = [];
      inlineToks.forEach(function(t) {
        runs = runs.concat(t.type === "text"
          ? inlineToRuns(t.tokens || [{ type: "text", text: t.text }])
          : inlineToRuns([t]));
      });
      out.push(new Paragraph({
        numbering: { reference: ordered ? "ordered-list" : "bullet-list", level: level },
        children: runs, spacing: { after: 60 },
      }));
      (item.tokens || []).filter(function(t) { return t.type === "list"; }).forEach(function(sub) {
        out = out.concat(listItems(sub.items, sub.ordered, level + 1));
      });
    });
    return out;
  }

  function buildTable(token) {
    var cols = token.header.length;
    var tblW = 9360, colW = Math.floor(tblW / cols);
    var colWidths = Array(cols).fill(colW);
    function bd(s) { return { style: BorderStyle.SINGLE, size: s, color: "D1D5DB" }; }
    var borders = { top: bd(4), bottom: bd(4), left: bd(4), right: bd(4) };

    var hRow = new TableRow({ tableHeader: true, children: token.header.map(function(cell) {
      return new TableCell({ borders: borders, width: { size: colW, type: WidthType.DXA },
        shading: { fill: "EFF6FF", type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          children: inlineToRuns(cell.tokens || [{ type: "text", text: cell.text }], { bold: true }),
          spacing: { after: 0 } })] });
    })});

    var bRows = token.rows.map(function(row, ri) {
      return new TableRow({ children: row.map(function(cell) {
        return new TableCell({ borders: borders, width: { size: colW, type: WidthType.DXA },
          shading: { fill: ri % 2 ? "F9FAFB" : "FFFFFF", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({
            children: inlineToRuns(cell.tokens || [{ type: "text", text: cell.text }]),
            spacing: { after: 0 } })] });
      })});
    });

    return new Table({ width: { size: tblW, type: WidthType.DXA }, columnWidths: colWidths, rows: [hRow].concat(bRows) });
  }

  var HLVL = {
    1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2,
    3: HeadingLevel.HEADING_3, 4: HeadingLevel.HEADING_4,
    5: HeadingLevel.HEADING_5, 6: HeadingLevel.HEADING_6,
  };

  function tokensToElements(tokens, opts) {
    var out = [];
    tokens.forEach(function(tok) {
      switch (tok.type) {
        case "heading":
          out.push(new Paragraph({
            heading: HLVL[tok.depth] || HeadingLevel.HEADING_1,
            children: inlineToRuns(tok.tokens || [{ type: "text", text: tok.text }]),
          }));
          break;

        case "paragraph": {
          var raw = tok.text || "";
          var mbm = raw.match(/^MATHBLOCK_(MATHPH\d+)$/);
          if (mbm) {
            out.push(new Paragraph({
              children: [new TextRun({ text: "MATHBLOCK_" + mbm[1] })],
              alignment: AlignmentType.CENTER,
              spacing: { before: 200, after: 200 },
            }));
          } else {
            out.push(new Paragraph({
              children: inlineToRuns(tok.tokens || [{ type: "text", text: raw }]),
              spacing: { after: 160 },
            }));
          }
          break;
        }

        case "blockquote":
          tokensToElements(tok.tokens, opts).forEach(function(inner) {
            if (inner instanceof Paragraph) {
              out.push(new Paragraph({
                children: (inner.options && inner.options.children) || [],
                indent: { left: 720 },
                border: { left: { style: BorderStyle.SINGLE, size: 8, color: "6B7280", space: 12 } },
                spacing: { after: 120 },
              }));
            } else { out.push(inner); }
          });
          break;

        case "code": {
          var lines = tok.text.split("\n");
          if (opts.code) {
            lines.forEach(function(line, li) {
              var isF = li === 0, isL = li === lines.length - 1;
              var border = {
                left:  { style: BorderStyle.SINGLE, size: 2, color: "E5E7EB" },
                right: { style: BorderStyle.SINGLE, size: 2, color: "E5E7EB" },
              };
              if (isF) border.top    = { style: BorderStyle.SINGLE, size: 2, color: "E5E7EB" };
              if (isL) border.bottom = { style: BorderStyle.SINGLE, size: 2, color: "E5E7EB" };
              out.push(new Paragraph({
                children: [new TextRun({ text: line || " ", font: "Courier New", size: 20 })],
                shading: { fill: "F3F4F6", type: ShadingType.CLEAR },
                spacing: { before: isF ? 160 : 0, after: isL ? 160 : 0 },
                indent: { left: 360 }, border: border,
              }));
            });
          } else {
            out.push(new Paragraph({
              children: [new TextRun({ text: tok.text, font: "Courier New", size: 20 })],
              spacing: { after: 160 },
            }));
          }
          break;
        }

        case "list":
          out = out.concat(listItems(tok.items, tok.ordered, 0));
          out.push(new Paragraph({ children: [], spacing: { after: 80 } }));
          break;

        case "table":
          out.push(buildTable(tok));
          out.push(new Paragraph({ children: [], spacing: { after: 160 } }));
          break;

        case "hr":
          out.push(new Paragraph({
            children: [],
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "D1D5DB", space: 1 } },
            spacing: { before: 200, after: 200 },
          }));
          break;

        case "space":
          out.push(new Paragraph({ children: [], spacing: { after: 80 } }));
          break;

        case "html":
          out.push(new Paragraph({
            children: [new TextRun({ text: tok.text.replace(/<[^>]+>/g, "").trim(), color: "6B7280" })],
            spacing: { after: 120 },
          }));
          break;
      }
    });
    return out;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SHARED CONVERSION ENTRY POINT
  // ═══════════════════════════════════════════════════════════════════════════

  async function convertMarkdown(rawMd, opts) {
    // 1. Extract math
    var ex = extractMath(rawMd);

    // 2. Lex
    var tokens = marked.lexer(ex.processed);

    // 3. Build elements
    var elems = tokensToElements(tokens, opts);
    if (!elems.length) elems.push(new Paragraph({ children: [new TextRun("(empty document)")] }));

    // 4. Build Document
    var sz = PAGE_SIZES[opts.pageSize] || PAGE_SIZES.letter;
    var doc = new Document({
      styles:    buildStyles(),
      numbering: buildNumbering(),
      sections: [{
        properties: { page: {
          size:   { width: sz.width, height: sz.height },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        }},
        children: elems,
      }],
    });

    // 5. Pack → Blob
    var blob = await Packer.toBlob(doc);

    // 6. Patch OMML
    blob = await patchMathInDocx(blob, ex.mathMap);

    return blob;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // DOWNLOAD HELPER
  // ═══════════════════════════════════════════════════════════════════════════
  function dlBlob(blob, name) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = name; a.click();
    setTimeout(function() { URL.revokeObjectURL(url); }, 1000);
  }

})();
