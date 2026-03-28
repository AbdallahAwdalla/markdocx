/**
 * mathml-to-omml.js
 * Converts MathML XML strings to OMML (Office Math Markup Language) XML strings.
 * OMML is the native math format used by Microsoft Word (.docx).
 *
 * Supports: mn, mi, mo, mrow, mfrac, msup, msub, msubsup, msqrt, mroot,
 *           mover, munder, munderover, mtext, mspace, mfenced, mtable, mtr, mtd
 */
(function (global) {
  "use strict";

  // ── Namespace prefix used in OMML ──────────────────────────────────────
  const NS = 'm:';

  // ── Operator map: MathML operators → OMML chr values ──────────────────
  const OP_MAP = {
    '+'  : '+',  '−'  : '−',  '-'  : '-',  '×'  : '×',  '÷'  : '÷',
    '='  : '=',  '≠'  : '≠',  '<'  : '&lt;', '>'  : '&gt;',
    '≤'  : '≤',  '≥'  : '≥',  '∈'  : '∈',  '∉'  : '∉',
    '⊂'  : '⊂',  '⊃'  : '⊃',  '∪'  : '∪',  '∩'  : '∩',
    '∑'  : '∑',  '∏'  : '∏',  '∫'  : '∫',  '∂'  : '∂',
    '∇'  : '∇',  '∞'  : '∞',  '±'  : '±',  '∓'  : '∓',
    '·'  : '·',  '∘'  : '∘',  '≈'  : '≈',  '∝'  : '∝',
    '→'  : '→',  '←'  : '←',  '↔'  : '↔',  '⇒'  : '⇒',
    '⇔'  : '⇔',  '…'  : '…',  '⋯'  : '⋯',  '⋮'  : '⋮',
    '⋱'  : '⋱',  '|'  : '|',  '‖'  : '‖',
    // spelled out
    '∀'  : '∀',  '∃'  : '∃',  '¬'  : '¬',  '∧'  : '∧',  '∨'  : '∨',
  };

  // ── XML helpers ────────────────────────────────────────────────────────
  function esc(s) {
    return String(s)
      .replace(/&(?!amp;|lt;|gt;|quot;|apos;)/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  function tag(name, attrs, children) {
    const a = attrs ? ' ' + Object.entries(attrs).map(([k,v]) => `${k}="${esc(v)}"`).join(' ') : '';
    if (!children && children !== 0) return `<${name}${a}/>`;
    return `<${name}${a}>${children}</${name}>`;
  }

  function run(text, style) {
    // m:r wraps a text run
    const rPr = style ? tag(NS + 'rPr', null, tag(NS + style, null, '')) : '';
    return tag(NS + 'r', null, rPr + tag(NS + 't', null, esc(text)));
  }

  // ── DOMParser shim: parse MathML string → document ────────────────────
  function parseMathML(xmlStr) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlStr, 'application/xml');
    const err = doc.querySelector('parsererror');
    if (err) throw new Error('MathML parse error: ' + err.textContent.slice(0, 120));
    return doc.documentElement;
  }

  // ── Main recursive converter ───────────────────────────────────────────
  function convertNode(node) {
    if (node.nodeType === 3 /* TEXT_NODE */) {
      const t = node.textContent.trim();
      return t ? run(t) : '';
    }
    if (node.nodeType !== 1 /* ELEMENT_NODE */) return '';

    const tag_name = node.localName.toLowerCase();

    // Strip temml class-only wrappers
    if (tag_name === 'mstyle' || tag_name === 'mpadded') {
      return convertChildren(node);
    }

    switch (tag_name) {
      case 'math':
        return convertMath(node);
      case 'mrow':
      case 'mphantom':
        return convertChildren(node);
      case 'mn':
        return run(node.textContent.trim());
      case 'mi': {
        const text = node.textContent.trim();
        // Single letter: italic; multi-letter function name: normal
        const isFunc = text.length > 1;
        if (isFunc) {
          return tag(NS + 'r', null,
            tag(NS + 'rPr', null, tag(NS + 'sty', {'m:val': 'p'}, null)) +
            tag(NS + 't', null, esc(text))
          );
        }
        return run(text);
      }
      case 'mo': {
        const op = node.textContent.trim();
        const mapped = OP_MAP[op] || op;
        return tag(NS + 'r', null, tag(NS + 't', null, esc(mapped)));
      }
      case 'mtext':
        return tag(NS + 'r', null,
          tag(NS + 'rPr', null, tag(NS + 'sty', {'m:val': 'p'}, null)) +
          tag(NS + 't', null, esc(node.textContent))
        );
      case 'mspace':
        return run('\u00a0');
      case 'mfrac':
        return convertFrac(node);
      case 'msup':
        return convertSup(node);
      case 'msub':
        return convertSub(node);
      case 'msubsup':
        return convertSubSup(node);
      case 'msqrt':
        return convertSqrt(node, null);
      case 'mroot':
        return convertRoot(node);
      case 'mover':
        return convertOver(node);
      case 'munder':
        return convertUnder(node);
      case 'munderover':
        return convertUnderOver(node);
      case 'mfenced':
        return convertFenced(node);
      case 'mtable':
        return convertTable(node);
      case 'mtr':
        return convertChildren(node);
      case 'mtd':
        return convertChildren(node);
      case 'menclose':
        return convertChildren(node); // simplified
      case 'mmultiscripts':
        return convertMultiScripts(node);
      default:
        return convertChildren(node);
    }
  }

  function convertChildren(node) {
    let out = '';
    for (const child of node.childNodes) out += convertNode(child);
    return out;
  }

  function elemChildren(node) {
    return [...node.childNodes].filter(n => n.nodeType === 1);
  }

  // ── math root element ──────────────────────────────────────────────────
  function convertMath(node) {
    const isDisplay = node.getAttribute('display') === 'block';
    const dispAttr  = isDisplay ? ' m:dispDef="1"' : '';
    // m:oMathPara wraps display math; m:oMath wraps inline
    const inner = tag('m:oMath', null, convertChildren(node));
    if (isDisplay) {
      return tag('m:oMathPara', null,
        tag('m:oMathParaPr', null, tag('m:jc', {'m:val': 'centerGroup'}, null)) +
        inner
      );
    }
    return inner;
  }

  // ── Fraction ──────────────────────────────────────────────────────────
  function convertFrac(node) {
    const [num, den] = elemChildren(node);
    return tag(NS + 'f', null,
      tag(NS + 'fPr', null, '') +
      tag(NS + 'num', null, num  ? tag(NS + 'r', null, '') + convertNode(num)  : '') +
      tag(NS + 'den', null, den ? tag(NS + 'r', null, '') + convertNode(den) : '')
    );
  }

  // ── Superscript ───────────────────────────────────────────────────────
  function convertSup(node) {
    const [base, sup] = elemChildren(node);
    return tag(NS + 'sSup', null,
      tag(NS + 'sSupPr', null, tag(NS + 'ctrlPr', null, '')) +
      tag(NS + 'e', null, base ? convertNode(base) : '') +
      tag(NS + 'sup', null, sup  ? convertNode(sup)  : '')
    );
  }

  // ── Subscript ─────────────────────────────────────────────────────────
  function convertSub(node) {
    const [base, sub] = elemChildren(node);
    return tag(NS + 'sSub', null,
      tag(NS + 'sSubPr', null, tag(NS + 'ctrlPr', null, '')) +
      tag(NS + 'e', null, base ? convertNode(base) : '') +
      tag(NS + 'sub', null, sub  ? convertNode(sub)  : '')
    );
  }

  // ── Sub+Superscript ───────────────────────────────────────────────────
  function convertSubSup(node) {
    const [base, sub, sup] = elemChildren(node);
    return tag(NS + 'sSubSup', null,
      tag(NS + 'sSubSupPr', null, tag(NS + 'ctrlPr', null, '')) +
      tag(NS + 'e',   null, base ? convertNode(base) : '') +
      tag(NS + 'sub', null, sub  ? convertNode(sub)  : '') +
      tag(NS + 'sup', null, sup  ? convertNode(sup)  : '')
    );
  }

  // ── Square root ───────────────────────────────────────────────────────
  function convertSqrt(node) {
    return tag(NS + 'rad', null,
      tag(NS + 'radPr', null,
        tag(NS + 'degHide', {'m:val': '1'}, null) +
        tag(NS + 'ctrlPr', null, '')
      ) +
      tag(NS + 'deg', null, '') +
      tag(NS + 'e', null, convertChildren(node))
    );
  }

  // ── nth Root ──────────────────────────────────────────────────────────
  function convertRoot(node) {
    const [radicand, degree] = elemChildren(node);
    return tag(NS + 'rad', null,
      tag(NS + 'radPr', null, tag(NS + 'ctrlPr', null, '')) +
      tag(NS + 'deg', null, degree   ? convertNode(degree)   : '') +
      tag(NS + 'e',   null, radicand ? convertNode(radicand) : '')
    );
  }

  // ── Overscript (accent / hat / bar) ───────────────────────────────────
  function convertOver(node) {
    const [base, over] = elemChildren(node);
    // Check if it's an accent (single combining char)
    const overText = over ? over.textContent.trim() : '';
    const isAccent = overText.length === 1;
    if (isAccent) {
      return tag(NS + 'acc', null,
        tag(NS + 'accPr', null,
          tag(NS + 'chr', {'m:val': overText}, null) +
          tag(NS + 'ctrlPr', null, '')
        ) +
        tag(NS + 'e', null, base ? convertNode(base) : '')
      );
    }
    return tag(NS + 'limUpp', null,
      tag(NS + 'limUppPr', null, tag(NS + 'ctrlPr', null, '')) +
      tag(NS + 'e', null, base ? convertNode(base) : '') +
      tag(NS + 'lim', null, over ? convertNode(over) : '')
    );
  }

  // ── Underscript ───────────────────────────────────────────────────────
  function convertUnder(node) {
    const [base, under] = elemChildren(node);
    return tag(NS + 'limLow', null,
      tag(NS + 'limLowPr', null, tag(NS + 'ctrlPr', null, '')) +
      tag(NS + 'e',   null, base  ? convertNode(base)  : '') +
      tag(NS + 'lim', null, under ? convertNode(under) : '')
    );
  }

  // ── Under+Overscript ──────────────────────────────────────────────────
  function convertUnderOver(node) {
    const [base, under, over] = elemChildren(node);
    // Render as nested limLow + limUpp
    const inner = tag(NS + 'limLow', null,
      tag(NS + 'limLowPr', null, tag(NS + 'ctrlPr', null, '')) +
      tag(NS + 'e',   null, base  ? convertNode(base)  : '') +
      tag(NS + 'lim', null, under ? convertNode(under) : '')
    );
    return tag(NS + 'limUpp', null,
      tag(NS + 'limUppPr', null, tag(NS + 'ctrlPr', null, '')) +
      tag(NS + 'e',   null, inner) +
      tag(NS + 'lim', null, over  ? convertNode(over)  : '')
    );
  }

  // ── Fenced (brackets) ─────────────────────────────────────────────────
  function convertFenced(node) {
    const open  = node.getAttribute('open')  ?? '(';
    const close = node.getAttribute('close') ?? ')';
    const sep   = node.getAttribute('separators') ?? ',';
    const children = elemChildren(node);

    // m:d with m:dPr
    return tag(NS + 'd', null,
      tag(NS + 'dPr', null,
        tag(NS + 'begChr', {'m:val': esc(open)}, null) +
        tag(NS + 'sepChr', {'m:val': esc(sep[0] || ',')}, null) +
        tag(NS + 'endChr', {'m:val': esc(close)}, null) +
        tag(NS + 'ctrlPr', null, '')
      ) +
      children.map(c => tag(NS + 'e', null, convertNode(c))).join('')
    );
  }

  // ── Matrix / table ────────────────────────────────────────────────────
  function convertTable(node) {
    const rows = elemChildren(node).filter(n => n.localName === 'mtr');
    const mRows = rows.map(row => {
      const cells = elemChildren(row).filter(n => n.localName === 'mtd');
      return tag(NS + 'mr', null,
        cells.map(cell => tag(NS + 'e', null, convertChildren(cell))).join('')
      );
    });
    return tag(NS + 'm', null,
      tag(NS + 'mPr', null,
        tag(NS + 'mcs', null,
          tag(NS + 'mc', null,
            tag(NS + 'mcPr', null,
              tag(NS + 'count', {'m:val': String(rows[0] ? elemChildren(rows[0]).length : 1)}, null) +
              tag(NS + 'mcJc', {'m:val': 'center'}, null)
            )
          )
        ) +
        tag(NS + 'ctrlPr', null, '')
      ) +
      mRows.join('')
    );
  }

  // ── Multi-scripts (pre-scripts) ───────────────────────────────────────
  function convertMultiScripts(node) {
    // Simplified: just render base
    const [base] = elemChildren(node);
    return base ? convertNode(base) : '';
  }

  // ── Public API ────────────────────────────────────────────────────────
  function mathmlToOmml(mathmlStr) {
    // Clean up temml class/style attributes that would confuse the parser
    const cleaned = mathmlStr
      .replace(/\s+class="[^"]*"/g, '')
      .replace(/\s+style="[^"]*"/g, '')
      .replace(/\s+xmlns="[^"]*"/g, '');

    const root = parseMathML(cleaned);
    return convertNode(root);
  }

  global.mathmlToOmml = mathmlToOmml;

})(typeof window !== 'undefined' ? window : global);
