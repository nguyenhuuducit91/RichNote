/*!
 * RichNote — a WYSIWYG rich-text editor for Standard Notes.
 *
 * Features: live formatting (no raw HTML), a Word-like toolbar (font, size,
 * text/highlight color, B/I/U/S, headings H1-H6, lists, alignment, indent,
 * links, undo/redo, clear formatting), cursor-aware active buttons, a line-number
 * gutter with current-line highlight, keyboard shortcuts and a Word Wrap toggle.
 * Content is stored as HTML and synced through @standardnotes/component-relay.
 *
 * @author  Nguyen Huu Duc <nguyenhuuduc.it.91@gmail.com> (VIETIS)
 * @license MIT
 */
(function () {
  'use strict';

  var menubar     = document.getElementById('menubar');
  var wrapItem    = document.getElementById('wrapItem');
  var editor      = document.getElementById('editor');
  var gutter      = document.getElementById('gutter');
  var gutterInner = document.getElementById('gutterInner');
  var toolbar     = document.getElementById('toolbar');
  var stName      = document.getElementById('st-name');
  var stPos       = document.getElementById('st-pos');
  var stSel       = document.getElementById('st-sel');
  var stLen       = document.getElementById('st-len');
  var styleSelect = document.getElementById('styleSelect');
  var fontSelect  = document.getElementById('fontSelect');
  var sizeSelect  = document.getElementById('sizeSelect');
  var sizeArrow   = document.getElementById('sizeArrow');
  var sizePop     = document.getElementById('sizePop');
  var foreBar     = document.getElementById('foreBar');
  var backBar     = document.getElementById('backBar');
  var alignBtn    = document.getElementById('alignBtn');
  var alignPop    = document.getElementById('alignPop');
  var alignCur    = document.getElementById('alignCur');
  var fmtBtn      = document.getElementById('fmtBtn');
  var fmtPop      = document.getElementById('fmtPop');
  var listBtn     = document.getElementById('listBtn');
  var listPop     = document.getElementById('listPop');
  var listCur     = document.getElementById('listCur');
  var indentBtn   = document.getElementById('indentBtn');
  var indentPop   = document.getElementById('indentPop');
  var tableBtn    = document.getElementById('tableBtn');
  var tablePop    = document.getElementById('tablePop');
  var tableGrid   = document.getElementById('tableGrid');
  var tableLabel  = document.getElementById('tableLabel');
  var linkPop     = document.getElementById('linkPop');
  var linkInput   = document.getElementById('linkInput');
  var linkApply   = document.getElementById('linkApply');
  var linkRemove  = document.getElementById('linkRemove');
  var aboutModal  = document.getElementById('aboutModal');
  var aboutClose  = document.getElementById('aboutClose');
  var scModal     = document.getElementById('scModal');
  var scClose     = document.getElementById('scClose');
  var fpBtn       = document.getElementById('fpBtn');
  var codeBar     = document.getElementById('codeBar');
  var codeLang    = document.getElementById('codeLang');
  var codeCopy    = document.getElementById('codeCopy');
  var sourceView  = document.getElementById('sourceView');

  /* Default color palette (Google Sheets style) */
  var PALETTE = [
    ['#000000','#434343','#666666','#999999','#b7b7b7','#cccccc','#d9d9d9','#efefef','#f3f3f3','#ffffff'],
    ['#980000','#ff0000','#ff9900','#ffff00','#00ff00','#00ffff','#4a86e8','#0000ff','#9900ff','#ff00ff'],
    ['#e6b8af','#f4cccc','#fce5cd','#fff2cc','#d9ead3','#d0e0e3','#c9daf8','#cfe2f3','#d9d2e9','#ead1dc'],
    ['#dd7e6b','#ea9999','#f9cb9c','#ffe599','#b6d7a8','#a2c4c9','#a4c2f4','#9fc5e8','#b4a7d6','#d5a6bd'],
    ['#cc4125','#e06666','#f6b26b','#ffd966','#93c47d','#76a5af','#6d9eeb','#6fa8dc','#8e7cc3','#c27ba0'],
    ['#a61c00','#cc0000','#e69138','#f1c232','#6aa84f','#45818e','#3c78d8','#3d85c6','#674ea7','#a64d79'],
    ['#85200c','#990000','#b45f06','#bf9000','#38761d','#134f5c','#1155cc','#0b5394','#351c75','#741b47'],
    ['#5b0f00','#660000','#783f04','#7f6000','#274e13','#0c343d','#1c4587','#073763','#20124d','#4c1130']
  ];

  var componentRelay = null;
  var workingNote = null;
  var lastSavedHTML = null;
  var wrapOn = false;
  var wrapUserSet = false;   // true once the user toggles Word Wrap themselves
  // Auto Word-Wrap heuristic: ON for touch devices (phones/tablets), and for any
  // genuinely narrow window (so it also flips as a desktop window is resized).
  var mqTouch  = window.matchMedia ? window.matchMedia('(pointer: coarse)') : null;
  var mqNarrow = window.matchMedia ? window.matchMedia('(max-width: 640px)') : null;
  function autoWrap() {
    return !!((mqTouch && mqTouch.matches) || (mqNarrow && mqNarrow.matches));
  }
  var savedRange = null;   // preserve selection while using selects / color popups
  var lastAction = null;   // repeats the last formatting action (F4)

  // Formatting commands that F4 can repeat
  var REPEATABLE = {
    bold: 1, italic: 1, underline: 1, strike: 1, sub: 1, super: 1, code: 1,
    h1: 1, h2: 1, h3: 1, h4: 1, h5: 1, h6: 1, p: 1, quote: 1,
    left: 1, center: 1, right: 1, justify: 1, ul: 1, ol: 1, checklist: 1, indent: 1, outdent: 1, clear: 1
  };

  /* ---------- Utilities ---------- */
  function escapeHtml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  var BLOCK_RE = /^(P|DIV|H1|H2|H3|H4|H5|H6|BLOCKQUOTE|UL|OL|PRE|LI|TABLE|HR)$/;
  // Wrap top-level text/inline nodes into blocks so every line is a block
  function normalizeBlocks() {
    var nodes = Array.prototype.slice.call(editor.childNodes);
    var run = null;
    for (var i = 0; i < nodes.length; i++) {
      var n = nodes[i];
      if (n.nodeType === 1 && BLOCK_RE.test(n.tagName)) { run = null; continue; }
      if (n.nodeType === 3 && !/\S/.test(n.nodeValue)) { editor.removeChild(n); continue; }
      if (!run) { run = document.createElement('p'); editor.insertBefore(run, n); }
      run.appendChild(n);
    }
  }
  // After removing a list, execCommand leaves bare inline text separated by <br> at the
  // editor's top level. Wrap each <br>-separated segment into its own <p> so every former
  // list item becomes a real block line (not a stray inline/span).
  function wrapInlineRunsAsBlocks() {
    var n = editor.firstChild;
    while (n) {
      if (n.nodeType === 1 && BLOCK_RE.test(n.tagName)) { n = n.nextSibling; continue; }
      if (n.nodeType === 3 && !/\S/.test(n.nodeValue)) { var w = n.nextSibling; editor.removeChild(n); n = w; continue; }
      var stop = n;                                   // first block after this inline run
      while (stop && !(stop.nodeType === 1 && BLOCK_RE.test(stop.tagName))) stop = stop.nextSibling;
      var seg = [], ps = [], cur = n;
      function pushP() {
        var p = document.createElement('p');
        if (seg.length) { for (var s = 0; s < seg.length; s++) p.appendChild(seg[s]); }
        else p.appendChild(document.createElement('br'));
        ps.push(p); seg = [];
      }
      while (cur !== stop) {
        var next = cur.nextSibling;
        editor.removeChild(cur);
        if (cur.nodeType === 1 && cur.tagName === 'BR') pushP();
        else seg.push(cur);
        cur = next;
      }
      if (seg.length) pushP();                        // trailing segment (skip a lone trailing <br>)
      for (var i = 0; i < ps.length; i++) editor.insertBefore(ps[i], stop);
      n = stop;
    }
  }
  function hasDirectList(el) {
    for (var i = 0; i < el.children.length; i++) {
      var t = el.children[i].tagName;
      if (t === 'UL' || t === 'OL') return true;
    }
    return false;
  }
  // A UL/OL nested inside a top-level block (e.g. execCommand's <p><ul>…</ul></p>, or a
  // heading wrapping a pasted list — <h1><ul>…</ul></h1>, which shows huge text and counts
  // as ONE line) is invalid and breaks line numbering. Lift every such list out to the
  // editor's top level, splitting the wrapper so its non-list content keeps its own tag
  // (a heading stays a heading) and each <li> becomes its own line.
  var LIFT_RE = /^(P|DIV|H1|H2|H3|H4|H5|H6|BLOCKQUOTE|PRE)$/;
  function liftLists() {
    var kids = Array.prototype.slice.call(editor.children);
    for (var i = 0; i < kids.length; i++) {
      var k = kids[i];
      if (k.nodeType !== 1 || !LIFT_RE.test(k.tagName) || !hasDirectList(k)) continue;
      var pieces = [], seg = null;
      Array.prototype.slice.call(k.childNodes).forEach(function (c) {
        if (c.nodeType === 1 && (c.tagName === 'UL' || c.tagName === 'OL')) { seg = null; pieces.push(c); }
        else if (c.nodeType === 1 && c.tagName === 'BR') { seg = null; }        // break → new segment
        else { if (!seg) { seg = k.cloneNode(false); seg.removeAttribute('class'); pieces.push(seg); } seg.appendChild(c); }
      });
      for (var p = 0; p < pieces.length; p++) {
        var pc = pieces[p], isList = pc.tagName === 'UL' || pc.tagName === 'OL';
        if (!isList && !/\S/.test(pc.textContent) && !pc.querySelector('img')) continue;  // drop empty clones
        editor.insertBefore(pc, k);
      }
      editor.removeChild(k);
    }
    wrapInlineRunsAsBlocks();
  }
  // A block that (illegally) contains other block-level elements — e.g. pasting several
  // <p> lines while the caret sits inside an <h3> yields <h3><p>…</p><p>…</p></h3>, which
  // breaks the one-block-per-line model and line numbering. Lift every nested block out to
  // the editor's top level: inline runs stay in a clone of the wrapper (a heading keeps its
  // tag), and each nested block becomes its own top-level line. Runs until nothing is nested.
  var CHILD_BLOCK_RE = /^(P|DIV|H1|H2|H3|H4|H5|H6|BLOCKQUOTE|PRE|UL|OL|TABLE|HR)$/;
  function hasDirectBlockChild(el) {
    for (var i = 0; i < el.children.length; i++) if (CHILD_BLOCK_RE.test(el.children[i].tagName)) return true;
    return false;
  }
  function flattenNestedBlocks() {
    for (var pass = 0; pass < 20; pass++) {
      var kids = Array.prototype.slice.call(editor.children), didWork = false;
      for (var i = 0; i < kids.length; i++) {
        var k = kids[i];
        if (k.nodeType !== 1 || !LIFT_RE.test(k.tagName) || !hasDirectBlockChild(k)) continue;
        var pieces = [], seg = null;
        Array.prototype.slice.call(k.childNodes).forEach(function (c) {
          if (c.nodeType === 1 && CHILD_BLOCK_RE.test(c.tagName)) { seg = null; pieces.push(c); }
          else if (c.nodeType === 1 && c.tagName === 'BR') { seg = null; }        // break → new segment
          else { if (!seg) { seg = k.cloneNode(false); seg.removeAttribute('class'); pieces.push(seg); } seg.appendChild(c); }
        });
        for (var p = 0; p < pieces.length; p++) {
          var pc = pieces[p], isBlk = pc.nodeType === 1 && CHILD_BLOCK_RE.test(pc.tagName);
          if (!isBlk && !/\S/.test(pc.textContent) && !pc.querySelector('img')) continue;  // drop empty clones
          editor.insertBefore(pc, k);
        }
        editor.removeChild(k);
        didWork = true;
      }
      if (!didWork) break;
    }
  }
  function ensureContent() {
    if (editor.children.length === 0) editor.innerHTML = '<p><br></p>';
  }
  // Split a block that packs several lines with <br> (e.g. execCommand's multi-line
  // <blockquote>Line1<br>Line2</blockquote>) into one block PER LINE of the same tag, so
  // each line gets its own gutter number. A lone filler <br> (an empty line) is kept as is.
  function splitBlockOnBr(block) {
    var segs = [[]], nodes = Array.prototype.slice.call(block.childNodes), hasBr = false;
    for (var i = 0; i < nodes.length; i++) {
      if (nodes[i].nodeType === 1 && nodes[i].tagName === 'BR') { hasBr = true; segs.push([]); }
      else segs[segs.length - 1].push(nodes[i]);
    }
    if (!hasBr || nodes.length <= 1) return false;          // no split for a lone/empty <br>
    if (segs.length > 1 && segs[segs.length - 1].length === 0) segs.pop();   // drop trailing empty segment
    var keepClass = (block.className || '').replace(/\bcurrent-line\b/, '').trim();
    var style = block.getAttribute('style');
    var frag = document.createDocumentFragment();
    for (var s = 0; s < segs.length; s++) {
      var nb = document.createElement(block.tagName);
      if (keepClass) nb.className = keepClass;
      if (style) nb.setAttribute('style', style);
      if (segs[s].length) { for (var j = 0; j < segs[s].length; j++) nb.appendChild(segs[s][j]); }
      else nb.appendChild(document.createElement('br'));
      frag.appendChild(nb);
    }
    block.parentNode.replaceChild(frag, block);
    return true;
  }
  // Keep every multi-line paragraph-style block (quote/heading/normal) as one block per
  // line. This mirrors how lists (one <li> per line) and code blocks already number lines.
  var SPLIT_RE = /^(P|DIV|BLOCKQUOTE|H1|H2|H3|H4|H5|H6)$/;
  function splitMultilineBlocks() {
    var kids = Array.prototype.slice.call(editor.children);
    for (var i = 0; i < kids.length; i++) {
      var k = kids[i];
      if (k.nodeType === 1 && SPLIT_RE.test(k.tagName)) splitBlockOnBr(k);
    }
  }
  // Run a DOM mutation while preserving the caret, using a temporary zero-width marker
  // that travels with its text as nodes are moved into new blocks.
  function withCaretPreserved(mutate) {
    var sel = window.getSelection(), marker = null;
    if (sel.rangeCount && editor.contains(sel.anchorNode)) {
      marker = document.createElement('span');
      marker.setAttribute('data-caret-marker', '1');
      marker.appendChild(document.createTextNode('​'));
      sel.getRangeAt(0).insertNode(marker);
    }
    mutate();
    if (marker && marker.parentNode) {
      var host = marker.parentNode, next = marker.nextSibling;
      host.removeChild(marker);
      var r = document.createRange();
      if (next && next.parentNode === host) { r.setStartBefore(next); }
      else { r.selectNodeContents(host); r.collapse(false); }
      r.collapse(true);
      if (!host.childNodes.length) host.appendChild(document.createElement('br'));
      sel.removeAllRanges(); sel.addRange(r);
      host.normalize();
    }
  }
  // Insert an array of text lines: multiple -> a block each; a single -> inline
  function insertLines(lines) {
    if (!lines || !lines.length) return;
    if (lines.length > 1 && lines[lines.length - 1] === '') lines.pop();
    editor.focus();
    if (lines.length === 1) {
      if (textHasUrl(lines[0])) document.execCommand('insertHTML', false, escapeAndLinkify(lines[0]));
      else document.execCommand('insertText', false, lines[0]);
    } else {
      var html = lines.map(function (l) {
        return '<p>' + (l ? escapeAndLinkify(l) : '<br>') + '</p>';
      }).join('');
      document.execCommand('insertHTML', false, html);
      stripSpuriousBg();
      absorbLeadingTabs();            // leading \t on a pasted line → editor indent
    }
    onChange();
  }
  // Insert pasted text: multiple lines -> multiple blocks; a single line -> inline
  function insertPastedText(text) {
    if (text == null || text === '') return;
    insertLines(text.replace(/\r\n?/g, '\n').split('\n'));
  }
  // Plain "value only" paste: collapse blank-line runs, because copying rich blocks
  // puts a BLANK LINE between paragraphs in text/plain (Chromium) — and Ctrl+Shift+V
  // gives us only text/plain (no HTML) so we can't use the block structure here.
  function insertValueText(text) {
    if (text == null || text === '') return;
    insertLines(text.replace(/\r\n?/g, '\n').replace(/\n{2,}/g, '\n').split('\n'));
  }
  /* ---------- Formatted paste (keep bold/colors/links) with sanitising ---------- */
  var PASTE_TAGS = { P:1, DIV:1, BR:1, HR:1, SPAN:1, B:1, STRONG:1, I:1, EM:1, U:1, S:1, STRIKE:1, DEL:1,
    MARK:1, CODE:1, PRE:1, A:1, H1:1, H2:1, H3:1, H4:1, H5:1, H6:1, UL:1, OL:1, LI:1, BLOCKQUOTE:1, FONT:1, SUB:1, SUP:1,
    TABLE:1, THEAD:1, TBODY:1, TFOOT:1, TR:1, TD:1, TH:1, CAPTION:1, COLGROUP:1, COL:1 };
  var STYLE_KEEP = ['color', 'background-color', 'font-weight', 'font-style',
    'text-decoration', 'text-decoration-line', 'font-size', 'font-family', 'text-align',
    'width', 'height', 'table-layout', 'vertical-align'];

  function filterStyle(el) {
    var out = '';
    for (var i = 0; i < STYLE_KEEP.length; i++) {
      var v = el.style.getPropertyValue(STYLE_KEEP[i]);
      if (v) out += STYLE_KEEP[i] + ':' + v + ';';
    }
    return out;
  }
  // Strip scripts/styles/classes/handlers; keep a safe subset of tags + inline formatting
  function sanitizeHtml(html) {
    var box = document.createElement('div');
    box.innerHTML = String(html || '');
    // Remove comment nodes — clipboard HTML carries <!--StartFragment-->/<!--EndFragment-->
    // markers that would otherwise become spurious empty lines when pasted.
    var walker = document.createTreeWalker(box, NodeFilter.SHOW_COMMENT, null);
    var comments = [], cm;
    while ((cm = walker.nextNode())) comments.push(cm);
    for (var ci = 0; ci < comments.length; ci++) if (comments[ci].parentNode) comments[ci].parentNode.removeChild(comments[ci]);
    var drop = box.querySelectorAll('script,style,meta,link,title,head,object,embed,iframe,noscript,input,button,form,svg,img');
    for (var d = 0; d < drop.length; d++) if (drop[d].parentNode) drop[d].parentNode.removeChild(drop[d]);
    var els = box.querySelectorAll('*');
    for (var j = els.length - 1; j >= 0; j--) {          // reverse: unwrap children before parents
      var el = els[j];
      if (!PASTE_TAGS[el.tagName]) {                      // unwrap unknown tags, keep their content
        var parent = el.parentNode;
        while (el.firstChild) parent.insertBefore(el.firstChild, el);
        parent.removeChild(el);
        continue;
      }
      var attrs = Array.prototype.slice.call(el.attributes);
      for (var k = 0; k < attrs.length; k++) {
        var name = attrs[k].name.toLowerCase();
        if (name === 'href' && el.tagName === 'A') {
          if (/^\s*javascript:/i.test(attrs[k].value)) el.removeAttribute('href');
        } else if (name === 'color' && el.tagName === 'FONT') {
          el.style.color = attrs[k].value; el.removeAttribute('color');
        } else if (name === 'style') {
          var s = filterStyle(el);
          if (s) el.setAttribute('style', s); else el.removeAttribute('style');
        } else if (name === 'class' && el.tagName === 'UL' && /\brn-checklist\b/.test(attrs[k].value)) {
          el.setAttribute('class', 'rn-checklist');            // keep checklist lists intact on paste
        } else if (name === 'class' && el.tagName === 'PRE' && /\brn-code\b/.test(attrs[k].value)) {
          el.setAttribute('class', 'rn-code');                 // keep code blocks intact on paste
        } else if (name === 'data-lang' && el.tagName === 'PRE') {
          /* keep a code block's language */
        } else if (name === 'data-checked' && el.tagName === 'LI') {
          /* keep a checklist item's checked state */
        } else {
          el.removeAttribute(attrs[k].name);
        }
      }
    }
    return box.innerHTML;
  }
  // Chromium's insertHTML/insertText sometimes injects a spurious white background on
  // inserted inline nodes — strip white/transparent backgrounds so pasted text looks clean.
  function stripSpuriousBg() {
    var els = editor.querySelectorAll('[style*="background"]');
    for (var i = 0; i < els.length; i++) {
      var bg = els[i].style.backgroundColor;
      if (bg === 'rgb(255, 255, 255)' || bg === 'white' || bg === 'transparent' || bg === 'rgba(0, 0, 0, 0)') {
        els[i].style.removeProperty('background-color');
        if (!els[i].getAttribute('style')) els[i].removeAttribute('style');
      }
    }
  }
  function pasteHtml(html) {
    var clean = sanitizeHtml(html);
    if (!clean) return;
    editor.focus();
    document.execCommand('insertHTML', false, clean);
    stripSpuriousBg();
    normalizeBlocks();
    liftLists();
    flattenNestedBlocks();              // e.g. <h3><p>…</p></h3> from pasting into a heading
    splitMultilineBlocks();             // one block per line for pasted multi-line quotes
    absorbLeadingTabs();                // leading \t → editor indent (mid-line tabs untouched)
    styleTables();
    ensureContent();
    highlightAllCode();                 // colour any pasted code blocks
    onChange();
  }

  // Turn pasted HTML into one plain line PER BLOCK. Used for "value only" so that
  // copying N paragraphs pastes N lines — not 2N (Chromium's text/plain separates
  // blocks with a blank line, which would otherwise add an empty line between each).
  var LINE_BLOCK = /^(P|DIV|H[1-6]|BLOCKQUOTE|PRE|LI|TR)$/;
  function collectLines(container, lines) {
    var run = null;   // accumulates a run of inline nodes into one line
    function flush() { if (run !== null) { lines.push(run); run = null; } }
    Array.prototype.forEach.call(container.childNodes, function (node) {
      if (node.nodeType === 1 && (node.tagName === 'UL' || node.tagName === 'OL')) {
        flush(); collectLines(node, lines);
      } else if (node.nodeType === 1 && LINE_BLOCK.test(node.tagName)) {
        flush(); lines.push(node.textContent.replace(/\s*\n\s*/g, ' '));
      } else if (node.nodeType === 1 && node.tagName === 'BR') {
        lines.push(run || ''); run = null;
      } else {
        run = (run || '') + (node.textContent || node.nodeValue || '');
      }
    });
    flush();
  }
  function pasteValueFromHtml(html) {
    var box = document.createElement('div');
    box.innerHTML = sanitizeHtml(html);
    var lines = [];
    collectLines(box, lines);
    if (lines.length) insertLines(lines);
    else insertPastedText(box.textContent);
  }

  // Async Clipboard API read — the last-resort fallback for normal browsers where
  // execCommand('paste') is disabled AND no native paste event fires (menu click).
  // formatted=true keeps HTML formatting; formatted=false pastes value only.
  function asyncClipboardPaste(formatted) {
    var clip = navigator.clipboard;
    function readPlain() {
      if (clip && clip.readText) {
        clip.readText().then(function (t) { (formatted ? insertPastedText : insertValueText)(t); })
          .catch(function () {});
      }
    }
    if (clip && clip.read) {
      clip.read().then(function (items) {
        for (var i = 0; i < items.length; i++) {
          if (items[i].types.indexOf('text/html') !== -1) {
            items[i].getType('text/html').then(function (b) {
              b.text().then(formatted ? pasteHtml : pasteValueFromHtml);
            });
            return;
          }
        }
        readPlain();
      }).catch(readPlain);
    } else {
      readPlain();
    }
  }

  // Read the clipboard and paste. formatted=true keeps HTML formatting; formatted=false
  // pastes plain text ("value only") but still uses the HTML block structure (one line
  // per block) so multi-paragraph copies don't gain blank lines.
  function pasteFromClipboard(formatted) {
    editor.focus();
    // Electron / Standard Notes desktop: an execCommand('paste') inside the current user
    // gesture (menu click / keydown) fires a TRUSTED 'paste' event whose clipboardData
    // needs no permission. Electron may fire that event ASYNCHRONOUSLY, so we must keep
    // pasteOverride set until the event actually arrives (the paste handler clears it and
    // cancels the timer). If no paste event comes, the timer falls back to the async API.
    armPasteOverride(formatted ? 'format' : 'plain');
    try { document.execCommand('paste'); } catch (e) {}
  }
  function doPaste()        { pasteFromClipboard(true); }   // default paste — keep formatting
  function pasteValueOnly() { pasteFromClipboard(false); }  // menu: Paste value only

  // Tag the next trusted 'paste' event with the desired mode. Used by menu paste,
  // Ctrl+Shift+V, and execCommand('paste'). Kept armed for a short window so a paste
  // event that Electron fires asynchronously still lands in the right mode; if none
  // arrives, the timer falls back to the async Clipboard API (normal browsers).
  var pasteOverride = null;      // 'plain' | 'format'
  var pasteOverrideTimer = null;
  function armPasteOverride(mode) {
    pasteOverride = mode;
    if (pasteOverrideTimer) clearTimeout(pasteOverrideTimer);
    pasteOverrideTimer = setTimeout(function () {
      if (pasteOverride === mode) { pasteOverride = null; asyncClipboardPaste(mode === 'format'); }
    }, 200);
  }
  function blockOf(node) {
    var n = node;
    if (!editor.contains(n)) return null;
    while (n && n.parentNode !== editor) n = n.parentNode;
    return (n && n.parentNode === editor) ? n : null;
  }
  function currentBlock() {
    var sel = window.getSelection();
    if (!sel.rangeCount) return null;
    return blockOf(sel.anchorNode);
  }
  // "Line" units for numbering/highlight: every top-level block, but a UL/OL expands so
  // each <li> counts as its own line.
  function lineUnits() {
    var out = [], kids = editor.children;
    for (var i = 0; i < kids.length; i++) {
      var el = kids[i];
      if (el.tagName === 'UL' || el.tagName === 'OL') {
        var lis = el.querySelectorAll('li');
        for (var j = 0; j < lis.length; j++) out.push(lis[j]);
      } else out.push(el);
    }
    return out;
  }
  function currentListItem() {
    var sel = window.getSelection();
    if (!sel.rangeCount || !editor.contains(sel.anchorNode)) return null;
    return closestTag(sel.anchorNode, 'LI');
  }
  // The line unit holding the caret: the innermost <li>, else the top-level block.
  function currentLineUnit() {
    var sel = window.getSelection();
    if (!sel.rangeCount || !editor.contains(sel.anchorNode)) return null;
    var n = sel.anchorNode, li = null, block = null;
    while (n && n !== editor) {
      if (n.nodeType === 1 && n.tagName === 'LI' && !li) li = n;
      if (n.parentNode === editor) block = n;
      n = n.parentNode;
    }
    return li || block;
  }
  function selectedBlocks() {
    var sel = window.getSelection();
    if (!sel.rangeCount) return [];
    var r = sel.getRangeAt(0);
    var a = blockOf(r.startContainer), b = blockOf(r.endContainer);
    var blocks = Array.prototype.slice.call(editor.children);
    var i1 = blocks.indexOf(a), i2 = blocks.indexOf(b);
    if (i1 < 0 && i2 < 0) return [];
    if (i1 < 0) i1 = i2;
    if (i2 < 0) i2 = i1;
    return blocks.slice(Math.min(i1, i2), Math.max(i1, i2) + 1);
  }
  function saveSel() {
    var s = window.getSelection();
    if (s.rangeCount && editor.contains(s.anchorNode)) savedRange = s.getRangeAt(0).cloneRange();
  }
  function restoreSel() {
    if (!savedRange) { editor.focus(); return; }
    editor.focus();
    var s = window.getSelection();
    s.removeAllRanges();
    s.addRange(savedRange);
  }

  /* ---------- Line-number gutter ----------
     Number by VISUAL line: merge blocks that share an offsetTop (inline nodes
     on the same line) and skip zero-height blocks so numbers never overlap. */
  function renderGutter(cur) {
    var blocks = lineUnits();
    var lines = [];
    var seen = {};
    for (var i = 0; i < blocks.length; i++) {
      var b = blocks[i];
      if (b.offsetHeight === 0) continue;                 // skip empty zero-height blocks
      var cs = window.getComputedStyle(b);
      var lh = parseFloat(cs.lineHeight) || 26;
      // A code block numbers EVERY text line, not just the block as a whole.
      if (isCodeBlock(b)) {
        var codeEl = codeTextEl(b);
        // offsetTop is the border-box top (margin excluded); align to the first text line.
        var base = b.offsetTop + (parseFloat(cs.borderTopWidth) || 0) + (parseFloat(cs.paddingTop) || 0);
        var txt = codeSrc(b);
        var nLines = txt.length ? txt.split('\n').length : 1;
        var activeLine = (b === cur) ? codeCaretLine(codeEl) : -1;
        for (var cl = 0; cl < nLines; cl++) lines.push({ top: base + cl * lh, lh: lh, active: cl === activeLine });
        continue;
      }
      var top = b.offsetTop + (parseFloat(cs.paddingTop) || 0) + (parseFloat(cs.marginTop) || 0);
      var key = Math.round(top);
      if (seen[key] === undefined) { seen[key] = lines.length; lines.push({ top: top, lh: lh, active: false }); }
      if (b === cur) lines[seen[key]].active = true;      // current line
    }
    if (!lines.length) lines.push({ top: 6, lh: 26, active: true });
    var html = '';
    for (var j = 0; j < lines.length; j++) {
      html += '<div class="ln' + (lines[j].active ? ' active' : '') + '" style="top:' + lines[j].top +
              'px;height:' + lines[j].lh + 'px;line-height:' + lines[j].lh + 'px">' + (j + 1) + '</div>';
    }
    gutterInner.innerHTML = html;
    gutterInner.style.height = editor.scrollHeight + 'px';
    // Compact gutter on mobile: narrower padding so line numbers take less width
    var gutterPad = (window.matchMedia && window.matchMedia('(max-width: 640px)').matches) ? 8 : 22;
    gutter.style.width = 'calc(' + String(lines.length).length + 'ch + ' + gutterPad + 'px)';
    syncGutterScroll();
  }
  // Keep the numbers aligned with the content. Translating the inner layer tracks the
  // editor's scroll exactly — scrollTop on the overflow:hidden gutter doesn't stick.
  function syncGutterScroll() {
    gutterInner.style.transform = 'translateY(' + (-editor.scrollTop) + 'px)';
  }

  /* ---------- Status bar ---------- */
  function updateStatus(cur) {
    var blocks = lineUnits();
    var seen = {}, count = 0, ln = 1, col = 1;
    var sel = window.getSelection();
    for (var i = 0; i < blocks.length; i++) {
      var b = blocks[i];
      if (b.offsetHeight === 0) continue;
      if (isCodeBlock(b)) {
        var codeEl = codeTextEl(b);
        var ctxt = codeSrc(b);
        var n = ctxt.length ? ctxt.split('\n').length : 1;
        if (b === cur) {
          var off = caretOffsetIn(codeEl); if (off == null) off = 0;
          var before = ctxt.slice(0, off);
          ln = count + (before.match(/\n/g) || []).length + 1;
          col = before.length - before.lastIndexOf('\n');   // chars after the last newline (+1)
        }
        count += n;
        continue;
      }
      var key = Math.round(b.offsetTop);
      if (seen[key] === undefined) seen[key] = ++count;
      if (b === cur) {
        ln = seen[key];
        if (sel.rangeCount) {
          var r = sel.getRangeAt(0);
          var pre = r.cloneRange();
          pre.selectNodeContents(cur);
          pre.setEnd(r.endContainer, r.endOffset);
          col = pre.toString().length + 1;
        }
      }
    }
    stPos.innerHTML = 'Ln : ' + ln + '&nbsp;&nbsp;Col : ' + col;
    var txt = editor.textContent;
    var words = (txt.match(/\S+/g) || []).length;
    stLen.innerHTML = 'words : ' + words + '&nbsp;&nbsp;length : ' + txt.length +
                      '&nbsp;&nbsp;lines : ' + Math.max(1, count);
    var selLen = (sel && sel.rangeCount) ? sel.toString().length : 0;
    stSel.textContent = selLen > 0 ? ('Sel : ' + selLen) : '';
  }

  // Sync the gutter to the editor's scroll, throttled to one update per animation
  // frame so fast scrolling stays smooth (no per-event layout thrash).
  var gutterSyncRaf = false;
  editor.addEventListener('scroll', function () {
    if (gutterSyncRaf) return;
    gutterSyncRaf = true;
    requestAnimationFrame(function () {
      gutterSyncRaf = false;
      syncGutterScroll();
      if (codeBarFor && codeBar && codeBar.classList.contains('show')) positionCodeBar(codeBarFor);
    });
  }, { passive: true });
  window.addEventListener('resize', function () {
    if (codeBarFor && codeBar && codeBar.classList.contains('show')) positionCodeBar(codeBarFor);
  });

  /* ---------- Toolbar button states ---------- */
  function q(cmd) { try { return document.queryCommandState(cmd); } catch (e) { return false; } }

  // Run a command that MUST emit a real element (<u>/<strike>/<sub>/<sup>) instead of a
  // styleWithCSS span. A text-decoration / vertical-align CSS span can't be detected by the
  // browser to toggle back OFF, so re-running the shortcut would fail to remove the format.
  // Emit the tag with styleWithCSS off, then restore it to the editor's default (true).
  function execTag(cmd) {
    document.execCommand('styleWithCSS', false, false);
    document.execCommand(cmd);
    document.execCommand('styleWithCSS', false, true);
  }
  function inCode() {
    var sel = window.getSelection();
    if (!sel.rangeCount) return false;
    var n = sel.anchorNode;
    while (n && n !== editor) { if (n.nodeType === 1 && n.tagName === 'CODE') return true; n = n.parentNode; }
    return false;
  }
  function setActive(cmd, on) {
    var btn = toolbar.querySelector('.tbtn[data-cmd="' + cmd + '"]');
    if (btn) { btn.classList.toggle('active', !!on); btn.setAttribute('aria-pressed', on ? 'true' : 'false'); }
  }
  function setEnabled(cmd, on) {
    var els = document.querySelectorAll('.tbtn[data-cmd="' + cmd + '"], .menu-item[data-cmd="' + cmd + '"]');
    for (var i = 0; i < els.length; i++) {
      els[i].disabled = !on;
      els[i].classList.toggle('disabled', !on);
    }
  }
  /* ---------- Toolbar dropdowns (align / lists / more-format) ---------- */
  // Toggle .active on a dropdown option (the .tdrop-opt for a given command)
  function setOpt(pop, cmd, on) {
    if (!pop) return;
    var opt = pop.querySelector('.tdrop-opt[data-cmd="' + cmd + '"]');
    if (opt) opt.classList.toggle('active', !!on);
  }
  var ALIGN_ICONS = {
    left:    '<svg class="ico" viewBox="0 0 16 16"><path d="M2 4h12M2 8h8M2 12h12"/></svg>',
    center:  '<svg class="ico" viewBox="0 0 16 16"><path d="M2 4h12M4 8h8M2 12h12"/></svg>',
    right:   '<svg class="ico" viewBox="0 0 16 16"><path d="M2 4h12M6 8h8M2 12h12"/></svg>',
    justify: '<svg class="ico" viewBox="0 0 16 16"><path d="M2 4h12M2 8h12M2 12h12"/></svg>'
  };
  var LIST_ICONS = {
    ul: '<svg class="ico" viewBox="0 0 16 16"><circle class="dot" cx="2.6" cy="4" r="1.1"/><circle class="dot" cx="2.6" cy="8" r="1.1"/><circle class="dot" cx="2.6" cy="12" r="1.1"/><path d="M6 4h8M6 8h8M6 12h8"/></svg>',
    ol: '<svg class="ico" viewBox="0 0 16 16"><path d="M6 4h8M6 8h8M6 12h8"/><text x="0.4" y="5.6">1</text><text x="0.4" y="9.6">2</text><text x="0.4" y="13.6">3</text></svg>'
  };
  function currentAlign() {
    if (q('justifyCenter')) return 'center';
    if (q('justifyRight'))  return 'right';
    if (q('justifyFull'))   return 'justify';
    return 'left';
  }
  function updateAlign() {
    if (!alignBtn) return;
    var cur = currentAlign();
    if (alignCur) alignCur.innerHTML = ALIGN_ICONS[cur] || ALIGN_ICONS.left;
    alignBtn.classList.toggle('active', cur !== 'left');
    setOpt(alignPop, 'left', cur === 'left');
    setOpt(alignPop, 'center', cur === 'center');
    setOpt(alignPop, 'right', cur === 'right');
    setOpt(alignPop, 'justify', cur === 'justify');
  }
  function updateLists() {
    if (!listBtn) return;
    var ul = q('insertUnorderedList'), ol = q('insertOrderedList'), chk = inChecklist();
    if (listCur) listCur.innerHTML = ol ? LIST_ICONS.ol : LIST_ICONS.ul;   // reflect active list; bullet as neutral default
    listBtn.classList.toggle('active', ul || ol || chk);
    setOpt(listPop, 'ul', ul && !chk);   // a checklist is a <ul>; don't light both up
    setOpt(listPop, 'ol', ol);
    setOpt(listPop, 'checklist', chk);
  }
  function updateFmtExtra() {
    if (!fmtBtn) return;
    var st = q('strikeThrough'), su = q('subscript'), sp = q('superscript');
    setOpt(fmtPop, 'strike', st);
    setOpt(fmtPop, 'sub', su);
    setOpt(fmtPop, 'super', sp);
    fmtBtn.classList.toggle('active', st || su || sp);
  }

  // The paragraph-style tag of a single top-level block (DIV counts as P; a code block
  // as CODE; a list as itself — lists aren't a paragraph style).
  function styleTagOf(block) {
    if (!block || block.nodeType !== 1) return 'P';
    var tn = block.tagName;
    if (tn === 'PRE' && block.classList.contains('rn-code')) return 'CODE';
    if (/^(H1|H2|H3|H4|H5|H6|BLOCKQUOTE)$/.test(tn)) return tn;
    if (tn === 'DIV' || tn === 'P') return 'P';
    return tn;   // UL/OL/TABLE/HR — not a paragraph style
  }
  // The common paragraph style across the whole selection, computed from the actual blocks
  // (queryCommandValue('formatBlock') returns '' for multi-block selections, so it can't be
  // trusted). Returns 'CODE' inside a code block, a tag if all lines share it, else '' (mixed).
  function selectedBlockStyle() {
    var sel = window.getSelection();
    if (sel.rangeCount && codeBlockOf(sel.anchorNode)) return 'CODE';
    var blocks = selectedBlocks();
    if (!blocks.length) { var cb = currentBlock(); blocks = cb ? [cb] : []; }
    if (!blocks.length) return 'P';
    var tag = null;
    for (var i = 0; i < blocks.length; i++) {
      var t = styleTagOf(blocks[i]);
      if (/^(UL|OL|TABLE|HR)$/.test(t)) t = 'P';        // lists/tables read as neutral for the style select
      if (tag === null) tag = t; else if (tag !== t) return '';   // mixed
    }
    return tag || 'P';
  }
  function updateToolbar() {
    setActive('bold', q('bold'));
    setActive('italic', q('italic'));
    setActive('underline', q('underline'));
    updateFmtExtra();   // strike / sub / super / code (grouped dropdown)
    updateLists();      // bullet / numbered (grouped dropdown)
    updateAlign();      // left / center / right / justify (grouped dropdown)
    // Block style computed from the actual selected blocks (reliable for multi-line
    // selections, unlike queryCommandValue('formatBlock')).
    var style = selectedBlockStyle();
    setActive('h1', style === 'H1');
    setActive('h2', style === 'H2');
    setActive('p', style === 'P');
    setActive('quote', style === 'BLOCKQUOTE');
    setActive('link', !!currentLinkEl());
    setActive('wrap', wrapOn);
    if (wrapItem) wrapItem.classList.toggle('checked', wrapOn);

    // Disable Undo/Redo when there is nothing to undo/redo (driven by our own history)
    updateHistButtons();

    // Sync the paragraph-style select: the shared style, or blank when the lines are mixed.
    styleSelect.value = style;   // '' (no option) shows blank for a mixed selection

    // Sync the font & size selects + the colour swatches with the selection (Word-like)
    var el = selEl();
    if (el) {
      var cs = window.getComputedStyle(el);
      setSelectFont(cs.fontFamily);
      setSelectSize(Math.round(parseFloat(cs.fontSize)));
      foreBar.style.background = cs.color;                 // text colour (inherited)
      backBar.style.background = effectiveBg(el);          // highlight colour
    }
  }

  // Nearest inline highlight colour above the caret. Skips the block itself so the
  // current-line highlight (a UI affordance on the <p>) isn't mistaken for a text highlight.
  function effectiveBg(node) {
    var n = node && node.nodeType === 3 ? node.parentElement : node;
    while (n && n !== editor && n.parentNode !== editor) {   // stop before the top-level block
      var bg = window.getComputedStyle(n).backgroundColor;
      if (bg && bg !== 'rgba(0, 0, 0, 0)' && bg !== 'transparent') return bg;
      n = n.parentElement;
    }
    return 'transparent';
  }

  function selEl() {
    var sel = window.getSelection();
    if (!sel.rangeCount || !editor.contains(sel.anchorNode)) return null;
    var n = sel.anchorNode;
    return n.nodeType === 3 ? n.parentElement : n;
  }
  function primaryFont(ff) { return (ff || '').split(',')[0].replace(/["']/g, '').trim().toLowerCase(); }
  function setSelectFont(ff) {
    var want = primaryFont(ff), opts = fontSelect.options, idx = 0;
    for (var i = 1; i < opts.length; i++) { if (primaryFont(opts[i].value) === want) { idx = i; break; } }
    fontSelect.selectedIndex = idx;
  }
  function setSelectSize(px) {
    // sizeSelect is an editable input — don't overwrite what the user is typing
    if (document.activeElement === sizeSelect) return;
    sizeSelect.value = px;
  }

  function refresh() {
    ensureContent();
    var cur = currentLineUnit();                    // the current line (an <li> or a top block)
    var prev = editor.querySelectorAll('.current-line');
    for (var i = 0; i < prev.length; i++) if (prev[i] !== cur) prev[i].classList.remove('current-line');
    if (cur) cur.classList.add('current-line');
    renderGutter(cur);
    updateStatus(cur);
    if (codeBarFor) positionCodeBar(codeBarFor);   // keep the hover bar glued as the block reflows
    updateToolbar();
    var topCur = currentBlock();
    updateTableTool(topCur);
    if (cellSel.length && !dragging && topCur !== cellSelTable) clearCellSel();
    updateEmptyState();
  }

  /* ---------- Undo / redo history ----------
     Native execCommand('undo') only tracks edits made THROUGH execCommand, so the many
     DIRECT DOM mutations here (indent padding, tables, code blocks, checklist toggles,
     line moves…) bypassed it — undo was unreliable and a multi-line indent couldn't be
     undone at all. This is a self-contained stack of {html, selection} snapshots: a burst
     of typing coalesces into one entry, and every discrete command records exactly one
     entry, so a single Ctrl+Z reverts a whole multi-line Tab / indent / paste at once. */
  var histStack = [], histIdx = -1, histTimer = null, histLock = false;
  var HIST_MAX = 300, HIST_DELAY = 350;

  function histNodePath(node) {
    if (!node || !editor.contains(node)) return null;
    var path = [], n = node;
    while (n && n !== editor) {
      var p = n.parentNode;
      path.unshift(Array.prototype.indexOf.call(p.childNodes, n));
      n = p;
    }
    return path;
  }
  function histResolve(path) {
    if (!path) return null;
    var n = editor;
    for (var i = 0; i < path.length; i++) {
      if (!n.childNodes || path[i] < 0 || path[i] >= n.childNodes.length) return null;
      n = n.childNodes[path[i]];
    }
    return n;
  }
  function histSaveSel() {
    var s = window.getSelection();
    if (!s.rangeCount || !editor.contains(s.anchorNode)) return null;
    return { ap: histNodePath(s.anchorNode), ao: s.anchorOffset,
             fp: histNodePath(s.focusNode),  fo: s.focusOffset };
  }
  function histRestoreSel(sel) {
    if (!sel) return;
    var an = histResolve(sel.ap), fn = histResolve(sel.fp), fo = sel.fo;
    if (!an) return;
    if (!fn) { fn = an; fo = sel.ao; }
    var maxA = an.nodeType === 3 ? an.nodeValue.length : an.childNodes.length;
    var maxF = fn.nodeType === 3 ? fn.nodeValue.length : fn.childNodes.length;
    var s = window.getSelection();
    try {
      var r = document.createRange();
      r.setStart(an, Math.min(sel.ao, maxA));
      r.setEnd(fn, Math.min(fo, maxF));
      s.removeAllRanges(); s.addRange(r);
    } catch (e) {
      try {                                    // reversed range → fall back to a collapsed caret
        var r2 = document.createRange();
        r2.setStart(an, Math.min(sel.ao, maxA)); r2.collapse(true);
        s.removeAllRanges(); s.addRange(r2);
      } catch (e2) {}
    }
  }
  // Strip purely-visual state (current-line highlight, cell-selection) so snapshots compare
  // by CONTENT only — moving the caret to another line must not create a bogus undo entry.
  // These are class tokens, so removing them never changes node structure → saved selection
  // paths still resolve after a restore.
  function histClean(html) {
    return html.replace(/\s*\b(?:current-line|rn-cell-sel)\b/g, '').replace(/\sclass=""/g, '');
  }
  function histCommit() {
    if (histLock) return;
    var html = histClean(editor.innerHTML);
    if (histIdx >= 0 && histStack[histIdx].html === html) {
      histStack[histIdx].sel = histSaveSel();       // refresh the caret, no new entry
      return;
    }
    histStack.length = histIdx + 1;                 // drop any redo tail
    histStack.push({ html: html, sel: histSaveSel() });
    if (histStack.length > HIST_MAX) histStack.shift();
    histIdx = histStack.length - 1;
    updateHistButtons();
  }
  function histCommitDebounced() {
    if (histTimer) clearTimeout(histTimer);
    histTimer = setTimeout(function () { histTimer = null; histCommit(); }, HIST_DELAY);
  }
  function histFlush() { if (histTimer) { clearTimeout(histTimer); histTimer = null; histCommit(); } }
  function histReset() {
    if (histTimer) { clearTimeout(histTimer); histTimer = null; }
    histStack = [{ html: histClean(editor.innerHTML), sel: histSaveSel() }];
    histIdx = 0;
    updateHistButtons();
  }
  function histApply(entry) {
    histLock = true;
    editor.focus();
    editor.innerHTML = entry.html;
    histRestoreSel(entry.sel);
    histLock = false;
    highlightAllCode();                             // re-attach live code-block colouring
    refresh();
    save();
    updateHistButtons();
  }
  function undo() {
    histFlush();
    histCommit();                                   // pin the latest edit as the top of the stack
    if (histIdx <= 0) return;
    histIdx--;
    histApply(histStack[histIdx]);
  }
  function redo() {
    histFlush();
    if (histIdx >= histStack.length - 1) return;
    histIdx++;
    histApply(histStack[histIdx]);
  }
  function updateHistButtons() {
    setEnabled('undo', histIdx > 0);
    setEnabled('redo', histIdx >= 0 && histIdx < histStack.length - 1);
  }

  function onChange() { if (!histLock) histCommit(); refresh(); save(); }
  editor.addEventListener('input', function () {
    if (!histLock) histCommitDebounced();
    refresh(); save();
  });

  document.addEventListener('selectionchange', function () {
    var sel = window.getSelection();
    if (!sel.rangeCount || !editor.contains(sel.anchorNode)) return;
    refresh();
  });

  /* ============================================================
     FORMATTING
     ============================================================ */
  function closestTag(node, tag) {
    var n = node;
    while (n && n !== editor) { if (n.nodeType === 1 && n.tagName === tag) return n; n = n.parentNode; }
    return null;
  }
  // Toggle code like Bold: clicking again removes <code>. Uses the Range API to build
  // a REAL <code> (execCommand insertHTML turns it into a <span>, which can't be toggled off).
  function toggleCode() {
    var sel = window.getSelection();
    if (!sel.rangeCount) return;
    var codeEl = closestTag(sel.anchorNode, 'CODE') || closestTag(sel.focusNode, 'CODE');
    if (codeEl) {
      var parent = codeEl.parentNode;
      var first = codeEl.firstChild, last = codeEl.lastChild;
      while (codeEl.firstChild) parent.insertBefore(codeEl.firstChild, codeEl);
      parent.removeChild(codeEl);
      if (first) {
        var r = document.createRange();
        r.setStartBefore(first); r.setEndAfter(last || first);
        sel.removeAllRanges(); sel.addRange(r);
      }
      parent.normalize();
    } else {
      var range = sel.getRangeAt(0);
      var code = document.createElement('code');
      if (range.collapsed) {
        code.textContent = 'code';
        range.insertNode(code);
      } else {
        try { range.surroundContents(code); }
        catch (e) { code.appendChild(range.extractContents()); range.insertNode(code); }
      }
      var r2 = document.createRange();
      r2.selectNodeContents(code);
      sel.removeAllRanges(); sel.addRange(r2);
    }
  }

  /* ---------- Code block (syntax-highlighted, per-line numbered) ----------
     A code block is a top-level <pre class="rn-code" data-lang="…"><code>…</code></pre>.
     highlight.js paints the tokens; the left gutter numbers every text line; a floating
     bar (language picker + copy) is anchored to the block while the caret is inside. */
  var HLJS = window.hljs || null;
  // Curated language list for the dropdown: [value, label]. 'auto' = auto-detect.
  var CODE_LANGS = [
    ['auto', 'Auto-detect'], ['plaintext', 'Plain text'],
    ['javascript', 'JavaScript'], ['typescript', 'TypeScript'],
    ['xml', 'HTML / XML'], ['css', 'CSS'], ['scss', 'SCSS'], ['json', 'JSON'],
    ['python', 'Python'], ['java', 'Java'], ['c', 'C'], ['cpp', 'C++'],
    ['csharp', 'C#'], ['go', 'Go'], ['rust', 'Rust'], ['php', 'PHP'],
    ['ruby', 'Ruby'], ['swift', 'Swift'], ['kotlin', 'Kotlin'], ['sql', 'SQL'],
    ['bash', 'Bash / Shell'], ['yaml', 'YAML'], ['markdown', 'Markdown'],
    ['lua', 'Lua'], ['r', 'R'], ['perl', 'Perl'], ['objectivec', 'Objective-C'],
    ['ini', 'INI / TOML'], ['makefile', 'Makefile'], ['diff', 'Diff']
  ];
  if (codeLang) {
    codeLang.innerHTML = CODE_LANGS.map(function (l) {
      return '<option value="' + l[0] + '">' + l[1] + '</option>';
    }).join('');
  }
  var codeBarFor = null;   // the code block the floating bar currently controls

  function isCodeBlock(el) {
    return !!(el && el.nodeType === 1 && el.tagName === 'PRE' && el.classList.contains('rn-code'));
  }
  function codeBlockOf(node) {
    var n = node;
    while (n && n !== editor) { if (isCodeBlock(n)) return n; n = n.parentNode; }
    return null;
  }
  function codeTextEl(pre) { return pre.querySelector('code') || pre; }
  // The logical source of a code block = its rendered text minus the trailing sentinel
  // newline. We always render one extra "\n" so the caret can sit on an empty last line
  // (a <pre> can't hold a caret AFTER a bare trailing newline — the "textarea trick").
  function codeSrc(pre) { return codeTextEl(pre).textContent.replace(/\n$/, ''); }

  // Character offset from the start of `el` to a DOM point (container, offset).
  function offsetOfPoint(el, container, offset) {
    var r = document.createRange();
    r.selectNodeContents(el);
    try { r.setEnd(container, offset); } catch (e) { return 0; }
    return r.toString().length;
  }
  // Character offset of the caret within `el`, or null if the caret isn't inside.
  function caretOffsetIn(el) {
    var sel = window.getSelection();
    if (!sel.rangeCount) return null;
    var r = sel.getRangeAt(0);
    if (!el.contains(r.startContainer) && r.startContainer !== el) return null;
    return offsetOfPoint(el, r.startContainer, r.startOffset);
  }
  // Place the caret `offset` characters into `el` (walking its text nodes).
  function setCaretOffsetIn(el, offset) {
    var walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, null);
    var node, seen = 0, target = null, tOff = 0;
    while ((node = walker.nextNode())) {
      var len = node.nodeValue.length;
      if (seen + len >= offset) { target = node; tOff = offset - seen; break; }
      seen += len;
    }
    var r = document.createRange();
    if (target) r.setStart(target, tOff);
    else { r.selectNodeContents(el); r.collapse(false); }
    r.collapse(true);
    var sel = window.getSelection();
    sel.removeAllRanges(); sel.addRange(r);
  }
  // Which text line (0-based) the caret sits on inside a code element.
  function codeCaretLine(codeEl) {
    var off = caretOffsetIn(codeEl);
    if (off == null) return -1;
    return (codeEl.textContent.slice(0, off).match(/\n/g) || []).length;
  }

  // Render `logical` code text into a block with syntax highlighting, then append a
  // sentinel newline so the last line is always caret-reachable. Restores the caret to
  // `caretOffset` (a logical offset) when provided.
  function renderCode(pre, logical, caretOffset) {
    if (!pre) return;
    var code = codeTextEl(pre);
    var lang = pre.getAttribute('data-lang') || 'auto';
    var out = null;
    if (HLJS && logical) {
      try {
        if (lang && lang !== 'auto' && HLJS.getLanguage(lang)) {
          out = HLJS.highlight(logical, { language: lang, ignoreIllegals: true }).value;
        } else {
          out = HLJS.highlightAuto(logical).value;
        }
      } catch (e) { out = null; }
    }
    code.innerHTML = ((out == null) ? escapeHtml(logical) : out) + '\n';   // trailing sentinel
    if (caretOffset != null) setCaretOffsetIn(code, caretOffset);
  }
  // Re-tokenize a block from its current DOM text (preserving the caret if given).
  function highlightCode(pre, caretOffset) { renderCode(pre, codeSrc(pre), caretOffset); }
  function highlightAllCode() {
    Array.prototype.forEach.call(editor.querySelectorAll('pre.rn-code'), function (pre) {
      if (!pre.getAttribute('data-lang')) pre.setAttribute('data-lang', 'auto');
      highlightCode(pre, null);
    });
  }

  // Turn the current block(s) into a code block — or, if already in one, unwrap it.
  function insertCodeBlock() {
    editor.focus();
    var sel = window.getSelection();
    var existing = sel.rangeCount ? codeBlockOf(sel.anchorNode) : null;
    if (existing) { unwrapCodeBlock(existing); onChange(); return; }

    var blocks = selectedBlocks();
    var pre = document.createElement('pre');
    pre.className = 'rn-code';
    pre.setAttribute('data-lang', 'auto');
    var code = document.createElement('code');
    pre.appendChild(code);

    var logical = '';
    if (blocks.length) {
      logical = blocks.map(function (b) { return b.textContent; }).join('\n');
      editor.insertBefore(pre, blocks[0]);
      for (var i = 0; i < blocks.length; i++) if (blocks[i].parentNode === editor) editor.removeChild(blocks[i]);
    } else {
      editor.appendChild(pre);
    }
    ensureContent();
    renderCode(pre, logical, logical.length);
    onChange();
  }
  // Convert a code block back into one plain paragraph per line.
  function unwrapCodeBlock(pre) {
    var lines = codeSrc(pre).split('\n');
    var frag = document.createDocumentFragment();
    for (var i = 0; i < lines.length; i++) {
      var p = document.createElement('p');
      if (lines[i]) p.textContent = lines[i]; else p.appendChild(document.createElement('br'));
      frag.appendChild(p);
    }
    var firstP = frag.firstChild;
    pre.parentNode.replaceChild(frag, pre);
    if (firstP) {
      var r = document.createRange();
      r.selectNodeContents(firstP); r.collapse(true);
      var s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
    }
  }

  // Enter inside a code block: insert a newline; an Enter on a trailing blank line
  // exits the block into a new paragraph below (VS Code-style escape).
  function handleCodeEnter(cb) {
    var code = codeTextEl(cb);
    var sel = window.getSelection();
    if (!sel.rangeCount) return;
    var r = sel.getRangeAt(0);
    var startOff = offsetOfPoint(code, r.startContainer, r.startOffset);
    var endOff = offsetOfPoint(code, r.endContainer, r.endOffset);
    var text = codeSrc(cb);                    // logical text (no sentinel)
    startOff = Math.min(startOff, text.length);
    endOff = Math.min(endOff, text.length);
    // Enter on an already-blank last line exits the block into a new paragraph below.
    if (startOff === endOff && startOff === text.length && /\n$/.test(text)) {
      renderCode(cb, text.replace(/\n$/, ''), null);
      var p = document.createElement('p'); p.appendChild(document.createElement('br'));
      if (cb.nextSibling) cb.parentNode.insertBefore(p, cb.nextSibling); else cb.parentNode.appendChild(p);
      var rr = document.createRange(); rr.selectNodeContents(p); rr.collapse(true);
      sel.removeAllRanges(); sel.addRange(rr);
      if (codeSrc(cb) === '') cb.parentNode.removeChild(cb);
      onChange();
      return;
    }
    renderCode(cb, text.slice(0, startOff) + '\n' + text.slice(endOff), startOff + 1);
    onChange();
  }

  // Live re-highlight while typing (debounced, and never mid-IME-composition).
  var codeHiTimer = null, codeComposing = false;
  function scheduleCodeHighlight() {
    var sel = window.getSelection();
    if (!sel.rangeCount || !codeBlockOf(sel.anchorNode)) return;
    if (codeHiTimer) clearTimeout(codeHiTimer);
    codeHiTimer = setTimeout(function () {
      codeHiTimer = null;
      if (codeComposing) { scheduleCodeHighlight(); return; }
      var s = window.getSelection();
      var block = s.rangeCount ? codeBlockOf(s.anchorNode) : null;
      if (!block) return;
      highlightCode(block, caretOffsetIn(codeTextEl(block)));
      save();
    }, 180);
  }
  editor.addEventListener('input', scheduleCodeHighlight);
  editor.addEventListener('compositionstart', function () { codeComposing = true; });
  editor.addEventListener('compositionend', function () { codeComposing = false; scheduleCodeHighlight(); });

  // The language + copy bar is hidden by default and revealed when the pointer hovers a
  // code block (or the bar itself). It's absolutely positioned inside .editor-area, over
  // the block's top-right corner, and kept out of the saved note content.
  var codeBarHideTimer = null;
  var editorArea = editor.parentNode;   // .editor-area (positioned ancestor)
  function showCodeBar(cb) {
    if (!codeBar || !cb) return;
    if (codeBarHideTimer) { clearTimeout(codeBarHideTimer); codeBarHideTimer = null; }
    codeBarFor = cb;
    var lang = cb.getAttribute('data-lang') || 'auto';
    if (codeLang && codeLang.value !== lang) codeLang.value = lang;
    codeBar.classList.add('show');
    codeBar.setAttribute('aria-hidden', 'false');
    positionCodeBar(cb);
  }
  function hideCodeBar(now) {
    if (!codeBar) return;
    if (codeBarHideTimer) { clearTimeout(codeBarHideTimer); codeBarHideTimer = null; }
    var doHide = function () { codeBar.classList.remove('show'); codeBar.setAttribute('aria-hidden', 'true'); codeBarFor = null; };
    if (now) doHide(); else codeBarHideTimer = setTimeout(doHide, 140);
  }
  function positionCodeBar(cb) {
    if (!codeBar || !cb || !codeBar.classList.contains('show')) return;
    var r = cb.getBoundingClientRect();
    var area = editorArea.getBoundingClientRect();
    // Hide when the block is scrolled out of the visible editor area
    if (r.bottom < area.top + 4 || r.top > area.bottom - 4) { hideCodeBar(true); return; }
    var bw = codeBar.offsetWidth || 140;
    var top = Math.max(r.top - area.top, 0) + 6;
    var left = (r.right - area.left) - bw - 10;
    codeBar.style.top = top + 'px';
    codeBar.style.left = Math.max(6, left) + 'px';
  }
  if (editor) {
    editor.addEventListener('mouseover', function (e) {
      var cb = codeBlockOf(e.target);
      if (cb) showCodeBar(cb);
    });
    editor.addEventListener('mouseout', function (e) {
      // Leaving the editor for something other than the bar → schedule hide
      var to = e.relatedTarget;
      if (to && (codeBar.contains(to) || codeBlockOf(to))) return;
      hideCodeBar(false);
    });
  }
  if (codeBar) {
    codeBar.addEventListener('mouseenter', function () { if (codeBarHideTimer) { clearTimeout(codeBarHideTimer); codeBarHideTimer = null; } });
    codeBar.addEventListener('mouseleave', function () { hideCodeBar(false); });
  }
  function activeCodeBlock() {
    if (codeBarFor) return codeBarFor;
    var s = window.getSelection();
    return s.rangeCount ? codeBlockOf(s.anchorNode) : null;
  }
  if (codeLang) codeLang.addEventListener('change', function () {
    var cb = activeCodeBlock();
    if (!cb) return;
    cb.setAttribute('data-lang', this.value);
    var s = window.getSelection();
    var inside = s.rangeCount && codeBlockOf(s.anchorNode) === cb;
    highlightCode(cb, inside ? caretOffsetIn(codeTextEl(cb)) : null);
    onChange();
  });
  // Copy text reliably. The async Clipboard API is frequently blocked inside Standard
  // Notes' iframe/Electron sandbox (no permission / not focused), so copy via a temporary
  // <textarea> + execCommand('copy') first, and only fall back to the async API.
  function copyPlainText(text) {
    var ok = false;
    try {
      var ta = document.createElement('textarea');
      ta.value = text;
      ta.setAttribute('readonly', '');
      ta.style.cssText = 'position:fixed;top:0;left:0;width:1px;height:1px;padding:0;border:0;opacity:0;';
      document.body.appendChild(ta);
      var sel = window.getSelection();
      var prev = (sel && sel.rangeCount) ? sel.getRangeAt(0).cloneRange() : null;
      ta.focus(); ta.select();
      try { ta.setSelectionRange(0, ta.value.length); } catch (e) {}
      ok = document.execCommand('copy');
      document.body.removeChild(ta);
      if (prev) { sel.removeAllRanges(); sel.addRange(prev); }
    } catch (e) { ok = false; }
    if (!ok && navigator.clipboard && navigator.clipboard.writeText) {
      try { navigator.clipboard.writeText(text); ok = true; } catch (e2) {}
    }
    return ok;
  }
  // Keep the current selection/focus when pressing the copy button
  if (codeCopy) codeCopy.addEventListener('mousedown', function (ev) { ev.preventDefault(); });
  if (codeCopy) codeCopy.addEventListener('click', function () {
    var cb = activeCodeBlock();
    if (!cb) return;
    var ok = copyPlainText(codeSrc(cb));
    var lbl = codeCopy.querySelector('.code-copy-label');
    codeCopy.classList.toggle('copied', ok);
    if (lbl) lbl.textContent = ok ? 'Copied' : 'Failed';
    setTimeout(function () { codeCopy.classList.remove('copied'); if (lbl) lbl.textContent = 'Copy'; }, 1200);
  });

  // Position the fixed popup right below its button so the toolbar overflow can't clip it
  function positionPopup(pop, anchor) {
    var r = anchor.getBoundingClientRect();
    pop.style.left = Math.max(6, Math.min(r.left, window.innerWidth - pop.offsetWidth - 8)) + 'px';
    pop.style.top = (r.bottom + 4) + 'px';
  }

  function currentLinkEl() {
    var sel = window.getSelection();
    if (!sel.rangeCount) return null;
    return closestTag(sel.anchorNode, 'A') || closestTag(sel.focusNode, 'A');
  }
  function openLinkPop() {
    saveSel();
    var a = currentLinkEl();               // read while the selection is still valid
    // Open after the current click finishes bubbling; otherwise the document-level
    // "close popups" handler would immediately close it when triggered from the menu.
    setTimeout(function () {
      closeMenus(); closePopups();
      linkInput.value = a ? (a.getAttribute('href') || '') : '';
      linkRemove.style.display = a ? '' : 'none';
      linkPop.classList.add('open');
      positionPopup(linkPop, toolbar.querySelector('.tbtn[data-cmd="link"]'));
      linkInput.focus(); linkInput.select();
    }, 0);
  }
  function applyLink() {
    var url = (linkInput.value || '').trim();
    restoreSel();
    if (url) {
      var a = currentLinkEl();
      if (a) a.setAttribute('href', url);
      else if (window.getSelection().toString()) document.execCommand('createLink', false, url);
      else document.execCommand('insertHTML', false, '<a href="' + escapeHtml(url) + '">' + escapeHtml(url) + '</a>');
      onChange();
    }
    linkPop.classList.remove('open');
  }
  function removeLink() {
    restoreSel();
    var a = currentLinkEl();
    if (a) {
      var r = document.createRange(); r.selectNode(a);
      var s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
    }
    document.execCommand('unlink');
    linkPop.classList.remove('open');
    onChange();
  }

  /* ---------- About / Donate dialog ---------- */
  // Keep keyboard focus inside an open modal (Tab / Shift+Tab cycle its controls).
  function modalFocusables(modal) {
    return Array.prototype.slice.call(
      modal.querySelectorAll('a[href],button:not([disabled]),input,[tabindex]:not([tabindex="-1"])')
    ).filter(function (el) { return el.offsetParent !== null; });
  }
  function trapModalTab(modal) {
    modal.addEventListener('keydown', function (ev) {
      if (ev.key !== 'Tab' || !modal.classList.contains('open')) return;
      var f = modalFocusables(modal);
      if (!f.length) return;
      var first = f[0], last = f[f.length - 1];
      if (ev.shiftKey && document.activeElement === first) { ev.preventDefault(); last.focus(); }
      else if (!ev.shiftKey && document.activeElement === last) { ev.preventDefault(); first.focus(); }
    });
  }
  function openAbout() {
    closeMenus();
    aboutModal.classList.add('open');
    aboutModal.setAttribute('aria-hidden', 'false');
    setTimeout(function () { if (aboutClose) aboutClose.focus(); }, 0);
  }
  function closeAbout() {
    aboutModal.classList.remove('open');
    aboutModal.setAttribute('aria-hidden', 'true');
    editor.focus();
  }
  function openShortcuts() {
    closeMenus();
    scModal.classList.add('open');
    scModal.setAttribute('aria-hidden', 'false');
    setTimeout(function () { if (scClose) scClose.focus(); }, 0);
  }
  function closeShortcuts() {
    scModal.classList.remove('open');
    scModal.setAttribute('aria-hidden', 'true');
    editor.focus();
  }
  trapModalTab(aboutModal);
  if (scModal) trapModalTab(scModal);

  // Font size in px: use the <font size=7> trick, then swap to inline style (on the current selection)
  function applyFontSize(px) {
    editor.focus();
    document.execCommand('styleWithCSS', false, false);
    document.execCommand('fontSize', false, '7');
    var fonts = editor.querySelectorAll('font[size="7"]');
    for (var i = 0; i < fonts.length; i++) { fonts[i].removeAttribute('size'); fonts[i].style.fontSize = px + 'px'; }
    document.execCommand('styleWithCSS', false, true);
    onChange();
  }

  // Current font size (px) at the selection, for the increase/decrease shortcuts
  function currentFontSizePx() {
    var el = selEl();
    var px = el ? Math.round(parseFloat(window.getComputedStyle(el).fontSize)) : 0;
    return px || 13;
  }
  // Bump the font size up/down by `delta` px (Ctrl+Shift+. / Ctrl+Shift+,)
  function changeFontSize(delta) {
    editor.focus();
    var next = Math.max(6, Math.min(200, currentFontSizePx() + delta));
    applyFontSize(next);                 // onChange() -> updateToolbar() syncs the Size box
    lastAction = function () { changeFontSize(delta); };   // F4 repeats
  }

  // Indent blocks via padding (keep the block flush-left so the highlight covers the whole line)
  function setIndent(b, lvl) {
    if (lvl > 0) { b.setAttribute('data-indent', lvl); b.style.setProperty('--indent', (lvl * 24) + 'px'); }
    else { b.removeAttribute('data-indent'); b.style.removeProperty('--indent'); }
  }
  function indentBlocks(dir) {
    var blocks = selectedBlocks();
    if (!blocks.length) { var c = currentBlock(); if (c) blocks = [c]; }
    for (var i = 0; i < blocks.length; i++) {
      var lvl = Math.max(0, (parseInt(blocks[i].getAttribute('data-indent'), 10) || 0) + dir);
      setIndent(blocks[i], lvl);
    }
    onChange();
  }
  // Outdent one level: reduce a block's --indent padding, OR — when the line is indented
  // with literal whitespace — strip a leading tab / up to 4 spaces. Returns true if changed,
  // so Shift+Tab works whether the indent is padding-based or space/tab-based.
  function outdentBlock(block) {
    if (!block) return false;
    var lvl = parseInt(block.getAttribute('data-indent'), 10) || 0;
    if (lvl > 0) { setIndent(block, lvl - 1); return true; }
    var walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT, null);
    var tn = walker.nextNode();
    if (!tn) return false;
    var m = /^(\t| {1,4})/.exec(tn.nodeValue);
    if (!m) return false;
    var sel = window.getSelection();
    var caretOff = (sel.rangeCount && sel.anchorNode === tn) ? sel.anchorOffset : -1;
    tn.nodeValue = tn.nodeValue.slice(m[0].length);
    if (caretOff >= 0) {
      var r = document.createRange();
      r.setStart(tn, Math.max(0, Math.min(caretOff - m[0].length, tn.nodeValue.length)));
      r.collapse(true);
      sel.removeAllRanges(); sel.addRange(r);
    }
    return true;
  }
  function outdentSelection() {
    var blocks = selectedBlocks();
    if (!blocks.length) { var c = currentBlock(); if (c) blocks = [c]; }
    var changed = false;
    for (var i = 0; i < blocks.length; i++) if (outdentBlock(blocks[i])) changed = true;
    if (changed) onChange();
  }

  // Convert LEADING tab characters (at the very start of a line) into the editor's own
  // padding indent, so pasted/loaded text that was indented with literal tabs adopts the
  // clean, uniform indent step. Tabs in the MIDDLE of a line (column separators like
  // "Super⇥Mở tìm kiếm") are left untouched. Code blocks keep their tabs verbatim.
  function absorbLeadingTabsIn(block) {
    var walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT, null);
    var tn = walker.nextNode();
    if (!tn) return;
    var m = /^\t+/.exec(tn.nodeValue);
    if (!m) return;
    tn.nodeValue = tn.nodeValue.slice(m[0].length);
    var cur = parseInt(block.getAttribute('data-indent'), 10) || 0;
    setIndent(block, Math.min(cur + m[0].length, 20));
  }
  function absorbLeadingTabs() {
    var kids = editor.children;
    for (var i = 0; i < kids.length; i++) {
      var b = kids[i];
      if (b.nodeType !== 1) continue;
      if (b.tagName === 'UL' || b.tagName === 'OL') {
        var lis = b.querySelectorAll('li');
        for (var j = 0; j < lis.length; j++) absorbLeadingTabsIn(lis[j]);
      } else if (b.tagName !== 'PRE') {              // never touch code-block indentation
        absorbLeadingTabsIn(b);
      }
    }
  }

  function clearFormat() {
    document.execCommand('removeFormat');
    var blocks = selectedBlocks();
    if (!blocks.length) { var c = currentBlock(); if (c) blocks = [c]; }
    for (var i = 0; i < blocks.length; i++) setIndent(blocks[i], 0);
    document.execCommand('formatBlock', false, 'P');
    onChange();
  }

  /* ---------- Paragraph style (heading / normal / quote) ----------
     execCommand('formatBlock') inside a list wraps the WHOLE list in the heading and
     nests on every repeat (<h2><h1><h2><ul>…). So in a list we apply the block to each
     <li>'s own content, replacing any existing heading instead of nesting. */
  var LI_BLK_RE = /^(H1|H2|H3|H4|H5|H6|P|DIV|BLOCKQUOTE)$/;
  function selectedListItems() {
    var sel = window.getSelection(), cur = currentListItem();
    if (!sel.rangeCount) return cur ? [cur] : [];
    var range = sel.getRangeAt(0), out = [], all = editor.querySelectorAll('li');
    for (var i = 0; i < all.length; i++) {
      try { if (range.intersectsNode(all[i])) out.push(all[i]); } catch (e) {}
    }
    return out.length ? out : (cur ? [cur] : []);
  }
  function setListItemBlock(li, tag) {
    var wrapper = null;
    for (var i = 0; i < li.children.length; i++) {
      if (LI_BLK_RE.test(li.children[i].tagName)) { wrapper = li.children[i]; break; }
    }
    if (tag === 'P') {                                  // Normal text → plain inline
      if (wrapper) { while (wrapper.firstChild) li.insertBefore(wrapper.firstChild, wrapper); li.removeChild(wrapper); }
      return;
    }
    var h = document.createElement(tag);
    if (wrapper) {                                      // replace the existing heading (no nesting)
      while (wrapper.firstChild) h.appendChild(wrapper.firstChild);
      li.replaceChild(h, wrapper);
    } else {                                            // wrap the li's own text (keep nested lists out)
      var refList = null, kids = Array.prototype.slice.call(li.childNodes);
      for (var k = 0; k < kids.length; k++) {
        var nd = kids[k];
        if (nd.nodeType === 1 && (nd.tagName === 'UL' || nd.tagName === 'OL')) { refList = nd; break; }
        h.appendChild(nd);
      }
      li.insertBefore(h, refList);
    }
  }
  function applyBlockFormat(tag) {
    editor.focus();
    if (tag === 'CODE') { insertCodeBlock(); return; }
    // Leaving a code block for another paragraph style: unwrap it back to plain lines first.
    var selCb = window.getSelection();
    var cb = selCb.rangeCount ? codeBlockOf(selCb.anchorNode) : null;
    if (cb) unwrapCodeBlock(cb);
    if (!currentListItem()) {
      document.execCommand('formatBlock', false, tag);
      // execCommand can merge a multi-line selection into one <blockquote>/<p> joined by
      // <br>; split it back so each line stays its own numbered block (both apply & remove).
      withCaretPreserved(splitMultilineBlocks);
      return;
    }
    var lis = selectedListItems();
    for (var i = 0; i < lis.length; i++) setListItemBlock(lis[i], tag);
  }

  /* ---------- Format painter (copy character formatting, Word-style) ---------- */
  var fp = { fmt: null, active: false, sticky: false };
  // Capture formatting at the caret/selection (Word-style: paragraph style + character
  // style). Character props are compared against the SOURCE BLOCK — so a heading's own
  // bold/size aren't captured as inline; only formatting applied ON TOP of the block is.
  function captureFormat() {
    var el = selEl();
    if (!el) return null;
    var block = currentBlock();
    var cs = window.getComputedStyle(el);
    var bcs = window.getComputedStyle(block || editor);
    var td = cs.textDecorationLine || cs.textDecoration || '';
    var btd = bcs.textDecorationLine || bcs.textDecoration || '';
    var bg = effectiveBg(el);
    var bt = block ? block.tagName : '';
    var blockType = /^(H1|H2|H3|H4|H5|H6|BLOCKQUOTE|PRE)$/.test(bt) ? bt : (/^(P|DIV)$/.test(bt) ? 'P' : null);
    return {
      blockType: blockType,     // paragraph style copied too (H2 → H2), like Word
      color: cs.color !== bcs.color ? cs.color : null,
      background: (bg && bg !== 'transparent') ? bg : null,
      bold:   parseInt(cs.fontWeight, 10) >= 600 && parseInt(bcs.fontWeight, 10) < 600,
      italic: cs.fontStyle !== 'normal' && bcs.fontStyle === 'normal',
      underline: td.indexOf('underline') >= 0 && btd.indexOf('underline') < 0,
      strike:    td.indexOf('line-through') >= 0 && btd.indexOf('line-through') < 0,
      code: inCode(),
      fontFamily: primaryFont(cs.fontFamily) !== primaryFont(bcs.fontFamily) ? cs.fontFamily : null,
      fontSize: Math.round(parseFloat(cs.fontSize)) !== Math.round(parseFloat(bcs.fontSize)) ? Math.round(parseFloat(cs.fontSize)) : null,
      vAlign: (cs.verticalAlign === 'sub' || cs.verticalAlign === 'super') ? cs.verticalAlign : null
    };
  }
  function applyCapturedFormat(fmt) {
    if (!fmt) return;
    editor.focus();
    var sel = window.getSelection();
    if (!sel.rangeCount || sel.getRangeAt(0).collapsed || !editor.contains(sel.anchorNode)) return;
    if (fmt.blockType) document.execCommand('formatBlock', false, fmt.blockType);   // paragraph style
    document.execCommand('styleWithCSS', false, true);
    document.execCommand('removeFormat');                 // clean slate for character style
    if (fmt.bold)      document.execCommand('bold');
    if (fmt.italic)    document.execCommand('italic');
    // Real <u>/<strike> tags so a later Ctrl+U / Ctrl+Shift+X can toggle them back OFF.
    if (fmt.underline) execTag('underline');
    if (fmt.strike)    execTag('strikeThrough');
    if (fmt.color)      document.execCommand('foreColor', false, fmt.color);
    if (fmt.background) document.execCommand('hiliteColor', false, fmt.background);
    if (fmt.fontFamily) document.execCommand('fontName', false, fmt.fontFamily);
    if (fmt.fontSize)   applyFontSize(fmt.fontSize);
    if (fmt.code)       toggleCode();
    if (fmt.vAlign) execTag(fmt.vAlign === 'sub' ? 'subscript' : 'superscript');
    onChange();
  }
  function fpArm(sticky) {
    var fmt = captureFormat();
    if (!fmt) return;
    fp.fmt = fmt; fp.active = true; fp.sticky = !!sticky;
    editor.classList.add('fp-armed');
    if (fpBtn) fpBtn.classList.add('active');
  }
  function fpDisarm() {
    fp.active = false; fp.sticky = false; fp.fmt = null;
    editor.classList.remove('fp-armed');
    if (fpBtn) fpBtn.classList.remove('active');
  }
  // Applying happens when the user finishes selecting the target text.
  editor.addEventListener('mouseup', function () {
    if (!fp.active) return;
    setTimeout(function () {
      var sel = window.getSelection();
      if (sel.rangeCount && !sel.getRangeAt(0).collapsed && editor.contains(sel.anchorNode)) {
        applyCapturedFormat(fp.fmt);
        if (!fp.sticky) fpDisarm();
      }
    }, 0);
  });
  if (fpBtn) {
    fpBtn.addEventListener('mousedown', function (ev) { ev.preventDefault(); saveSel(); });
    fpBtn.addEventListener('click', function () { if (fp.active) fpDisarm(); else fpArm(false); });
    fpBtn.addEventListener('dblclick', function () { fpArm(true); });
  }

  function toggleWrap() {
    wrapOn = !wrapOn;
    wrapUserSet = true;   // stop auto-following the screen size once the user decides
    editor.classList.toggle('wrap', wrapOn);
    try { localStorage.setItem('richnote-wrap', wrapOn ? '1' : '0'); } catch (e) {}
    refresh();
  }

  /* ---------- View / edit HTML source ----------
     A togglable raw-HTML view. The document's HTML is pretty-printed (one block per line,
     nested/indented) so it reads cleanly instead of one giant minified line; edits made in
     the textarea are parsed back into the live document when the view is turned off. */
  var SRC_INLINE = { A:1, ABBR:1, B:1, BDI:1, BDO:1, BR:1, CITE:1, CODE:1, DATA:1, DEL:1,
    DFN:1, EM:1, FONT:1, I:1, IMG:1, INS:1, KBD:1, LABEL:1, MARK:1, Q:1, S:1, SAMP:1, SMALL:1,
    SPAN:1, STRIKE:1, STRONG:1, SUB:1, SUP:1, TIME:1, U:1, VAR:1, WBR:1 };
  var SRC_VOID = { AREA:1, BASE:1, BR:1, COL:1, EMBED:1, HR:1, IMG:1, INPUT:1, LINK:1,
    META:1, PARAM:1, SOURCE:1, TRACK:1, WBR:1 };
  var SRC_VERBATIM = { PRE:1 };   // never reflow the inside of a code block

  function srcOpenTag(el) {
    var s = '<' + el.tagName.toLowerCase(), attrs = el.attributes;
    for (var i = 0; i < attrs.length; i++) {
      s += ' ' + attrs[i].name + '="' +
        String(attrs[i].value).replace(/&/g, '&amp;').replace(/"/g, '&quot;') + '"';
    }
    return s + '>';
  }
  function srcHasBlockChild(el) {
    for (var i = 0; i < el.children.length; i++) {
      if (!SRC_INLINE[el.children[i].tagName]) return true;
    }
    return false;
  }
  // Pretty-print HTML: block elements each on their own indented line, inline runs kept
  // whole on the parent's line so words/formatting aren't split.
  function formatHtml(html) {
    var root = document.createElement('div');
    root.innerHTML = String(html || '');
    var out = [];
    function pad(d) { return new Array(d + 1).join('  '); }
    function walk(parent, depth) {
      var nodes = parent.childNodes;
      for (var i = 0; i < nodes.length; i++) {
        var n = nodes[i];
        if (n.nodeType === 3) {                                   // text
          if (!/\S/.test(n.nodeValue)) continue;                  // ignore layout whitespace
          out.push(pad(depth) + escapeHtml(n.nodeValue.replace(/\s+/g, ' ').trim()));
        } else if (n.nodeType === 8) {                            // comment
          out.push(pad(depth) + '<!--' + n.nodeValue + '-->');
        } else if (n.nodeType === 1) {                            // element
          var tag = n.tagName;
          if (SRC_VOID[tag]) { out.push(pad(depth) + srcOpenTag(n)); continue; }
          var close = '</' + tag.toLowerCase() + '>';
          if (SRC_VERBATIM[tag] || !srcHasBlockChild(n)) {        // keep inline content on one line
            out.push(pad(depth) + srcOpenTag(n) + n.innerHTML.trim() + close);
          } else {
            out.push(pad(depth) + srcOpenTag(n));
            walk(n, depth + 1);
            out.push(pad(depth) + close);
          }
        }
      }
    }
    walk(root, 0);
    return out.join('\n');
  }
  // Drop the newline+indent whitespace the pretty-printer added between BLOCK elements,
  // without touching whitespace that sits next to inline content (where it is meaningful)
  // or inside <pre>/<code>.
  function stripPrettyWhitespace(root) {
    var kill = [];
    (function walk(node) {
      for (var i = 0; i < node.childNodes.length; i++) {
        var n = node.childNodes[i];
        if (n.nodeType === 1) {
          if (n.tagName !== 'PRE' && n.tagName !== 'CODE') walk(n);
        } else if (n.nodeType === 3 && !/\S/.test(n.nodeValue)) {
          var prev = n.previousSibling, next = n.nextSibling;
          var prevOK = !prev || (prev.nodeType === 1 && !SRC_INLINE[prev.tagName]);
          var nextOK = !next || (next.nodeType === 1 && !SRC_INLINE[next.tagName]);
          if (prevOK && nextOK) kill.push(n);
        }
      }
    })(root);
    for (var k = 0; k < kill.length; k++) if (kill[k].parentNode) kill[k].parentNode.removeChild(kill[k]);
  }

  var sourceOn = false;
  function applySource() {
    var tmp = document.createElement('div');
    tmp.innerHTML = sourceView.value;
    stripPrettyWhitespace(tmp);
    editor.innerHTML = tmp.innerHTML;
    normalizeBlocks();
    liftLists();
    flattenNestedBlocks();
    splitMultilineBlocks();
    absorbLeadingTabs();
    styleTables();
    ensureContent();
    highlightAllCode();
    onChange();                     // records one undo entry for the whole source edit
  }
  function setSourceView(on) {
    sourceOn = on;
    document.body.classList.toggle('source-mode', on);
    setActive('source', on);
    var item = document.querySelector('.menu-item[data-cmd="source"] .menu-check');
    if (item) item.textContent = on ? '✓' : '';
    if (on) {
      sourceView.value = formatHtml(editor.innerHTML);
      editor.style.display = 'none';
      sourceView.style.display = 'block';
      sourceView.focus();
    } else {
      sourceView.style.display = 'none';
      editor.style.display = '';
      editor.focus();
    }
  }
  function toggleSource() {
    if (sourceOn) { applySource(); setSourceView(false); }
    else setSourceView(true);
  }

  /* ============================================================
     TABLES
     ============================================================ */
  var TABLE_MAX_ROWS = 20, TABLE_MAX_COLS = 10;
  // Ensure every table carries the rn-table class (so pasted/loaded tables get styled)
  function styleTables() {
    var tbls = editor.querySelectorAll('table:not(.rn-table)');
    for (var i = 0; i < tbls.length; i++) tbls[i].classList.add('rn-table');
  }
  function cellOf(node) { return closestTag(node, 'TD') || closestTag(node, 'TH'); }
  function currentCell() {
    var sel = window.getSelection();
    if (!sel.rangeCount || !editor.contains(sel.anchorNode)) return null;
    return cellOf(sel.anchorNode);
  }
  function tableOf(node) { return closestTag(node, 'TABLE'); }
  function placeCaretInCell(cell) {
    if (!cell) return;
    var r = document.createRange();
    r.selectNodeContents(cell);
    r.collapse(true);
    var s = window.getSelection();
    s.removeAllRanges(); s.addRange(r);
    editor.focus();
  }
  function isBlockEmpty(b) {
    return b && b.nodeType === 1 && !/\S/.test(b.textContent) && !b.querySelector('img,table,hr');
  }
  function makeCell(tag) {
    var c = document.createElement(tag);
    c.appendChild(document.createElement('br'));
    return c;
  }
  function insertTable(rows, cols) {
    rows = Math.max(1, Math.min(TABLE_MAX_ROWS, rows | 0));
    cols = Math.max(1, Math.min(TABLE_MAX_COLS, cols | 0));
    editor.focus();
    // Insert at the live caret; only fall back to the saved range if the selection
    // isn't inside the editor (so a stale saved range can't move the insert point).
    var isel = window.getSelection();
    if (!isel.rangeCount || !editor.contains(isel.anchorNode)) restoreSel();
    var table = document.createElement('table');
    table.className = 'rn-table';
    var tbody = document.createElement('tbody');
    for (var r = 0; r < rows; r++) {
      var tr = document.createElement('tr');
      for (var c = 0; c < cols; c++) tr.appendChild(makeCell(r === 0 ? 'th' : 'td'));
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);

    var block = currentBlock();
    if (block && isBlockEmpty(block)) {
      editor.replaceChild(table, block);
    } else if (block) {
      block.parentNode.insertBefore(table, block.nextSibling);
    } else {
      editor.appendChild(table);
    }
    // Always keep a paragraph after the table so the caret can leave it
    if (!table.nextElementSibling) {
      var p = document.createElement('p'); p.appendChild(document.createElement('br'));
      editor.insertBefore(p, table.nextSibling);
    }
    placeCaretInCell(table.querySelector('th,td'));
    ensureContent();
    onChange();
  }

  // Tab / Shift+Tab move between cells; Tab past the last cell adds a new row.
  function moveCell(cell, dir) {
    var table = tableOf(cell);
    if (!table) return;
    var cells = Array.prototype.slice.call(table.querySelectorAll('th,td'));
    var idx = cells.indexOf(cell);
    var next = cells[idx + dir];
    if (next) { placeCaretInCell(next); return; }
    if (dir > 0) {                       // past the last cell → append a row and enter it
      addRow(cell, 1);
      var rows = table.rows;
      placeCaretInCell(rows[rows.length - 1].cells[0]);
    } else {                             // before the first cell → move to the block above
      var prev = table.previousElementSibling;
      if (prev) { var rr = document.createRange(); rr.selectNodeContents(prev); rr.collapse(false);
        var s = window.getSelection(); s.removeAllRanges(); s.addRange(rr); }
    }
  }

  function colIndexOf(cell) {
    var tr = cell.parentNode, n = 0;
    for (var i = 0; i < tr.cells.length; i++) { if (tr.cells[i] === cell) return n; n++; }
    return 0;
  }
  function addRow(cell, dir) {
    var tr = cell.parentNode, table = tableOf(cell);
    var count = tr.cells.length;
    var newTr = document.createElement('tr');
    for (var i = 0; i < count; i++) newTr.appendChild(makeCell('td'));
    tr.parentNode.insertBefore(newTr, dir < 0 ? tr : tr.nextSibling);
    onChange();
    return newTr;
  }
  function addCol(cell, dir) {
    var table = tableOf(cell), at = colIndexOf(cell);
    var rows = table.rows;
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var ref = row.cells[at];
      var tag = (ref && ref.tagName === 'TH') ? 'th' : 'td';
      var nc = makeCell(tag);
      if (dir < 0) row.insertBefore(nc, ref || null);
      else row.insertBefore(nc, ref ? ref.nextSibling : null);
    }
    var cg = table.querySelector('colgroup');   // keep colgroup in sync for resizing
    if (cg) {
      var refCol = cg.children[at];
      var ncol = document.createElement('col');
      ncol.style.width = (refCol && refCol.style.width) || '80px';
      cg.insertBefore(ncol, dir < 0 ? (refCol || null) : (refCol ? refCol.nextSibling : null));
      syncTableWidth(table);
    }
    onChange();
  }
  function delRow(cell) {
    var tr = cell.parentNode, table = tableOf(cell);
    if (table.rows.length <= 1) { delTable(cell); return; }
    var focusTr = tr.nextElementSibling || tr.previousElementSibling;
    tr.parentNode.removeChild(tr);
    if (focusTr) placeCaretInCell(focusTr.cells[0]);
    onChange();
  }
  function delCol(cell) {
    var table = tableOf(cell), at = colIndexOf(cell);
    if (table.rows[0] && table.rows[0].cells.length <= 1) { delTable(cell); return; }
    var rows = table.rows, focus = null;
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i], c = row.cells[at];
      if (c) { if (!focus) focus = row.cells[at + 1] || row.cells[at - 1]; row.removeChild(c); }
    }
    var cg = table.querySelector('colgroup');
    if (cg && cg.children[at]) { cg.removeChild(cg.children[at]); syncTableWidth(table); }
    if (focus) placeCaretInCell(focus);
    onChange();
  }
  function delTable(cell) {
    var table = tableOf(cell);
    if (!table) return;
    var focus = table.nextElementSibling || table.previousElementSibling;
    if (!focus) { focus = document.createElement('p'); focus.appendChild(document.createElement('br')); editor.insertBefore(focus, table); }
    table.parentNode.removeChild(table);
    hideTableTool();
    var r = document.createRange(); r.selectNodeContents(focus); r.collapse(true);
    var s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
    editor.focus();
    ensureContent();
    onChange();
  }

  /* ---------- Floating table toolbar (shown while the caret is inside a table) ---------- */
  var tableTool = null;
  function buildTableTool() {
    tableTool = document.createElement('div');
    tableTool.className = 'rn-table-tool';
    tableTool.innerHTML =
      '<button type="button" class="rn-tt-btn" data-tact="rowAbove" title="Insert row above"><svg class="ico" viewBox="0 0 16 16"><path d="M2 9.5h12M2 12.5h12"/><path d="M8 2v4M6 4h4"/></svg></button>' +
      '<button type="button" class="rn-tt-btn" data-tact="rowBelow" title="Insert row below"><svg class="ico" viewBox="0 0 16 16"><path d="M2 3.5h12M2 6.5h12"/><path d="M8 10v4M6 12h4"/></svg></button>' +
      '<span class="rn-tt-sep"></span>' +
      '<button type="button" class="rn-tt-btn" data-tact="colLeft" title="Insert column left"><svg class="ico" viewBox="0 0 16 16"><path d="M9.5 2v12M12.5 2v12"/><path d="M2 8h4M4 6v4"/></svg></button>' +
      '<button type="button" class="rn-tt-btn" data-tact="colRight" title="Insert column right"><svg class="ico" viewBox="0 0 16 16"><path d="M3.5 2v12M6.5 2v12"/><path d="M10 8h4M12 6v4"/></svg></button>' +
      '<span class="rn-tt-sep"></span>' +
      '<button type="button" class="rn-tt-btn rn-tt-del" data-tact="delRow" title="Delete row"><svg class="ico" viewBox="0 0 16 16"><path d="M2 6h12M2 10h12"/><path d="M6 13.5h4"/></svg></button>' +
      '<button type="button" class="rn-tt-btn rn-tt-del" data-tact="delCol" title="Delete column"><svg class="ico" viewBox="0 0 16 16"><path d="M6 2v12M10 2v12"/><path d="M12.5 6v4"/></svg></button>' +
      '<button type="button" class="rn-tt-btn rn-tt-del" data-tact="delTable" title="Delete table"><svg class="ico" viewBox="0 0 16 16"><path d="M3 4.5h10"/><path d="M5.5 4.5V3h5v1.5"/><path d="M4.2 4.5 5 13.5h6l.8-9"/></svg></button>';
    (document.querySelector('.editor-area') || document.body).appendChild(tableTool);
    tableTool.addEventListener('mousedown', function (ev) { ev.preventDefault(); });   // keep the caret
    tableTool.addEventListener('click', function (ev) {
      var b = ev.target.closest('.rn-tt-btn'); if (!b) return;
      var cell = currentCell(); if (!cell) return;
      switch (b.dataset.tact) {
        case 'rowAbove': addRow(cell, -1); break;
        case 'rowBelow': addRow(cell, 1); break;
        case 'colLeft':  addCol(cell, -1); break;
        case 'colRight': addCol(cell, 1); break;
        case 'delRow':   delRow(cell); break;
        case 'delCol':   delCol(cell); break;
        case 'delTable': delTable(cell); break;
      }
    });
  }
  function hideTableTool() { if (tableTool) tableTool.classList.remove('open'); }
  function positionTableTool(table) {
    if (!tableTool) return;
    var t = table.getBoundingClientRect();
    var host = editor.getBoundingClientRect();
    var top = Math.max(host.top + 2, t.top - tableTool.offsetHeight - 4);
    var left = Math.min(t.left, window.innerWidth - tableTool.offsetWidth - 8);
    tableTool.style.top = top + 'px';
    tableTool.style.left = Math.max(6, left) + 'px';
  }
  function updateTableTool(cur) {
    var table = (cur && cur.tagName === 'TABLE') ? cur : null;
    if (!table) { hideTableTool(); return; }
    if (!tableTool) buildTableTool();
    tableTool.classList.add('open');
    positionTableTool(table);
  }

  /* ============================================================
     TABLE INTERACTION — column/row resize + multi-cell selection
     ============================================================ */
  var RZ_EDGE = 5, RZ_MIN_COL = 34, RZ_MIN_ROW = 22;
  var rz = null;                            // active resize drag
  var hoverEdge = null, hoverCell = null;   // resize hover target
  var cellSel = [], cellSelTable = null;    // selected cells (multi-cell selection)
  var dragAnchor = null, dragging = false;  // cell drag-selection state

  function colCount(table) {
    var n = 0, rows = table.rows;
    for (var i = 0; i < rows.length; i++) n = Math.max(n, rows[i].cells.length);
    return n;
  }
  function syncTableWidth(table) {
    var cg = table.querySelector('colgroup'); if (!cg) return;
    var sum = 0;
    for (var i = 0; i < cg.children.length; i++) sum += parseFloat(cg.children[i].style.width) || 0;
    if (sum) table.style.width = sum + 'px';
  }
  // Freeze current column widths into a <colgroup> and switch to fixed layout so a drag
  // gives predictable, WYSIWYG resizing (like Word). Done lazily on the first resize.
  function ensureColgroup(table) {
    var cg = table.querySelector('colgroup');
    if (cg) return cg;
    cg = document.createElement('colgroup');
    var first = table.rows[0], n = colCount(table);
    for (var i = 0; i < n; i++) {
      var col = document.createElement('col');
      var cell = first && first.cells[i];
      col.style.width = (cell ? Math.round(cell.getBoundingClientRect().width) : 80) + 'px';
      cg.appendChild(col);
    }
    table.insertBefore(cg, table.firstChild);
    table.style.tableLayout = 'fixed';
    syncTableWidth(table);
    return cg;
  }

  /* ---------- Resize: hover cursor + drag ---------- */
  editor.addEventListener('mousemove', function (ev) {
    if (rz || dragging) return;
    var cell = ev.target.closest ? ev.target.closest('td,th') : null;
    hoverEdge = null; hoverCell = null;
    if (cell && editor.contains(cell)) {
      var r = cell.getBoundingClientRect();
      if (Math.abs(ev.clientX - r.right) <= RZ_EDGE) { hoverEdge = 'col'; hoverCell = cell; }
      else if (Math.abs(ev.clientY - r.bottom) <= RZ_EDGE) { hoverEdge = 'row'; hoverCell = cell; }
    }
    editor.style.cursor = hoverEdge === 'col' ? 'col-resize' : hoverEdge === 'row' ? 'row-resize' : '';
  });

  function startResize(ev, mode, cell) {
    var table = tableOf(cell); if (!table) return;
    if (mode === 'col') {
      var cg = ensureColgroup(table);
      var col = cg.children[colIndexOf(cell)];
      rz = { mode: 'col', table: table, col: col, startX: ev.clientX,
             startW: parseFloat(col.style.width) || cell.getBoundingClientRect().width };
    } else {
      var tr = cell.parentNode;
      rz = { mode: 'row', tr: tr, startY: ev.clientY, startH: tr.getBoundingClientRect().height };
    }
    document.body.style.userSelect = 'none';
    window.addEventListener('mousemove', onResizeMove, true);
    window.addEventListener('mouseup', onResizeUp, true);
  }
  function onResizeMove(ev) {
    if (!rz) return;
    if (rz.mode === 'col') {
      rz.col.style.width = Math.max(RZ_MIN_COL, rz.startW + (ev.clientX - rz.startX)) + 'px';
      syncTableWidth(rz.table);
    } else {
      rz.tr.style.height = Math.max(RZ_MIN_ROW, rz.startH + (ev.clientY - rz.startY)) + 'px';
    }
  }
  function onResizeUp() {
    window.removeEventListener('mousemove', onResizeMove, true);
    window.removeEventListener('mouseup', onResizeUp, true);
    document.body.style.userSelect = '';
    rz = null;
    onChange();
  }

  /* ---------- Multi-cell selection (drag a rectangle; covers row / column / any range) ---------- */
  function cellRC(cell) {
    var table = tableOf(cell), tr = cell.parentNode;
    return { r: Array.prototype.indexOf.call(table.rows, tr), c: colIndexOf(cell) };
  }
  function clearCellSel() {
    for (var i = 0; i < cellSel.length; i++) cellSel[i].classList.remove('rn-cell-sel');
    cellSel = []; cellSelTable = null;
  }
  function selectCellRange(a, b) {
    var table = tableOf(a);
    if (!table || tableOf(b) !== table) return;
    clearCellSel();
    var ra = cellRC(a), rb = cellRC(b);
    var r1 = Math.min(ra.r, rb.r), r2 = Math.max(ra.r, rb.r);
    var c1 = Math.min(ra.c, rb.c), c2 = Math.max(ra.c, rb.c);
    var rows = table.rows;
    for (var r = r1; r <= r2; r++)
      for (var c = c1; c <= c2; c++) {
        var cell = rows[r] && rows[r].cells[c];
        if (cell) { cell.classList.add('rn-cell-sel'); cellSel.push(cell); }
      }
    cellSelTable = table;
  }
  function reselectCells(cells) {
    clearCellSel();
    for (var i = 0; i < cells.length; i++) { cells[i].classList.add('rn-cell-sel'); cellSel.push(cells[i]); }
    cellSelTable = cells.length ? tableOf(cells[0]) : null;
  }
  function hasCellSel() { return cellSel.length > 0; }

  // mousedown (capture): start a resize on an edge, else begin a cell drag-selection.
  editor.addEventListener('mousedown', function (ev) {
    if (ev.button !== 0) return;
    if (hoverEdge && hoverCell) { ev.preventDefault(); startResize(ev, hoverEdge, hoverCell); return; }
    clearCellSel();
    var cell = ev.target.closest ? ev.target.closest('td,th') : null;
    if (cell && editor.contains(cell)) {
      dragAnchor = cell; dragging = false;
      window.addEventListener('mousemove', onCellDragMove, true);
      window.addEventListener('mouseup', onCellDragUp, true);
    }
  }, true);

  function onCellDragMove(ev) {
    if (!dragAnchor) return;
    var el = document.elementFromPoint(ev.clientX, ev.clientY);
    var over = el && el.closest ? el.closest('td,th') : null;
    if (!over || !editor.contains(over) || tableOf(over) !== tableOf(dragAnchor)) return;
    if (over === dragAnchor && !dragging) return;   // still inside the anchor → normal text selection
    if (!dragging) { dragging = true; editor.style.userSelect = 'none'; }
    ev.preventDefault();
    window.getSelection().removeAllRanges();          // suppress the native cross-cell highlight
    selectCellRange(dragAnchor, over);
  }
  function onCellDragUp() {
    window.removeEventListener('mousemove', onCellDragMove, true);
    window.removeEventListener('mouseup', onCellDragUp, true);
    editor.style.userSelect = '';
    dragAnchor = null;
    if (dragging && cellSel.length) placeCaretInCell(cellSel[0]);   // give a caret so typing works
    setTimeout(function () { dragging = false; }, 0);
  }

  /* ---------- Apply formatting / colour to every selected cell ---------- */
  function forEachSelCell(fn) {
    var cells = cellSel.slice(), sel = window.getSelection();
    for (var i = 0; i < cells.length; i++) {
      var r = document.createRange(); r.selectNodeContents(cells[i]);
      sel.removeAllRanges(); sel.addRange(r);
      fn(cells[i]);
    }
    reselectCells(cells);
    onChange();
  }
  var CELL_CMDS = { bold: 1, italic: 1, underline: 1, strike: 1, sub: 1, super: 1, code: 1,
    left: 1, center: 1, right: 1, justify: 1, clear: 1 };
  var CELL_TOGGLE = { bold: 'bold', italic: 'italic', underline: 'underline',
    strike: 'strikeThrough', sub: 'subscript', super: 'superscript' };
  var CELL_ALIGN = { left: 'justifyLeft', center: 'justifyCenter', right: 'justifyRight', justify: 'justifyFull' };
  function selectCellContents(cell) {
    var r = document.createRange(); r.selectNodeContents(cell);
    var s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
  }
  function applyCmdToCells(cmd) {
    var cells = cellSel.slice();
    if (CELL_TOGGLE[cmd]) {
      // Uniform target: turn the format ON for all cells unless every cell already has it
      // (avoids toggling CSS-bold <th> headers the wrong way in a mixed selection).
      var qname = CELL_TOGGLE[cmd], allOn = true;
      for (var i = 0; i < cells.length; i++) { selectCellContents(cells[i]); if (!q(qname)) { allOn = false; break; } }
      var want = !allOn;
      for (var j = 0; j < cells.length; j++) {
        selectCellContents(cells[j]);
        if (q(qname) === want) continue;
        // underline/strike/sub/super must be real tags (<u>/<strike>/<sub>/<sup>) — a
        // styleWithCSS text-decoration span can't be detected to toggle back OFF.
        if (cmd === 'sub' || cmd === 'super' || cmd === 'underline' || cmd === 'strike') execTag(qname);
        else document.execCommand(qname);
      }
    } else {
      for (var k = 0; k < cells.length; k++) {
        selectCellContents(cells[k]);
        if (CELL_ALIGN[cmd]) document.execCommand(CELL_ALIGN[cmd]);
        else if (cmd === 'code') toggleCode();
        else if (cmd === 'clear') document.execCommand('removeFormat');
      }
    }
    reselectCells(cells);
    onChange();
  }
  function applyColorToCells(kind, color) {
    forEachSelCell(function () { doColor(kind, color); });
  }
  // Fill the whole cell background (Word-style shading), not just the text highlight.
  function fillCells(cells, color) {
    for (var i = 0; i < cells.length; i++) {
      if (color === 'transparent' || color === 'none') cells[i].style.removeProperty('background-color');
      else cells[i].style.backgroundColor = color;
      if (!cells[i].getAttribute('style')) cells[i].removeAttribute('style');
    }
    if (backBar) backBar.style.background = (color === 'transparent' || color === 'none') ? 'transparent' : color;
    onChange();
  }
  // The target cells for a fill: the multi-cell selection, or the single cell at the caret.
  function fillTargetCells() {
    if (hasCellSel()) return cellSel.slice();
    var cc = currentCell();
    return cc ? [cc] : null;
  }

  /* ---------- Apply formatting across a discontiguous (Ctrl+drag) selection ---------- */
  function mcActive() { return !!(window.__richnoteMC && window.__richnoteMC.active()); }
  var MC_FMT = { bold: 1, italic: 1, underline: 1, strike: 1, sub: 1, super: 1, code: 1,
    left: 1, center: 1, right: 1, justify: 1, clear: 1 };
  // The raw command run once per selected range (the multi-cursor driver sets the native
  // selection to each range before calling this).
  function rawFormat(cmd) {
    switch (cmd) {
      case 'bold':      document.execCommand('bold'); break;
      case 'italic':    document.execCommand('italic'); break;
      // Real <u>/<strike>/<sub>/<sup> tags — a styleWithCSS span can't be toggled back OFF.
      case 'underline': execTag('underline'); break;
      case 'strike':    execTag('strikeThrough'); break;
      case 'sub':       execTag('subscript'); break;
      case 'super':     execTag('superscript'); break;
      case 'code':      toggleCode(); break;
      case 'left':      document.execCommand('justifyLeft'); break;
      case 'center':    document.execCommand('justifyCenter'); break;
      case 'right':     document.execCommand('justifyRight'); break;
      case 'justify':   document.execCommand('justifyFull'); break;
      case 'clear':     document.execCommand('removeFormat'); break;
    }
  }
  function rawColor(kind, color) {
    document.execCommand('styleWithCSS', false, true);
    document.execCommand(kind === 'fore' ? 'foreColor' : 'hiliteColor', false, color);
  }

  /* ---------- Show / hide the formatting toolbar (default: shown) ---------- */
  var toolbarOn = true;
  try { toolbarOn = localStorage.getItem('richnote-toolbar') !== '0'; } catch (e) {}
  function applyToolbar() {
    var app = document.getElementById('app');
    if (app) app.classList.toggle('toolbar-collapsed', !toolbarOn);
    var b = document.getElementById('toolbarToggle');
    if (b) { b.classList.toggle('collapsed', !toolbarOn); b.title = toolbarOn ? 'Hide toolbar' : 'Show toolbar'; }
  }
  function toggleToolbar() {
    toolbarOn = !toolbarOn;
    try { localStorage.setItem('richnote-toolbar', toolbarOn ? '1' : '0'); } catch (e) {}
    applyToolbar();
    refresh();
  }

  // Help/dialog commands that must NOT pull focus back into the editor — doing so
  // would re-trigger the software keyboard on touch devices.
  var NO_FOCUS_CMDS = { about: 1, donate: 1, shortcuts: 1 };

  function exec(cmd) {
    if (NO_FOCUS_CMDS[cmd]) {
      switch (cmd) {
        case 'about':     case 'donate': openAbout(); return;
        case 'shortcuts': openShortcuts(); return;
      }
    }
    editor.focus();
    // On touch devices we blur the editor when opening a menu (to hide the
    // keyboard), which drops the live selection — restore it before formatting.
    var live = window.getSelection();
    if (savedRange && (!live.rangeCount || !editor.contains(live.anchorNode))) restoreSel();
    // With a multi-cell selection active, formatting applies to every selected cell
    if (hasCellSel() && CELL_CMDS[cmd]) { applyCmdToCells(cmd); return; }
    // With a discontiguous (Ctrl+drag) selection, formatting applies to every range
    if (MC_FMT[cmd] && mcActive()) { window.__richnoteMC.run(function () { rawFormat(cmd); }); return; }
    if (REPEATABLE[cmd]) lastAction = (function (c) { return function () { exec(c); }; })(cmd);
    switch (cmd) {
      case 'undo':      undo(); return;
      case 'redo':      redo(); return;
      case 'cut':       document.execCommand('cut'); break;
      case 'copy':      document.execCommand('copy'); break;
      case 'paste':       doPaste(); return;
      case 'pasteValue':  pasteValueOnly(); return;
      case 'selectAll': document.execCommand('selectAll'); break;
      case 'bold':      document.execCommand('bold'); break;
      case 'italic':    document.execCommand('italic'); break;
      // Real <u>/<strike>/<sub>/<sup> tags — a styleWithCSS span can't be toggled back OFF,
      // and for sub/super it also wouldn't shrink the text or survive a paste.
      case 'underline': execTag('underline'); break;
      case 'strike':    execTag('strikeThrough'); break;
      case 'sub':       execTag('subscript'); break;
      case 'super':     execTag('superscript'); break;
      // "Code" now produces a code block; inside a table cell fall back to inline code.
      case 'code':      if (currentCell()) { toggleCode(); break; } insertCodeBlock(); return;
      case 'h1':        applyBlockFormat('H1'); break;
      case 'h2':        applyBlockFormat('H2'); break;
      case 'h3':        applyBlockFormat('H3'); break;
      case 'h4':        applyBlockFormat('H4'); break;
      case 'h5':        applyBlockFormat('H5'); break;
      case 'h6':        applyBlockFormat('H6'); break;
      case 'p':         applyBlockFormat('P'); break;
      case 'quote':     applyBlockFormat('BLOCKQUOTE'); break;
      case 'codeblock': applyBlockFormat('CODE'); return;
      case 'ul':        document.execCommand('insertUnorderedList'); liftLists(); break;
      case 'ol':        document.execCommand('insertOrderedList');   liftLists(); break;
      case 'checklist': toggleChecklist(); return;
      case 'hr':        insertHR(); return;
      case 'left':      document.execCommand('justifyLeft'); break;
      case 'center':    document.execCommand('justifyCenter'); break;
      case 'right':     document.execCommand('justifyRight'); break;
      case 'justify':   document.execCommand('justifyFull'); break;
      case 'indent':    indentBlocks(1); return;
      case 'outdent':   indentBlocks(-1); return;
      case 'link':      openLinkPop(); return;
      case 'about':     case 'donate': openAbout(); return;
      case 'shortcuts': openShortcuts(); return;
      case 'clear':     clearFormat(); return;
      case 'formatPainter': fpArm(false); return;
      case 'wrap':      toggleWrap(); return;
      case 'source':    toggleSource(); return;
      case 'minimap':   if (window.__richnoteMinimap) window.__richnoteMinimap.toggle(); return;
      case 'find':        if (window.__richnoteFind) window.__richnoteFind.open(); return;
      case 'findReplace': if (window.__richnoteFind) window.__richnoteFind.openReplace(); return;
      case 'table':       insertTable(3, 3); return;
      case 'toggleToolbar': toggleToolbar(); return;
    }
    onChange();
  }

  /* ---------- Menu bar (Edit / View) ---------- */
  function closeMenus() {
    var open = menubar.querySelectorAll('.menu.open');
    for (var i = 0; i < open.length; i++) open[i].classList.remove('open');
  }
  menubar.addEventListener('mousedown', function (ev) {
    if (ev.target.closest('.menu-title, .menu-item, .tbtn')) ev.preventDefault(); // keep the editor selection
  });
  menubar.addEventListener('click', function (ev) {
    var tbtn = ev.target.closest('.tbtn');   // undo / redo living in the menu bar
    if (tbtn) { closeMenus(); exec(tbtn.dataset.cmd); return; }
    var title = ev.target.closest('.menu-title');
    if (title) {
      var menu = title.parentNode;
      var wasOpen = menu.classList.contains('open');
      closeMenus(); closePopups();
      if (!wasOpen) {
        menu.classList.add('open');
        // On touch devices the software keyboard would stay up (the editor keeps
        // focus) and cover the dropdown — save the selection and blur to hide it.
        if (mqTouch && mqTouch.matches) { saveSel(); editor.blur(); }
      }
      return;
    }
    var item = ev.target.closest('.menu-item');
    if (item) { exec(item.dataset.cmd); closeMenus(); }
  });
  // When a menu is open, hovering another menu title switches to it
  menubar.addEventListener('mouseover', function (ev) {
    if (!menubar.querySelector('.menu.open')) return;
    var title = ev.target.closest('.menu-title');
    if (title) { closeMenus(); title.parentNode.classList.add('open'); }
  });
  document.addEventListener('click', function (ev) {
    if (!ev.target.closest('#menubar')) closeMenus();
    if (!ev.target.closest('.tcolor-wrap') && !ev.target.closest('.link-wrap') && !ev.target.closest('.tsize-wrap') && !ev.target.closest('.tdrop') && !ev.target.closest('.ttable-wrap')) closePopups();
  });

  /* ---------- Toolbar events ---------- */
  toolbar.addEventListener('mousedown', function (ev) {
    if (ev.target.closest('.tbtn')) { ev.preventDefault(); return; }     // keep focus on the button
    if (ev.target.closest('.tcolor')) { ev.preventDefault(); saveSel(); return; } // color button: keep selection
    if (ev.target.closest('input[type="color"]')) return;               // let the native picker open
    if (ev.target.closest('.color-pop')) { ev.preventDefault(); return; } // swatch: keep selection
    if (ev.target.closest('.tsel')) saveSel();                          // font / size selects
  });
  toolbar.addEventListener('click', function (ev) {
    var btn = ev.target.closest('.tbtn');
    if (btn && btn.dataset.cmd) exec(btn.dataset.cmd);
  });

  /* ---------- Toolbar dropdowns (align / lists / indent / more-format) ---------- */
  function wireDropdown(btn, pop) {
    if (!btn || !pop) return;
    btn.addEventListener('mousedown', function (ev) { ev.preventDefault(); saveSel(); });
    btn.addEventListener('click', function () {
      var open = pop.classList.contains('open');
      closePopups(); closeMenus();
      if (!open) { pop.classList.add('open'); positionPopup(pop, btn); }
    });
    pop.addEventListener('mousedown', function (ev) { ev.preventDefault(); });   // keep the selection
    pop.addEventListener('click', function (ev) {
      var opt = ev.target.closest('.tdrop-opt');
      if (!opt) return;
      restoreSel();
      exec(opt.dataset.cmd);
      pop.classList.remove('open');
    });
  }
  wireDropdown(alignBtn, alignPop);
  wireDropdown(listBtn, listPop);
  wireDropdown(indentBtn, indentPop);
  wireDropdown(fmtBtn, fmtPop);

  /* ---------- Table insert picker (hover a grid to choose size) ---------- */
  (function initTablePicker() {
    if (!tableBtn || !tablePop || !tableGrid) return;
    var GRID_R = 8, GRID_C = 8, hoverR = 0, hoverC = 0;
    var html = '';
    for (var r = 1; r <= GRID_R; r++)
      for (var c = 1; c <= GRID_C; c++)
        html += '<span class="ttable-cell" data-r="' + r + '" data-c="' + c + '"></span>';
    tableGrid.innerHTML = html;
    tableGrid.style.gridTemplateColumns = 'repeat(' + GRID_C + ', 1fr)';
    function mark(rr, cc) {
      hoverR = rr; hoverC = cc;
      var cells = tableGrid.children;
      for (var i = 0; i < cells.length; i++) {
        var cr = +cells[i].dataset.r, cc2 = +cells[i].dataset.c;
        cells[i].classList.toggle('on', cr <= rr && cc2 <= cc);
      }
      tableLabel.textContent = rr && cc ? (cc + ' × ' + rr) : 'Insert table';
    }
    tableBtn.addEventListener('mousedown', function (ev) { ev.preventDefault(); saveSel(); });
    tableBtn.addEventListener('click', function () {
      var open = tablePop.classList.contains('open');
      closePopups(); closeMenus();
      if (!open) { mark(0, 0); tablePop.classList.add('open'); positionPopup(tablePop, tableBtn); }
    });
    tableGrid.addEventListener('mousedown', function (ev) { ev.preventDefault(); });
    tableGrid.addEventListener('mouseover', function (ev) {
      var cell = ev.target.closest('.ttable-cell'); if (!cell) return;
      mark(+cell.dataset.r, +cell.dataset.c);
    });
    tableGrid.addEventListener('click', function (ev) {
      var cell = ev.target.closest('.ttable-cell'); if (!cell) return;
      tablePop.classList.remove('open');
      insertTable(+cell.dataset.r, +cell.dataset.c);
    });
  })();

  // Keep the floating table toolbar glued to its table as the editor scrolls
  editor.addEventListener('scroll', function () {
    if (tableTool && tableTool.classList.contains('open')) {
      var cur = currentBlock();
      if (cur && cur.tagName === 'TABLE') positionTableTool(cur); else hideTableTool();
    }
  }, { passive: true });
  window.addEventListener('resize', function () {
    if (tableTool && tableTool.classList.contains('open')) {
      var cur = currentBlock();
      if (cur && cur.tagName === 'TABLE') positionTableTool(cur);
    }
  });

  styleSelect.addEventListener('change', function () {
    restoreSel();
    var v = this.value;
    applyBlockFormat(v);
    lastAction = function () { applyBlockFormat(v); onChange(); };
    onChange();
  });
  fontSelect.addEventListener('change', function () {
    if (!this.value) return;
    restoreSel();
    var v = this.value;
    document.execCommand('styleWithCSS', false, true);
    document.execCommand('fontName', false, v);
    lastAction = function () { editor.focus(); document.execCommand('styleWithCSS', false, true); document.execCommand('fontName', false, v); onChange(); };
    onChange();
  });
  sizeSelect.addEventListener('change', function () {
    var px = parseInt(this.value, 10);
    if (!px || isNaN(px)) { this.value = currentFontSizePx(); return; }  // reject non-numeric input
    px = Math.max(6, Math.min(200, px));
    this.value = px;
    restoreSel();
    applyFontSize(px);
    lastAction = function () { applyFontSize(px); };
  });
  // Enter applies the typed size (blur -> change) and returns focus to the editor
  sizeSelect.addEventListener('keydown', function (ev) {
    if (ev.key === 'Enter') { ev.preventDefault(); this.blur(); }
  });

  // Custom size dropdown (native <datalist>/<select> popups mis-position inside the
  // Standard Notes iframe, so we render and position our own — fixed, like the color popups).
  (function initSizeDropdown() {
    if (!sizeArrow || !sizePop) return;
    var SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 40, 48, 64, 72];
    sizePop.innerHTML = SIZES.map(function (s) {
      return '<button type="button" class="tsize-opt" data-s="' + s + '">' + s + '</button>';
    }).join('');
    sizeArrow.addEventListener('mousedown', function (ev) { ev.preventDefault(); saveSel(); });
    sizeArrow.addEventListener('click', function () {
      var open = sizePop.classList.contains('open');
      closePopups(); closeMenus();
      if (!open) { sizePop.classList.add('open'); positionPopup(sizePop, sizeSelect.parentNode); }
    });
    sizePop.addEventListener('mousedown', function (ev) { ev.preventDefault(); });   // keep the selection
    sizePop.addEventListener('click', function (ev) {
      var opt = ev.target.closest('.tsize-opt');
      if (!opt) return;
      var px = parseInt(opt.dataset.s, 10);
      sizeSelect.value = px;
      restoreSel();
      applyFontSize(px);
      lastAction = function () { applyFontSize(px); };
      sizePop.classList.remove('open');
    });
  })();

  /* ---------- Color palette (Google Sheets style) ---------- */
  function buildGrid(el) {
    var html = '';
    for (var r = 0; r < PALETTE.length; r++) {
      for (var c = 0; c < PALETTE[r].length; c++) {
        var hex = PALETTE[r][c];
        html += '<button type="button" class="color-swatch" data-c="' + hex +
                '" title="' + hex + '" style="background:' + hex + '"></button>';
      }
    }
    el.innerHTML = html;
  }
  function closePopups() {
    var ps = toolbar.querySelectorAll('.color-pop.open');
    for (var i = 0; i < ps.length; i++) ps[i].classList.remove('open');
    if (linkPop) linkPop.classList.remove('open');
    if (sizePop) sizePop.classList.remove('open');
    if (tablePop) tablePop.classList.remove('open');
    var dds = toolbar.querySelectorAll('.tdrop-pop.open');
    for (var d = 0; d < dds.length; d++) dds[d].classList.remove('open');
  }
  function doColor(kind, color) {
    editor.focus();
    document.execCommand('styleWithCSS', false, true);
    document.execCommand(kind === 'fore' ? 'foreColor' : 'hiliteColor', false, color);
    var bar = kind === 'fore' ? foreBar : backBar;
    bar.style.background = color === 'transparent' ? 'transparent' : color;
    onChange();
  }
  function applyColor(kind, color) {
    // Discontiguous (Ctrl+drag) selection → colour every range
    if (mcActive()) { window.__richnoteMC.run(function () { rawColor(kind, color); }); closePopups(); return; }
    // Highlight colour inside a table = fill the whole cell (Word-style shading), unless
    // the user has an explicit text selection to highlight. Text colour still colours text.
    if (kind === 'back') {
      var sel = window.getSelection();
      var textSel = sel.rangeCount && !sel.getRangeAt(0).collapsed && !hasCellSel();
      if (!textSel) {
        var targets = fillTargetCells();
        if (targets) { fillCells(targets, color); closePopups(); return; }
      }
    }
    if (kind === 'fore' && hasCellSel()) { applyColorToCells('fore', color); closePopups(); return; }
    restoreSel();
    doColor(kind, color);
    lastAction = function () { doColor(kind, color); };
    closePopups();
  }
  // Highlight the palette swatch matching the caret's current colour
  function hexToRgb(hex) {
    var m = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex || '');
    return m ? 'rgb(' + parseInt(m[1], 16) + ', ' + parseInt(m[2], 16) + ', ' + parseInt(m[3], 16) + ')' : null;
  }
  function currentColorFor(kind) {
    var el = selEl();
    if (!el) return null;
    return kind === 'fore' ? window.getComputedStyle(el).color : effectiveBg(el);
  }
  function markSelectedSwatch(pop, kind) {
    var cur = currentColorFor(kind);
    var sws = pop.querySelectorAll('.color-swatch');
    for (var i = 0; i < sws.length; i++) {
      sws[i].classList.toggle('selected', !!cur && hexToRgb(sws[i].dataset.c) === cur);
    }
  }
  (function initColors() {
    var wraps = toolbar.querySelectorAll('.tcolor-wrap');
    for (var i = 0; i < wraps.length; i++) {
      (function (wrap) {
        var btn = wrap.querySelector('.tcolor');
        var kind = btn.dataset.color;
        var pop = wrap.querySelector('.color-pop');
        buildGrid(pop.querySelector('.color-grid'));
        btn.addEventListener('click', function () {
          var isOpen = pop.classList.contains('open');
          closePopups(); closeMenus();
          if (!isOpen) { pop.classList.add('open'); positionPopup(pop, btn); markSelectedSwatch(pop, kind); }
        });
        pop.addEventListener('mousedown', function (ev) {
          if (ev.target.closest('input[type="color"]')) return;
          ev.preventDefault();
        });
        pop.addEventListener('click', function (ev) {
          var sw = ev.target.closest('.color-swatch');
          if (sw) { applyColor(kind, sw.dataset.c); return; }
          if (ev.target.closest('.color-none')) applyColor(kind, 'transparent');
        });
        var input = pop.querySelector('input[type="color"]');
        if (input) input.addEventListener('change', function () { applyColor(kind, input.value); });
      })(wraps[i]);
    }
  })();

  /* ---------- Icons for the dropdown menu items ---------- */
  (function initMenuIcons() {
    function s(inner) { return '<svg class="ico" viewBox="0 0 16 16">' + inner + '</svg>'; }
    function g(txt, extra) { return '<span class="menu-ico-txt"' + (extra ? ' style="' + extra + '"' : '') + '>' + txt + '</span>'; }
    var ICONS = {
      undo: s('<path d="M4 7h6a3.5 3.5 0 1 1 0 7H6"/><path d="M6.5 4.5 4 7l2.5 2.5"/>'),
      redo: s('<path d="M12 7H6a3.5 3.5 0 1 0 0 7h4"/><path d="M9.5 4.5 12 7 9.5 9.5"/>'),
      cut: s('<circle cx="4" cy="4.2" r="2"/><circle cx="4" cy="11.8" r="2"/><path d="M5.7 5.4 14 11.8M5.7 10.6 14 4.2"/>'),
      copy: s('<rect x="5.5" y="5.5" width="8" height="8" rx="1.5"/><path d="M10.5 5.5V4A1.5 1.5 0 0 0 9 2.5H4A1.5 1.5 0 0 0 2.5 4v5A1.5 1.5 0 0 0 4 10.5h1.5"/>'),
      paste: s('<rect x="3.5" y="3" width="9" height="11" rx="1.5"/><rect x="5.5" y="1.9" width="5" height="2.5" rx="1"/>'),
      pasteValue: s('<rect x="3.5" y="3" width="9" height="11" rx="1.5"/><rect x="5.5" y="1.9" width="5" height="2.5" rx="1"/><path d="M5.7 7.6h4.6M5.7 10h3"/>'),
      selectAll: s('<rect x="2.6" y="2.6" width="10.8" height="10.8" rx="1.5" stroke-dasharray="2.3 1.7"/>'),
      find: s('<circle cx="7" cy="7" r="4.3"/><path d="M10.2 10.2 14 14"/>'),
      findReplace: s('<path d="M3 6.5A4.5 4.5 0 0 1 11 4l1.6 1.6"/><path d="M12.8 3v3h-3"/><path d="M13 9.5A4.5 4.5 0 0 1 5 12l-1.6-1.6"/><path d="M3.2 13v-3h3"/>'),
      bold: g('B', 'font-weight:800'),
      italic: g('I', 'font-style:italic'),
      underline: g('U', 'text-decoration:underline'),
      strike: g('S', 'text-decoration:line-through'),
      sub: g('X<sub>2</sub>'),
      super: g('X<sup>2</sup>'),
      h1: g('H1'), h2: g('H2'), h3: g('H3'), h4: g('H4'), h5: g('H5'), h6: g('H6'),
      p: g('¶'),
      quote: s('<path d="M3.3 4.2v7.6" stroke-width="2"/><path d="M6.3 5.4h7M6.3 8h7M6.3 10.6h5"/>'),
      codeblock: s('<rect x="1.8" y="3" width="12.4" height="10" rx="1.6"/><path d="M6.4 6.4 4.4 8l2 1.6M9.6 6.4 11.6 8l-2 1.6"/>'),
      left: s('<path d="M2 4h12M2 8h8M2 12h12"/>'),
      center: s('<path d="M2 4h12M4 8h8M2 12h12"/>'),
      right: s('<path d="M2 4h12M6 8h8M2 12h12"/>'),
      justify: s('<path d="M2 4h12M2 8h12M2 12h12"/>'),
      clear: s('<path d="M3 4.8V3.2h9v1.6"/><path d="M8 3.2 5.4 12.8"/><path d="M3.4 12.8h4"/><path d="M9.8 9.8 12.8 12.8"/><path d="M12.8 9.8 9.8 12.8"/>'),
      formatPainter: s('<path d="M13.6 2.4 7.9 8.1"/><path d="M8 8 4.1 10.8l2 3.5 5.2-2.4z"/><path d="M5.6 9.9 7.5 13.2M7.1 8.9 8.9 12.1M8.7 8.2 10.2 11"/>'),
      link: s('<path d="M6.7 9.3 9.3 6.7"/><path d="M8.4 4.6l1-1a2.7 2.7 0 0 1 3.9 3.9l-1 1"/><path d="M7.6 11.4l-1 1a2.7 2.7 0 0 1-3.9-3.9l1-1"/>'),
      ul: s('<circle class="dot" cx="2.6" cy="4" r="1.1"/><circle class="dot" cx="2.6" cy="8" r="1.1"/><circle class="dot" cx="2.6" cy="12" r="1.1"/><path d="M6 4h8M6 8h8M6 12h8"/>'),
      ol: s('<path d="M6 4h8M6 8h8M6 12h8"/><text x="0.4" y="5.6">1</text><text x="0.4" y="9.6">2</text><text x="0.4" y="13.6">3</text>'),
      checklist: s('<rect x="2" y="2.6" width="5" height="5" rx="1"/><path d="M3 5l1.2 1.2L6.4 4"/><path d="M9.5 4.6h5"/><path d="M2 11.4h5M9.5 11.4h5"/>'),
      table: s('<rect x="2" y="3" width="12" height="10" rx="1"/><path d="M2 6.5h12M2 9.7h12M6.4 3v10M10 3v10"/>'),
      hr: s('<path d="M2 8h12"/><path d="M3.4 4.4h9.2M3.4 11.6h9.2" opacity=".45"/>'),
      wrap: s('<path d="M2 4h12M2 8h8.4a2.4 2.4 0 0 1 0 4.8H7.4"/><path d="M9.2 10.8 7.4 12.8l1.8 2"/>'),
      minimap: s('<rect x="2.5" y="2.5" width="11" height="11" rx="1.5"/><path d="M10.2 3.5v9" opacity=".55"/><path d="M4.6 5.6h3.4M4.6 8h4.2M4.6 10.4h2.6" stroke-width="1"/>'),
      shortcuts: s('<rect x="1.5" y="4" width="13" height="8" rx="1.5"/><path d="M4 6.5h0M6 6.5h0M8 6.5h0M10 6.5h0M12 6.5h0M4 9h0M12 9h0M6 9.4h4" stroke-linecap="round"/>'),
      donate: s('<path class="fill" d="M8 13.4C8 13.4 2.6 10 2.6 6.3A2.6 2.6 0 0 1 8 5.1a2.6 2.6 0 0 1 5.4 1.2C13.4 10 8 13.4 8 13.4z"/>'),
      about: s('<circle cx="8" cy="8" r="6"/><path d="M8 7.3v3.4"/><circle class="dot" cx="8" cy="5.2" r=".75"/>')
    };
    var items = menubar.querySelectorAll('.menu-item[data-cmd]');
    for (var i = 0; i < items.length; i++) {
      var html = ICONS[items[i].dataset.cmd];
      if (!html) continue;
      var span = document.createElement('span');
      span.className = 'menu-ico';
      span.innerHTML = html;
      var check = items[i].querySelector('.menu-check');
      items[i].insertBefore(span, check ? check.nextSibling : items[i].firstChild);
    }
  })();

  /* ---------- About / Shortcuts dialog events ---------- */
  aboutClose.addEventListener('click', closeAbout);
  aboutModal.addEventListener('click', function (ev) { if (ev.target === aboutModal) closeAbout(); });
  if (scClose) scClose.addEventListener('click', closeShortcuts);
  if (scModal) scModal.addEventListener('click', function (ev) { if (ev.target === scModal) closeShortcuts(); });
  document.addEventListener('keydown', function (ev) {
    if (ev.key !== 'Escape') return;
    if (aboutModal.classList.contains('open')) closeAbout();
    if (scModal && scModal.classList.contains('open')) closeShortcuts();
    if (fp.active) fpDisarm();
  });

  /* ---------- Link popup events ---------- */
  linkApply.addEventListener('click', applyLink);
  linkRemove.addEventListener('click', removeLink);
  linkInput.addEventListener('keydown', function (ev) {
    if (ev.key === 'Enter') { ev.preventDefault(); applyLink(); }
    else if (ev.key === 'Escape') { ev.preventDefault(); linkPop.classList.remove('open'); editor.focus(); }
  });

  /* ---------- Hand cursor over links only while Ctrl/Cmd is held ---------- */
  function setModCursor(on) { editor.classList.toggle('mod-down', !!on); }
  editor.addEventListener('mousemove', function (ev) { setModCursor(ev.ctrlKey || ev.metaKey); });
  window.addEventListener('keydown', function (ev) { setModCursor(ev.ctrlKey || ev.metaKey); });
  window.addEventListener('keyup', function (ev) { setModCursor(ev.ctrlKey || ev.metaKey); });
  window.addEventListener('blur', function () { setModCursor(false); });

  /* ---------- Open link on Ctrl/Cmd + Click (a plain click just places the caret) ---------- */
  editor.addEventListener('click', function (ev) {
    if (!(ev.ctrlKey || ev.metaKey)) return;           // only open while Ctrl/Cmd is held
    var a = (ev.target && ev.target.closest) ? ev.target.closest('a') : null;
    if (!a || !editor.contains(a)) return;
    var href = a.getAttribute('href');
    if (!href) return;
    ev.preventDefault();
    var url = /^[a-z][\w+.-]*:|^\/\//i.test(href) ? href : 'https://' + href;  // add scheme if missing
    try { window.open(url, '_blank', 'noopener,noreferrer'); } catch (e) {}
  });

  /* ---------- Paste: keep formatting when available, else each line -> a block ---------- */
  /* ---------- Copy / Cut ----------
     The browser's default copy sometimes drops the block structure (multi-paragraph copies
     collapse into one line when pasted, and editor-only classes like current-line leak into
     the clipboard). We build the payload ourselves: clean HTML with every loose line wrapped
     in its own <p> (so lines never merge on paste), plus a block-structured text/plain
     fallback (one line per block) for plain-text targets. */
  function wrapLooseInline(box) {
    var run = null, nodes = Array.prototype.slice.call(box.childNodes);
    for (var i = 0; i < nodes.length; i++) {
      var n = nodes[i];
      var isBlock = n.nodeType === 1 && !SRC_INLINE[n.tagName];
      if (isBlock) { run = null; continue; }
      if (n.nodeType === 3 && !/\S/.test(n.nodeValue) && !run) { box.removeChild(n); continue; }
      if (!run) { run = document.createElement('p'); box.insertBefore(run, n); }
      run.appendChild(n);
    }
  }
  function buildCopyPayload(range) {
    var box = document.createElement('div');
    box.appendChild(range.cloneContents());
    var cl = box.querySelectorAll('.current-line');
    for (var i = 0; i < cl.length; i++) cl[i].classList.remove('current-line');
    // The editor's indent is a BLOCK attribute (data-indent/--indent), which the clipboard
    // drops. Bake it back into leading tab characters so the indentation travels with the
    // text (and any plain-text target keeps it); a paste back re-absorbs the tabs as indent.
    var ind = box.querySelectorAll('[data-indent]');
    for (var d = 0; d < ind.length; d++) {
      var lvl = parseInt(ind[d].getAttribute('data-indent'), 10) || 0;
      if (lvl > 0) ind[d].insertBefore(document.createTextNode(new Array(lvl + 1).join('\t')), ind[d].firstChild);
      ind[d].removeAttribute('data-indent');
      ind[d].style.removeProperty('--indent');
      if (!ind[d].getAttribute('style')) ind[d].removeAttribute('style');
    }
    wrapLooseInline(box);                       // loose inline runs → one <p> per line
    var lines = [];
    collectLines(box, lines);                   // one plain line per block
    var text = lines.length ? lines.join('\n') : box.textContent;
    return { html: box.innerHTML, text: text };
  }
  function onCopyEvent(ev, isCut) {
    if (hasCellSel()) return;                    // multi-cell selection: leave to default copy
    var sel = window.getSelection();
    if (!sel.rangeCount || sel.isCollapsed || !editor.contains(sel.anchorNode)) return;
    var cd = ev.clipboardData || window.clipboardData;
    if (!cd) return;
    var payload = buildCopyPayload(sel.getRangeAt(0));
    try {
      cd.setData('text/html', payload.html);
      cd.setData('text/plain', payload.text);
    } catch (e) { return; }                      // couldn't set → fall back to the browser default
    ev.preventDefault();
    if (isCut) { document.execCommand('delete'); onChange(); }
  }
  editor.addEventListener('copy', function (ev) { onCopyEvent(ev, false); });
  editor.addEventListener('cut',  function (ev) { onCopyEvent(ev, true); });

  editor.addEventListener('paste', function (ev) {
    if (pasteOverrideTimer) { clearTimeout(pasteOverrideTimer); pasteOverrideTimer = null; }
    var mode = pasteOverride; pasteOverride = null;   // 'plain' / 'format' / null (default)
    var cd = ev.clipboardData || window.clipboardData;
    if (!cd) return;
    var html = cd.getData('text/html');
    var text = cd.getData('text/plain');
    if ((html == null || html === '') && (text == null || text === '')) return;
    ev.preventDefault();
    if (mode === 'plain') {                                             // value only
      if (html && html.trim()) pasteValueFromHtml(html);               // one line per block
      else insertValueText(text);                                      // text/plain -> collapse blanks
    } else if (html && html.trim()) {
      pasteHtml(html);                                                  // default: keep formatting
    } else {
      insertPastedText(text);
    }
  });

  /* ---------- Keyboard shortcuts (Word-like) & Tab ---------- */
  editor.addEventListener('keydown', function (ev) {
    var mod = ev.ctrlKey || ev.metaKey;

    // Multi-cell selection: Esc clears it, Delete/Backspace empties the cells,
    // and typing a character collapses back to normal single-caret editing.
    if (hasCellSel()) {
      if (ev.key === 'Escape') { ev.preventDefault(); clearCellSel(); return; }
      if ((ev.key === 'Backspace' || ev.key === 'Delete') && cellSel.length > 1) {
        ev.preventDefault();
        for (var ci = 0; ci < cellSel.length; ci++) cellSel[ci].innerHTML = '<br>';
        placeCaretInCell(cellSel[0]);
        onChange();
        return;
      }
      if (ev.key.length === 1 && !mod && !ev.altKey) clearCellSel();
    }

    // F4: repeat the last formatting action (like Google Sheets)
    if (ev.key === 'F4') { ev.preventDefault(); if (lastAction) lastAction(); return; }

    // Inside a code block, Enter always inserts a newline (never a new paragraph); the
    // block manages its own text + syntax highlighting. Markdown auto-format is skipped.
    var sel0 = window.getSelection();
    var codeNow = sel0.rangeCount && editor.contains(sel0.anchorNode) ? codeBlockOf(sel0.anchorNode) : null;
    if (codeNow && ev.key === 'Enter') { ev.preventDefault(); handleCodeEnter(codeNow); return; }

    // Auto-format: a Markdown marker + Space at the line start → heading / list / quote / checklist
    if (ev.key === ' ' && !mod && !ev.altKey && !ev.shiftKey) {
      if (maybeMarkdownBlock()) { ev.preventDefault(); return; }
      linkifyBeforeCaret();               // a bare URL just typed → link it (the space still types)
      return;
    }
    // Enter on a "---" / "***" / "___" line → horizontal rule; else auto-link a URL just typed
    if (ev.key === 'Enter' && !ev.shiftKey && !mod && !ev.altKey) {
      if (maybeHorizontalRule()) { ev.preventDefault(); return; }
      linkifyBeforeCaret();
      return;
    }
    // Alt+Up / Alt+Down — move the current line(s) up / down (VS Code style)
    if (ev.altKey && !ev.shiftKey && !mod && (ev.key === 'ArrowUp' || ev.key === 'ArrowDown')) {
      if (mcActive()) return;
      ev.preventDefault();
      moveLine(ev.key === 'ArrowUp' ? -1 : 1);
      return;
    }

    // Shift+Enter -> create a NEW block (new line), not a soft line break
    if (ev.key === 'Enter' && ev.shiftKey && !mod) {
      ev.preventDefault();
      document.execCommand('insertParagraph');
      return;
    }

    // Tab / Shift+Tab: inside a table move between cells; otherwise indent / insert a tab
    if (ev.key === 'Tab') {
      ev.preventDefault();
      var tabCell = currentCell();
      if (tabCell) { moveCell(tabCell, ev.shiftKey ? -1 : 1); return; }
      // In a list: Tab nests the item, Shift+Tab un-nests; Shift+Tab on a top-level item
      // drops it out of the list into a normal paragraph (like Word).
      if (currentListItem()) {
        // Drop the current-line class first, else execCommand bakes its CSS background
        // into an inline span on the moved item.
        var clNow = editor.querySelector('.current-line');
        if (clNow) clNow.classList.remove('current-line');
        document.execCommand(ev.shiftKey ? 'outdent' : 'indent');
        liftLists();
        ensureContent();
        onChange();
        return;
      }
      // Indent by the SAME padding step whether one line or many are selected (a literal
      // '\t' rendered ~4 chars wide, so a single-line Tab used to jump further than a
      // multi-line one). A whole multi-line indent is a single undo step.
      if (ev.shiftKey) outdentSelection();               // outdent: padding OR leading spaces/tab
      else indentBlocks(1);
      return;
    }

    if (!mod) return;
    var c = ev.code, handled = true;
    if (ev.altKey && !ev.shiftKey) {                 // Ctrl+Alt+1..6 / 0
      if (c === 'Digit1') exec('h1');
      else if (c === 'Digit2') exec('h2');
      else if (c === 'Digit3') exec('h3');
      else if (c === 'Digit4') exec('h4');
      else if (c === 'Digit5') exec('h5');
      else if (c === 'Digit6') exec('h6');
      else if (c === 'Digit0') exec('p');
      else handled = false;
    } else if (ev.shiftKey) {
      if (c === 'KeyX') exec('strike');
      else if (c === 'Equal') exec('super');          // Ctrl+Shift+=  superscript
      else if (c === 'KeyV') {                                                    // Ctrl+Shift+V  paste value only
        // Arm 'plain', then try execCommand('paste') in this keydown gesture. In Electron/SN
        // it fires a trusted paste event (possibly async) that the handler applies as plain →
        // preventDefault so the browser doesn't also paste. In a normal browser execCommand is
        // disabled (returns false) → don't preventDefault so the native paste event fires.
        armPasteOverride('plain');
        var okPaste = false;
        try { okPaste = document.execCommand('paste'); } catch (e) { okPaste = false; }
        handled = !!okPaste;
      }
      else if (c === 'Digit7') exec('ol');
      else if (c === 'Digit8') exec('ul');
      else if (c === 'KeyP') exec('formatPainter');   // Ctrl+Shift+P  copy formatting
      else if (c === 'Period') changeFontSize(1);     // Ctrl+Shift+.  increase font size
      else if (c === 'Comma') changeFontSize(-1);     // Ctrl+Shift+,  decrease font size
      else if (c === 'KeyD') duplicateLine();         // Ctrl+Shift+D  duplicate line
      else if (c === 'KeyK') deleteLine();            // Ctrl+Shift+K  delete line
      else if (c === 'KeyZ') redo();                  // Ctrl+Shift+Z  redo
      else handled = false;
    } else {
      if (c === 'KeyL') exec('left');
      else if (c === 'KeyE') exec('center');
      else if (c === 'KeyR') exec('right');
      else if (c === 'KeyJ') exec('justify');
      else if (c === 'Equal') exec('sub');            // Ctrl+=  subscript
      else if (c === 'KeyK') exec('link');
      else if (c === 'KeyH') exec('findReplace');     // Ctrl+H  Find & Replace
      else if (c === 'Backslash') exec('clear');
      else if (c === 'KeyZ') undo();                  // Ctrl+Z  undo
      else if (c === 'KeyY') redo();                  // Ctrl+Y  redo
      // With a multi-cell or discontiguous (Ctrl+drag) selection, route Ctrl+B/I/U through
      // exec so the format hits every cell/range (else the browser only affects the caret).
      else if (c === 'KeyB' && (hasCellSel() || mcActive())) exec('bold');
      else if (c === 'KeyI' && (hasCellSel() || mcActive())) exec('italic');
      else if (c === 'KeyU' && (hasCellSel() || mcActive())) exec('underline');
      else handled = false;   // b/i/u/z/a/c/x/v: let the browser handle them
    }
    if (handled) ev.preventDefault();
  });

  /* ============================================================
     EMPTY-STATE PLACEHOLDER
     ============================================================ */
  function isEditorEmpty() {
    if (editor.querySelector('img,table,hr,li')) return false;
    if (editor.children.length > 1) return false;
    return !/\S/.test(editor.textContent);
  }
  function updateEmptyState() { editor.classList.toggle('is-empty', isEditorEmpty()); }

  /* ============================================================
     SAVE-STATE INDICATOR (status bar)
     ============================================================ */
  var stSave = document.getElementById('st-save');
  var saveStateTimer = null;
  function setSaveState(saving) {
    if (!stSave) return;
    stSave.textContent = saving ? 'Saving…' : 'Saved';
    stSave.classList.toggle('saving', !!saving);
    stSave.classList.toggle('saved', !saving);
  }

  /* ============================================================
     AUTO-LINKIFY  (paste + type-then-space)
     ============================================================ */
  var URL_RE = /\b(https?:\/\/[^\s<>()]+|www\.[^\s<>()]+)/gi;
  function trimTrailingPunct(u) {                 // don't swallow trailing . , ) ] etc.
    var m = /[.,;:!?)\]}"'»]+$/.exec(u);
    return m ? u.slice(0, u.length - m[0].length) : u;
  }
  function textHasUrl(text) { URL_RE.lastIndex = 0; return URL_RE.test(text); }
  // Escape a plain-text line to HTML, turning bare URLs into <a> links.
  function escapeAndLinkify(text) {
    var out = '', last = 0, m;
    URL_RE.lastIndex = 0;
    while ((m = URL_RE.exec(text))) {
      var raw = trimTrailingPunct(m[0]);
      var start = m.index;
      out += escapeHtml(text.slice(last, start));
      var href = /^www\./i.test(raw) ? 'https://' + raw : raw;
      out += '<a href="' + escapeHtml(href) + '">' + escapeHtml(raw) + '</a>';
      last = start + raw.length;
      URL_RE.lastIndex = last;
    }
    out += escapeHtml(text.slice(last));
    return out;
  }
  // Turn the URL token that ends right at the caret into a link (called on Space/Enter).
  function linkifyBeforeCaret() {
    var sel = window.getSelection();
    if (!sel.rangeCount) return false;
    var r = sel.getRangeAt(0);
    if (!r.collapsed) return false;
    var node = r.startContainer;
    if (node.nodeType !== 3 || closestTag(node, 'A')) return false;
    var offset = r.startOffset;
    var tok = /(\S+)$/.exec(node.nodeValue.slice(0, offset));
    if (!tok) return false;
    var token = tok[1];
    URL_RE.lastIndex = 0;
    var um = URL_RE.exec(token);
    if (!um || um.index !== 0) return false;
    var raw = trimTrailingPunct(token);
    if (raw.length < 6) return false;
    var startPos = offset - token.length;
    var wrap = document.createRange();
    wrap.setStart(node, startPos);
    wrap.setEnd(node, startPos + raw.length);
    sel.removeAllRanges(); sel.addRange(wrap);
    var href = /^www\./i.test(raw) ? 'https://' + raw : raw;
    document.execCommand('insertHTML', false, '<a href="' + escapeHtml(href) + '">' + escapeHtml(raw) + '</a>');
    return true;
  }

  /* ============================================================
     CARET HELPERS + LINE OPERATIONS (VS Code style)
     ============================================================ */
  function placeCaret(el, atEnd) {
    if (!el) return;
    var r = document.createRange();
    r.selectNodeContents(el); r.collapse(!atEnd);
    var s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
    editor.focus();
  }
  function placeCaretAtEnd(el) { placeCaret(el, true); }

  // The top-level blocks the operation targets (the selection's span, else the caret line).
  function opBlocks() {
    var blocks = selectedBlocks();
    if (!blocks.length) { var c = currentBlock(); if (c) blocks = [c]; }
    return blocks;
  }
  function duplicateLine() {
    var blocks = opBlocks();
    if (!blocks.length) return;
    var last = blocks[blocks.length - 1], ref = last.nextSibling, clones = [];
    for (var i = 0; i < blocks.length; i++) {
      var cl = blocks[i].cloneNode(true);
      if (cl.classList) cl.classList.remove('current-line');
      clones.push(cl);
    }
    for (var j = 0; j < clones.length; j++) last.parentNode.insertBefore(clones[j], ref);
    placeCaretAtEnd(clones[clones.length - 1]);
    onChange();
  }
  function moveLine(dir) {
    var blocks = opBlocks();
    if (!blocks.length) return;
    var first = blocks[0], last = blocks[blocks.length - 1];
    if (dir < 0) {
      var prev = first.previousElementSibling;
      if (!prev) return;
      last.parentNode.insertBefore(prev, last.nextSibling);   // slide the block above down past us
    } else {
      var next = last.nextElementSibling;
      if (!next) return;
      first.parentNode.insertBefore(next, first);             // slide the block below up past us
    }
    onChange();
  }
  function deleteLine() {
    var blocks = opBlocks();
    if (!blocks.length) return;
    var focus = blocks[blocks.length - 1].nextElementSibling || blocks[0].previousElementSibling;
    for (var i = 0; i < blocks.length; i++) blocks[i].parentNode.removeChild(blocks[i]);
    ensureContent();
    placeCaret(focus || editor.firstElementChild, false);
    onChange();
  }

  /* ============================================================
     HORIZONTAL RULE
     ============================================================ */
  function insertHR() {
    editor.focus();
    var block = currentBlock();
    var hr = document.createElement('hr');
    if (block && isBlockEmpty(block)) editor.replaceChild(hr, block);
    else if (block) block.parentNode.insertBefore(hr, block.nextSibling);
    else editor.appendChild(hr);
    var p = hr.nextElementSibling;
    if (!p || p.tagName === 'HR' || p.tagName === 'TABLE') {
      p = document.createElement('p'); p.appendChild(document.createElement('br'));
      hr.parentNode.insertBefore(p, hr.nextSibling);
    }
    placeCaret(p, false);
    ensureContent(); onChange();
  }
  // Enter on a line that is only "---" / "***" / "___" → a horizontal rule (Markdown-style).
  function maybeHorizontalRule() {
    var block = currentBlock();
    if (!block || currentListItem() || block.tagName === 'TABLE') return false;
    if (!/^(-{3,}|\*{3,}|_{3,})$/.test(block.textContent.trim())) return false;
    var hr = document.createElement('hr');
    editor.replaceChild(hr, block);
    var p = document.createElement('p'); p.appendChild(document.createElement('br'));
    hr.parentNode.insertBefore(p, hr.nextSibling);
    placeCaret(p, false);
    onChange();
    return true;
  }

  /* ============================================================
     CHECKLIST (task list)
     ============================================================ */
  function inChecklist() {
    var li = currentListItem();
    return !!(li && li.parentNode && li.parentNode.classList && li.parentNode.classList.contains('rn-checklist'));
  }
  function makeChecklist(checked) {
    var block = currentBlock();
    if (!block || block.tagName === 'TABLE' || block.tagName === 'HR') return;
    var ul = document.createElement('ul');
    ul.className = 'rn-checklist';
    var li = document.createElement('li');
    if (checked) li.setAttribute('data-checked', '');
    while (block.firstChild) li.appendChild(block.firstChild);
    if (!li.firstChild) li.appendChild(document.createElement('br'));
    ul.appendChild(li);
    block.parentNode.replaceChild(ul, block);
    placeCaretAtEnd(li);
    onChange();
  }
  function toggleChecklist() {
    editor.focus();
    if (inChecklist()) {                                     // checklist → back to paragraphs
      var ul = currentListItem().parentNode, made = [];
      Array.prototype.forEach.call(ul.children, function (item) {
        if (item.tagName !== 'LI') return;
        var p = document.createElement('p');
        while (item.firstChild) p.appendChild(item.firstChild);
        if (!p.firstChild) p.appendChild(document.createElement('br'));
        made.push(p);
      });
      for (var i = 0; i < made.length; i++) ul.parentNode.insertBefore(made[i], ul);
      ul.parentNode.removeChild(ul);
      placeCaretAtEnd(made[0] || null);
      ensureContent(); onChange();
      return;
    }
    makeChecklist(false);
  }
  // Click on a checklist item's checkbox (its left gutter) toggles the done state.
  editor.addEventListener('click', function (ev) {
    var li = ev.target.closest ? ev.target.closest('li') : null;
    if (!li || !editor.contains(li)) return;
    var ul = li.parentNode;
    if (!ul.classList || !ul.classList.contains('rn-checklist')) return;
    var r = li.getBoundingClientRect();
    if (ev.clientX - r.left <= 24) {                        // inside the checkbox gutter
      ev.preventDefault();
      li.toggleAttribute('data-checked');
      onChange();
    }
  });

  /* ============================================================
     MARKDOWN AUTO-FORMAT  (typed marker + Space at line start)
     ============================================================ */
  var MD_HEADING = /^#{1,6}$/;
  function maybeMarkdownBlock() {
    var sel = window.getSelection();
    if (!sel.rangeCount || !sel.getRangeAt(0).collapsed) return false;
    var block = currentBlock();
    if (!block || currentListItem()) return false;          // don't fire inside a list
    if (block.tagName === 'TABLE' || block.tagName === 'HR') return false;
    var r = sel.getRangeAt(0);
    var pre = document.createRange();
    pre.selectNodeContents(block);
    try { pre.setEnd(r.startContainer, r.startOffset); } catch (e) { return false; }
    var before = pre.toString();
    var apply = null;
    if (MD_HEADING.test(before)) { var lvl = before.length; apply = function () { applyBlockFormat('H' + lvl); }; }
    else if (before === '-' || before === '*') apply = function () { exec('ul'); };
    else if (before === '1.') apply = function () { exec('ol'); };
    else if (before === '>') apply = function () { applyBlockFormat('BLOCKQUOTE'); };
    else if (before === '[]' || before === '[ ]') apply = function () { makeChecklist(false); };
    else if (before === '[x]' || before === '[X]') apply = function () { makeChecklist(true); };
    else return false;
    var del = document.createRange();                       // remove the leading marker text
    del.setStart(block, 0);
    del.setEnd(r.startContainer, r.startOffset);
    del.deleteContents();
    block.normalize();
    var cr = document.createRange();
    cr.selectNodeContents(block); cr.collapse(true);
    sel.removeAllRanges(); sel.addRange(cr);
    apply();
    onChange();
    return true;
  }

  /* ---------- Standard Notes integration ---------- */
  function loadContent(html, title) {
    editor.innerHTML = html || '<p><br></p>';
    normalizeBlocks();
    liftLists();
    flattenNestedBlocks();              // repair any block-nested-in-block from older saves
    splitMultilineBlocks();             // legacy multi-line quotes → one numbered block per line
    absorbLeadingTabs();                // leading \t → editor indent (mid-line tabs untouched)
    styleTables();
    // Convert legacy indentation (margin-left) to the new padding-based --indent
    // so the current-line highlight reaches the left edge.
    Array.prototype.forEach.call(editor.children, function (b) {
      var ml = parseInt(b.style.marginLeft, 10);
      if (ml > 0) { b.style.marginLeft = ''; setIndent(b, Math.max(1, Math.round(ml / 24))); }
    });
    ensureContent();
    highlightAllCode();                 // re-tokenize any saved code blocks in the current theme
    lastSavedHTML = editor.innerHTML;
    if (title) stName.textContent = title;
    setSaveState(false);
    histReset();                        // a freshly loaded note starts a new undo history
    refresh();
  }
  function save() {
    if (!componentRelay || !workingNote) return;
    var html = editor.innerHTML;
    lastSavedHTML = html;
    setSaveState(true);
    if (saveStateTimer) clearTimeout(saveStateTimer);
    saveStateTimer = setTimeout(function () { setSaveState(false); }, 700);
    componentRelay.saveItemWithPresave(workingNote, function () {
      workingNote.content.text = html;
    });
  }
  function connectStandardNotes() {
    if (typeof ComponentRelay === 'undefined') return;
    componentRelay = new ComponentRelay({
      targetWindow: window,
      options: { coallesedSaving: true, coallesedSavingDelay: 250 },
      onReady: function () {
        var platform = componentRelay.platform;
        if (platform) document.body.classList.add('platform-' + platform);
      }
    });
    componentRelay.streamContextItem(function (note) {
      workingNote = note;
      var title = note.content && note.content.title;
      if (note.isMetadataUpdate) { if (title) stName.textContent = title; return; }
      var incoming = (note.content && note.content.text) || '';
      if (incoming === lastSavedHTML || incoming === editor.innerHTML) {
        if (title) stName.textContent = title;
        return;
      }
      loadContent(incoming, title);
    });
  }

  /* ---------- Startup ---------- */
  // Word Wrap: default from the device/size — ON for touch devices (phones/tablets) and
  // narrow windows, OFF on desktop — then honor a saved preference if one exists. The
  // default is computed BEFORE touching localStorage, because on iOS (Standard Notes runs
  // editors in a cross-origin iframe) localStorage access throws a SecurityError; if that
  // read were the only place we set wrapOn, wrap would silently stay off on iPhone/iPad.
  wrapOn = autoWrap();
  try {
    var savedWrap = localStorage.getItem('richnote-wrap');
    if (savedWrap !== null) { wrapUserSet = true; wrapOn = savedWrap === '1'; }
  } catch (e) {}
  editor.classList.toggle('wrap', wrapOn);
  // Auto-follow device / window-size changes until the user sets Word Wrap explicitly
  var onWrapMq = function () {
    if (wrapUserSet) return;
    wrapOn = autoWrap();
    editor.classList.toggle('wrap', wrapOn);
    refresh();
  };
  [mqTouch, mqNarrow].forEach(function (mq) {
    if (!mq) return;
    if (mq.addEventListener) mq.addEventListener('change', onWrapMq);
    else if (mq.addListener) mq.addListener(onWrapMq);   // older engines
  });
  try { document.execCommand('styleWithCSS', false, true); } catch (e) {}
  applyToolbar();
  ensureContent();
  histReset();                          // seed the undo history with the initial content
  connectStandardNotes();
  refresh();
})();
