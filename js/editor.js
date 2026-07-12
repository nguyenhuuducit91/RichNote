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
  var linkPop     = document.getElementById('linkPop');
  var linkInput   = document.getElementById('linkInput');
  var linkApply   = document.getElementById('linkApply');
  var linkRemove  = document.getElementById('linkRemove');
  var aboutModal  = document.getElementById('aboutModal');
  var aboutClose  = document.getElementById('aboutClose');

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
    left: 1, center: 1, right: 1, justify: 1, ul: 1, ol: 1, indent: 1, outdent: 1, clear: 1
  };

  /* ---------- Utilities ---------- */
  function escapeHtml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  var BLOCK_RE = /^(P|DIV|H1|H2|H3|H4|H5|H6|BLOCKQUOTE|UL|OL|PRE|LI)$/;
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
  function ensureContent() {
    if (editor.children.length === 0) editor.innerHTML = '<p><br></p>';
  }
  // Insert an array of text lines: multiple -> a block each; a single -> inline
  function insertLines(lines) {
    if (!lines || !lines.length) return;
    if (lines.length > 1 && lines[lines.length - 1] === '') lines.pop();
    editor.focus();
    if (lines.length === 1) {
      document.execCommand('insertText', false, lines[0]);
    } else {
      var html = lines.map(function (l) {
        return '<p>' + (l ? escapeHtml(l) : '<br>') + '</p>';
      }).join('');
      document.execCommand('insertHTML', false, html);
      stripSpuriousBg();
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
  var PASTE_TAGS = { P:1, DIV:1, BR:1, SPAN:1, B:1, STRONG:1, I:1, EM:1, U:1, S:1, STRIKE:1, DEL:1,
    MARK:1, CODE:1, PRE:1, A:1, H1:1, H2:1, H3:1, H4:1, H5:1, H6:1, UL:1, OL:1, LI:1, BLOCKQUOTE:1, FONT:1, SUB:1, SUP:1 };
  var STYLE_KEEP = ['color', 'background-color', 'font-weight', 'font-style',
    'text-decoration', 'text-decoration-line', 'font-size', 'font-family'];

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
    ensureContent();
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
    var blocks = editor.children;
    var lines = [];
    var seen = {};
    for (var i = 0; i < blocks.length; i++) {
      var b = blocks[i];
      if (b.offsetHeight === 0) continue;                 // skip empty zero-height blocks
      var cs = window.getComputedStyle(b);
      var top = b.offsetTop + (parseFloat(cs.paddingTop) || 0) + (parseFloat(cs.marginTop) || 0);
      var lh = parseFloat(cs.lineHeight) || 26;
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
    gutter.scrollTop = editor.scrollTop;
  }

  /* ---------- Status bar ---------- */
  function updateStatus(cur) {
    var blocks = editor.children;
    var seen = {}, count = 0, ln = 1;
    for (var i = 0; i < blocks.length; i++) {
      var b = blocks[i];
      if (b.offsetHeight === 0) continue;
      var key = Math.round(b.offsetTop);
      if (seen[key] === undefined) seen[key] = ++count;
      if (b === cur) ln = seen[key];
    }
    var col = 1;
    var sel = window.getSelection();
    if (cur && sel.rangeCount) {
      var r = sel.getRangeAt(0);
      var pre = r.cloneRange();
      pre.selectNodeContents(cur);
      pre.setEnd(r.endContainer, r.endOffset);
      col = pre.toString().length + 1;
    }
    stPos.innerHTML = 'Ln : ' + ln + '&nbsp;&nbsp;Col : ' + col;
    stLen.innerHTML = 'length : ' + editor.textContent.length +
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
    requestAnimationFrame(function () { gutterSyncRaf = false; gutter.scrollTop = editor.scrollTop; });
  }, { passive: true });

  /* ---------- Toolbar button states ---------- */
  function q(cmd) { try { return document.queryCommandState(cmd); } catch (e) { return false; } }
  function inCode() {
    var sel = window.getSelection();
    if (!sel.rangeCount) return false;
    var n = sel.anchorNode;
    while (n && n !== editor) { if (n.nodeType === 1 && n.tagName === 'CODE') return true; n = n.parentNode; }
    return false;
  }
  function setActive(cmd, on) {
    var btn = toolbar.querySelector('.tbtn[data-cmd="' + cmd + '"]');
    if (btn) btn.classList.toggle('active', !!on);
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
    var ul = q('insertUnorderedList'), ol = q('insertOrderedList');
    if (listCur) listCur.innerHTML = ol ? LIST_ICONS.ol : LIST_ICONS.ul;   // reflect active list; bullet as neutral default
    listBtn.classList.toggle('active', ul || ol);
    setOpt(listPop, 'ul', ul);
    setOpt(listPop, 'ol', ol);
  }
  function updateFmtExtra() {
    if (!fmtBtn) return;
    var st = q('strikeThrough'), su = q('subscript'), sp = q('superscript'), cd = inCode();
    setOpt(fmtPop, 'strike', st);
    setOpt(fmtPop, 'sub', su);
    setOpt(fmtPop, 'super', sp);
    setOpt(fmtPop, 'code', cd);
    fmtBtn.classList.toggle('active', st || su || sp || cd);
  }

  function updateToolbar() {
    setActive('bold', q('bold'));
    setActive('italic', q('italic'));
    setActive('underline', q('underline'));
    updateFmtExtra();   // strike / sub / super / code (grouped dropdown)
    updateLists();      // bullet / numbered (grouped dropdown)
    updateAlign();      // left / center / right / justify (grouped dropdown)
    var fb = '';
    try { fb = (document.queryCommandValue('formatBlock') || '').toLowerCase(); } catch (e) {}
    setActive('h1', fb === 'h1');
    setActive('h2', fb === 'h2');
    setActive('p', fb === 'p' || fb === 'div');
    setActive('quote', fb === 'blockquote');
    setActive('link', !!currentLinkEl());
    setActive('wrap', wrapOn);
    if (wrapItem) wrapItem.classList.toggle('checked', wrapOn);

    // Disable Undo/Redo when there is nothing to undo/redo
    try { setEnabled('undo', document.queryCommandEnabled('undo')); } catch (e) {}
    try { setEnabled('redo', document.queryCommandEnabled('redo')); } catch (e) {}

    // Sync the paragraph-style select with the current block
    var fbU = fb ? fb.toUpperCase() : '';
    styleSelect.value = /^(H1|H2|H3|H4|H5|H6|BLOCKQUOTE)$/.test(fbU) ? fbU : 'P';

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
    var cur = currentBlock();
    var blocks = editor.children;
    for (var i = 0; i < blocks.length; i++) {
      blocks[i].classList.toggle('current-line', blocks[i] === cur);
    }
    renderGutter(cur);
    updateStatus(cur);
    updateToolbar();
  }

  function onChange() { refresh(); save(); }
  editor.addEventListener('input', onChange);

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
  function openAbout() {
    closeMenus();
    aboutModal.classList.add('open');
    aboutModal.setAttribute('aria-hidden', 'false');
  }
  function closeAbout() {
    aboutModal.classList.remove('open');
    aboutModal.setAttribute('aria-hidden', 'true');
  }

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
    return px || 12;
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

  function clearFormat() {
    document.execCommand('removeFormat');
    var blocks = selectedBlocks();
    if (!blocks.length) { var c = currentBlock(); if (c) blocks = [c]; }
    for (var i = 0; i < blocks.length; i++) setIndent(blocks[i], 0);
    document.execCommand('formatBlock', false, 'P');
    onChange();
  }

  function toggleWrap() {
    wrapOn = !wrapOn;
    wrapUserSet = true;   // stop auto-following the screen size once the user decides
    editor.classList.toggle('wrap', wrapOn);
    try { localStorage.setItem('richnote-wrap', wrapOn ? '1' : '0'); } catch (e) {}
    refresh();
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

  function exec(cmd) {
    editor.focus();
    if (REPEATABLE[cmd]) lastAction = (function (c) { return function () { exec(c); }; })(cmd);
    switch (cmd) {
      case 'undo':      document.execCommand('undo'); break;
      case 'redo':      document.execCommand('redo'); break;
      case 'cut':       document.execCommand('cut'); break;
      case 'copy':      document.execCommand('copy'); break;
      case 'paste':       doPaste(); return;
      case 'pasteValue':  pasteValueOnly(); return;
      case 'selectAll': document.execCommand('selectAll'); break;
      case 'bold':      document.execCommand('bold'); break;
      case 'italic':    document.execCommand('italic'); break;
      case 'underline': document.execCommand('underline'); break;
      case 'strike':    document.execCommand('strikeThrough'); break;
      // Emit real <sub>/<sup> (styleWithCSS would instead make a bare
      // vertical-align span that neither shrinks the text nor survives paste).
      case 'sub':       document.execCommand('styleWithCSS', false, false); document.execCommand('subscript');   document.execCommand('styleWithCSS', false, true); break;
      case 'super':     document.execCommand('styleWithCSS', false, false); document.execCommand('superscript'); document.execCommand('styleWithCSS', false, true); break;
      case 'code':      toggleCode(); break;
      case 'h1':        document.execCommand('formatBlock', false, 'H1'); break;
      case 'h2':        document.execCommand('formatBlock', false, 'H2'); break;
      case 'h3':        document.execCommand('formatBlock', false, 'H3'); break;
      case 'h4':        document.execCommand('formatBlock', false, 'H4'); break;
      case 'h5':        document.execCommand('formatBlock', false, 'H5'); break;
      case 'h6':        document.execCommand('formatBlock', false, 'H6'); break;
      case 'p':         document.execCommand('formatBlock', false, 'P'); break;
      case 'quote':     document.execCommand('formatBlock', false, 'BLOCKQUOTE'); break;
      case 'ul':        document.execCommand('insertUnorderedList'); break;
      case 'ol':        document.execCommand('insertOrderedList'); break;
      case 'left':      document.execCommand('justifyLeft'); break;
      case 'center':    document.execCommand('justifyCenter'); break;
      case 'right':     document.execCommand('justifyRight'); break;
      case 'justify':   document.execCommand('justifyFull'); break;
      case 'indent':    indentBlocks(1); return;
      case 'outdent':   indentBlocks(-1); return;
      case 'link':      openLinkPop(); return;
      case 'about':     case 'donate': openAbout(); return;
      case 'clear':     clearFormat(); return;
      case 'wrap':      toggleWrap(); return;
      case 'minimap':   if (window.__richnoteMinimap) window.__richnoteMinimap.toggle(); return;
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
      if (!wasOpen) menu.classList.add('open');
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
    if (!ev.target.closest('.tcolor-wrap') && !ev.target.closest('.link-wrap') && !ev.target.closest('.tsize-wrap') && !ev.target.closest('.tdrop')) closePopups();
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

  styleSelect.addEventListener('change', function () {
    restoreSel();
    var v = this.value;
    document.execCommand('formatBlock', false, v);
    lastAction = function () { editor.focus(); document.execCommand('formatBlock', false, v); onChange(); };
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
      bold: g('B', 'font-weight:800'),
      italic: g('I', 'font-style:italic'),
      underline: g('U', 'text-decoration:underline'),
      strike: g('S', 'text-decoration:line-through'),
      sub: g('X<sub>2</sub>'),
      super: g('X<sup>2</sup>'),
      h1: g('H1'), h2: g('H2'), h3: g('H3'), h4: g('H4'), h5: g('H5'), h6: g('H6'),
      p: g('¶'),
      quote: s('<path d="M3.3 4.2v7.6" stroke-width="2"/><path d="M6.3 5.4h7M6.3 8h7M6.3 10.6h5"/>'),
      left: s('<path d="M2 4h12M2 8h8M2 12h12"/>'),
      center: s('<path d="M2 4h12M4 8h8M2 12h12"/>'),
      right: s('<path d="M2 4h12M6 8h8M2 12h12"/>'),
      justify: s('<path d="M2 4h12M2 8h12M2 12h12"/>'),
      clear: s('<path d="M3 4.8V3.2h9v1.6"/><path d="M8 3.2 5.4 12.8"/><path d="M3.4 12.8h4"/><path d="M9.8 9.8 12.8 12.8"/><path d="M12.8 9.8 9.8 12.8"/>'),
      link: s('<path d="M6.7 9.3 9.3 6.7"/><path d="M8.4 4.6l1-1a2.7 2.7 0 0 1 3.9 3.9l-1 1"/><path d="M7.6 11.4l-1 1a2.7 2.7 0 0 1-3.9-3.9l1-1"/>'),
      ul: s('<circle class="dot" cx="2.6" cy="4" r="1.1"/><circle class="dot" cx="2.6" cy="8" r="1.1"/><circle class="dot" cx="2.6" cy="12" r="1.1"/><path d="M6 4h8M6 8h8M6 12h8"/>'),
      ol: s('<path d="M6 4h8M6 8h8M6 12h8"/><text x="0.4" y="5.6">1</text><text x="0.4" y="9.6">2</text><text x="0.4" y="13.6">3</text>'),
      wrap: s('<path d="M2 4h12M2 8h8.4a2.4 2.4 0 0 1 0 4.8H7.4"/><path d="M9.2 10.8 7.4 12.8l1.8 2"/>'),
      minimap: s('<rect x="2.5" y="2.5" width="11" height="11" rx="1.5"/><path d="M10.2 3.5v9" opacity=".55"/><path d="M4.6 5.6h3.4M4.6 8h4.2M4.6 10.4h2.6" stroke-width="1"/>'),
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

  /* ---------- About dialog events ---------- */
  aboutClose.addEventListener('click', closeAbout);
  aboutModal.addEventListener('click', function (ev) { if (ev.target === aboutModal) closeAbout(); });
  document.addEventListener('keydown', function (ev) {
    if (ev.key === 'Escape' && aboutModal.classList.contains('open')) closeAbout();
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

    // F4: repeat the last formatting action (like Google Sheets)
    if (ev.key === 'F4') { ev.preventDefault(); if (lastAction) lastAction(); return; }

    // Shift+Enter -> create a NEW block (new line), not a soft line break
    if (ev.key === 'Enter' && ev.shiftKey && !mod) {
      ev.preventDefault();
      document.execCommand('insertParagraph');
      return;
    }

    // Tab / Shift+Tab: indent (multi-line when several blocks are selected)
    if (ev.key === 'Tab') {
      ev.preventDefault();
      var blocks = selectedBlocks();
      if (ev.shiftKey || blocks.length > 1) indentBlocks(ev.shiftKey ? -1 : 1);
      else document.execCommand('insertText', false, '\t');
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
      else if (c === 'KeyC') exec('code');
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
      else if (c === 'Period') changeFontSize(1);     // Ctrl+Shift+.  increase font size
      else if (c === 'Comma') changeFontSize(-1);     // Ctrl+Shift+,  decrease font size
      else if (c === 'KeyZ') { document.execCommand('redo'); onChange(); }
      else handled = false;
    } else {
      if (c === 'KeyL') exec('left');
      else if (c === 'KeyE') exec('center');
      else if (c === 'KeyR') exec('right');
      else if (c === 'KeyJ') exec('justify');
      else if (c === 'Equal') exec('sub');            // Ctrl+=  subscript
      else if (c === 'KeyK') exec('link');
      else if (c === 'Backslash') exec('clear');
      else if (c === 'KeyY') { document.execCommand('redo'); onChange(); }
      else handled = false;   // b/i/u/z/a/c/x/v: let the browser handle them
    }
    if (handled) ev.preventDefault();
  });

  /* ---------- Standard Notes integration ---------- */
  function loadContent(html, title) {
    editor.innerHTML = html || '<p><br></p>';
    normalizeBlocks();
    // Convert legacy indentation (margin-left) to the new padding-based --indent
    // so the current-line highlight reaches the left edge.
    Array.prototype.forEach.call(editor.children, function (b) {
      var ml = parseInt(b.style.marginLeft, 10);
      if (ml > 0) { b.style.marginLeft = ''; setIndent(b, Math.max(1, Math.round(ml / 24))); }
    });
    ensureContent();
    lastSavedHTML = editor.innerHTML;
    if (title) stName.textContent = title;
    refresh();
  }
  function save() {
    if (!componentRelay || !workingNote) return;
    var html = editor.innerHTML;
    lastSavedHTML = html;
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
  // Word Wrap: honor a saved preference; otherwise follow the device/size — ON for
  // touch devices (phones/tablets) and narrow windows, OFF on desktop — and keep
  // following it as the window resizes until the user toggles it themselves.
  try {
    var savedWrap = localStorage.getItem('richnote-wrap');
    if (savedWrap === null) {
      wrapOn = autoWrap();
    } else {
      wrapUserSet = true;
      wrapOn = savedWrap === '1';
    }
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
  connectStandardNotes();
  refresh();
})();
