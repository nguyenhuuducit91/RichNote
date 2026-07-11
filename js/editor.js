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
  var foreBar     = document.getElementById('foreBar');
  var backBar     = document.getElementById('backBar');
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
    bold: 1, italic: 1, underline: 1, strike: 1, code: 1,
    h1: 1, h2: 1, h3: 1, h4: 1, h5: 1, h6: 1, p: 1, quote: 1,
    left: 1, center: 1, right: 1, ul: 1, ol: 1, indent: 1, outdent: 1, clear: 1
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
  // Insert pasted text: multiple lines -> multiple blocks; a single line -> inline
  function insertPastedText(text) {
    if (text == null || text === '') return;
    editor.focus();
    var lines = text.replace(/\r\n?/g, '\n').split('\n');
    if (lines.length > 1 && lines[lines.length - 1] === '') lines.pop();
    if (lines.length === 1) {
      document.execCommand('insertText', false, lines[0]);
    } else {
      var html = lines.map(function (l) {
        return '<p>' + (l ? escapeHtml(l) : '<br>') + '</p>';
      }).join('');
      document.execCommand('insertHTML', false, html);
    }
    onChange();
  }
  function doPaste() {
    editor.focus();
    if (navigator.clipboard && navigator.clipboard.readText) {
      navigator.clipboard.readText()
        .then(function (t) { insertPastedText(t); })
        .catch(function () { try { document.execCommand('paste'); } catch (e) {} });
    } else {
      try { document.execCommand('paste'); } catch (e) {}
    }
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

  editor.addEventListener('scroll', function () { gutter.scrollTop = editor.scrollTop; });

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
  function updateToolbar() {
    setActive('bold', q('bold'));
    setActive('italic', q('italic'));
    setActive('underline', q('underline'));
    setActive('strike', q('strikeThrough'));
    setActive('ul', q('insertUnorderedList'));
    setActive('ol', q('insertOrderedList'));
    setActive('left', q('justifyLeft'));
    setActive('center', q('justifyCenter'));
    setActive('right', q('justifyRight'));
    var fb = '';
    try { fb = (document.queryCommandValue('formatBlock') || '').toLowerCase(); } catch (e) {}
    setActive('h1', fb === 'h1');
    setActive('h2', fb === 'h2');
    setActive('p', fb === 'p' || fb === 'div');
    setActive('quote', fb === 'blockquote');
    setActive('code', inCode());
    setActive('link', !!currentLinkEl());
    setActive('wrap', wrapOn);
    if (wrapItem) wrapItem.classList.toggle('checked', wrapOn);

    // Disable Undo/Redo when there is nothing to undo/redo
    try { setEnabled('undo', document.queryCommandEnabled('undo')); } catch (e) {}
    try { setEnabled('redo', document.queryCommandEnabled('redo')); } catch (e) {}

    // Sync the paragraph-style select with the current block
    var fbU = fb ? fb.toUpperCase() : '';
    styleSelect.value = /^(H1|H2|H3|H4|H5|H6|BLOCKQUOTE)$/.test(fbU) ? fbU : 'P';

    // Sync the font & size selects with the selection (Word-like)
    var el = selEl();
    if (el) {
      var cs = window.getComputedStyle(el);
      setSelectFont(cs.fontFamily);
      setSelectSize(Math.round(parseFloat(cs.fontSize)));
    }
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
    var opts = sizeSelect.options, idx = 0;
    for (var i = 1; i < opts.length; i++) { if (parseInt(opts[i].value, 10) === px) { idx = i; break; } }
    sizeSelect.selectedIndex = idx;
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

  function exec(cmd) {
    editor.focus();
    if (REPEATABLE[cmd]) lastAction = (function (c) { return function () { exec(c); }; })(cmd);
    switch (cmd) {
      case 'undo':      document.execCommand('undo'); break;
      case 'redo':      document.execCommand('redo'); break;
      case 'cut':       document.execCommand('cut'); break;
      case 'copy':      document.execCommand('copy'); break;
      case 'paste':     doPaste(); return;
      case 'selectAll': document.execCommand('selectAll'); break;
      case 'bold':      document.execCommand('bold'); break;
      case 'italic':    document.execCommand('italic'); break;
      case 'underline': document.execCommand('underline'); break;
      case 'strike':    document.execCommand('strikeThrough'); break;
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
      case 'indent':    indentBlocks(1); return;
      case 'outdent':   indentBlocks(-1); return;
      case 'link':      openLinkPop(); return;
      case 'about':     case 'donate': openAbout(); return;
      case 'clear':     clearFormat(); return;
      case 'wrap':      toggleWrap(); return;
    }
    onChange();
  }

  /* ---------- Menu bar (Edit / View) ---------- */
  function closeMenus() {
    var open = menubar.querySelectorAll('.menu.open');
    for (var i = 0; i < open.length; i++) open[i].classList.remove('open');
  }
  menubar.addEventListener('mousedown', function (ev) {
    if (ev.target.closest('.menu-title, .menu-item')) ev.preventDefault(); // keep the editor selection
  });
  menubar.addEventListener('click', function (ev) {
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
    if (!ev.target.closest('.tcolor-wrap') && !ev.target.closest('.link-wrap')) closePopups();
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
    if (btn) exec(btn.dataset.cmd);
  });

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
    if (!this.value) return;
    var px = parseInt(this.value, 10);
    restoreSel();
    applyFontSize(px);
    lastAction = function () { applyFontSize(px); };
  });

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
          if (!isOpen) { pop.classList.add('open'); positionPopup(pop, btn); }
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

  /* ---------- Paste: each line becomes its own block ---------- */
  editor.addEventListener('paste', function (ev) {
    var cd = ev.clipboardData || window.clipboardData;
    if (!cd) return;
    var text = cd.getData('text/plain');
    if (text == null || text === '') return;
    ev.preventDefault();
    insertPastedText(text);
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
      else if (c === 'Digit7') exec('ol');
      else if (c === 'Digit8') exec('ul');
      else if (c === 'KeyZ') { document.execCommand('redo'); onChange(); }
      else handled = false;
    } else {
      if (c === 'KeyL') exec('left');
      else if (c === 'KeyE') exec('center');
      else if (c === 'KeyR') exec('right');
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
  ensureContent();
  connectStandardNotes();
  refresh();
})();
