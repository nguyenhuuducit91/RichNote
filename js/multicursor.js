/*!
 * RichNote — Multi-cursor & Find (Sublime-Text style).
 *
 * Adds true multi-caret editing on top of the contenteditable editor:
 *   • Shift+Alt+Up / Down  — add a caret on the line above / below (column cursors)
 *   • Alt+Click            — add/toggle a caret at the click point
 *   • Ctrl+F               — Find bar; "Select all" (Alt+Enter) turns every match
 *                            into a caret so you edit them all at once
 *   • With many carets     — typing, Backspace/Delete, Enter and the arrow / Home /
 *                            End keys (with Shift to select) all act on every caret
 *   • Esc / plain click    — collapse back to a single caret
 *
 * Carets are kept as live DOM Ranges (the browser keeps their boundary points in
 * sync as we mutate the document), movement is delegated to Selection.modify() and
 * editing to execCommand(), so grapheme / bidi / block behaviour matches the native
 * editor. A non-editable overlay paints the extra carets and their selections.
 *
 * @author  Nguyen Huu Duc <nguyenhuuduc.it.91@gmail.com> (VIETIS)
 * @license MIT
 */
(function () {
  'use strict';

  var editor = document.getElementById('editor');
  var editorArea = editor && editor.closest ? editor.closest('.editor-area') : null;
  if (!editor) return;

  /* ---------- State ---------- */
  var carets = [];        // array of live Range objects (each may be a selection)
  var mcMode = false;     // true when >= 2 carets are active
  var goalX = null;       // remembered column (viewport x) for vertical caret adds
  var suppressInput = false;
  var composing = false;  // true while an IME composition (preedit) is in flight

  var sel = window.getSelection();

  /* ---------- Overlay ---------- */
  var overlay = document.createElement('div');
  overlay.className = 'mc-overlay';
  overlay.setAttribute('aria-hidden', 'true');
  document.body.appendChild(overlay);

  function clearOverlay() { overlay.textContent = ''; overlay.style.display = 'none'; }

  // Robust caret rectangle for a (possibly collapsed) range's focus point.
  // getBoundingClientRect() on a collapsed range is unreliable at text-node offset 0,
  // so fall back to probing an adjacent character, then to the containing element.
  function collapsedEndRect(range) {
    var r = range.cloneRange();
    r.collapse(false);
    var rect = r.getBoundingClientRect();
    if (rect && rect.height) return { left: rect.left, top: rect.top, height: rect.height };

    var node = r.startContainer, off = r.startOffset;
    if (node.nodeType === 3) {
      var len = node.nodeValue.length, probe = document.createRange();
      if (off < len) { probe.setStart(node, off); probe.setEnd(node, off + 1); var b = probe.getBoundingClientRect(); if (b.height) return { left: b.left, top: b.top, height: b.height }; }
      if (off > 0)   { probe.setStart(node, off - 1); probe.setEnd(node, off); var b2 = probe.getBoundingClientRect(); if (b2.height) return { left: b2.right, top: b2.top, height: b2.height }; }
    }
    var el = node.nodeType === 1 ? (node.childNodes[off] || node) : node.parentElement;
    if (el && el.nodeType === 1) {
      var er = el.getBoundingClientRect(), cs = getComputedStyle(el);
      return { left: er.left + (parseFloat(cs.paddingLeft) || 0), top: er.top, height: parseFloat(cs.lineHeight) || 20 };
    }
    return rect && rect.height ? rect : null;
  }

  // Paint an arbitrary list of ranges onto the overlay (used by render() and by the live
  // Ctrl+drag preview, which must show the existing selections while a new one is dragged).
  function paintRanges(ranges) {
    if (!ranges.length) { clearOverlay(); return; }
    var host = editor.getBoundingClientRect();
    overlay.style.display = 'block';
    overlay.style.left = host.left + 'px';
    overlay.style.top = host.top + 'px';
    overlay.style.width = host.width + 'px';
    overlay.style.height = host.height + 'px';
    var html = [];
    for (var i = 0; i < ranges.length; i++) {
      var range = ranges[i];
      if (!range.collapsed) {
        var rects = range.getClientRects();
        for (var k = 0; k < rects.length; k++) {
          var s = rects[k];
          if (!s.width && !s.height) continue;
          html.push('<div class="mc-sel" style="left:' + (s.left - host.left) + 'px;top:' +
            (s.top - host.top) + 'px;width:' + s.width + 'px;height:' + s.height + 'px"></div>');
        }
      }
      var c = collapsedEndRect(range);
      if (c) {
        html.push('<div class="mc-caret" style="left:' + (c.left - host.left) + 'px;top:' +
          (c.top - host.top) + 'px;height:' + (c.height || 20) + 'px"></div>');
      }
    }
    overlay.innerHTML = html.join('');
  }
  function render() {
    if (!mcMode || !carets.length) { clearOverlay(); return; }
    // Drop stale carets if the document was replaced underneath us (external reload)
    for (var d = 0; d < carets.length; d++) {
      if (!editor.contains(carets[d].startContainer)) { leaveMode(false); return; }
    }
    paintRanges(carets);
  }

  /* ---------- Caret bookkeeping ---------- */
  function cmp(a, b) { return a.compareBoundaryPoints(Range.START_TO_START, b); }
  function sortCarets() { carets.sort(cmp); }

  // Merge carets that are equal or overlapping (Sublime collapses coincident cursors)
  function dedupe() {
    if (carets.length < 2) return;
    sortCarets();
    var out = [carets[0]];
    for (var i = 1; i < carets.length; i++) {
      var prev = out[out.length - 1], cur = carets[i];
      // overlap if cur.start <= prev.end
      var overlap = cur.compareBoundaryPoints(Range.START_TO_END, prev) <= 0;
      if (overlap) {
        if (cur.compareBoundaryPoints(Range.END_TO_END, prev) > 0) {
          prev.setEnd(cur.endContainer, cur.endOffset);
        }
      } else {
        out.push(cur);
      }
    }
    carets = out;
  }

  function primary() { return carets.length ? carets[carets.length - 1] : null; }

  function enterMode() {
    dedupe();
    mcMode = carets.length >= 2;
    if (mcMode) {
      editor.classList.add('mc-active');
      editor.focus();
      syncNativeToPrimary();   // keep the native caret (and any IME preedit) on the primary
    } else {
      leaveMode(false);
    }
    render();
  }

  function leaveMode(collapseToPrimary) {
    var p = collapseToPrimary ? primary() : null;
    carets = [];
    mcMode = false;
    goalX = null;
    editor.classList.remove('mc-active');
    if (p) { sel.removeAllRanges(); p.collapse(false); sel.addRange(p); }
    clearOverlay();
  }

  function syncNativeToPrimary() {
    var p = primary();
    if (!p) return;
    var r = p.cloneRange();
    r.collapse(false);
    sel.removeAllRanges();
    sel.addRange(r);
  }

  // After a caret op, if everything merged down to one caret, drop to a single native caret
  function maybeExit() {
    if (carets.length < 2) { leaveMode(true); return true; }
    return false;
  }

  /* ---------- Editing across every caret ---------- */
  // Run `fn(sel)` once per caret (document order, LAST → FIRST so earlier ranges stay
  // valid), driving the native selection so execCommand / modify apply to that caret.
  // A DOM Range has no direction, but a selection does (anchor vs focus). We remember
  // it on the range object (_backward) so that extending LEFT/HOME keeps growing the
  // selection instead of collapsing it on the next keystroke.
  function selectionIsBackward() {
    if (!sel.rangeCount || sel.isCollapsed || !sel.anchorNode || !sel.focusNode) return false;
    if (sel.anchorNode === sel.focusNode) return sel.anchorOffset > sel.focusOffset;
    return !!(sel.anchorNode.compareDocumentPosition(sel.focusNode) & Node.DOCUMENT_POSITION_PRECEDING);
  }
  function applyCaretToSelection(r) {
    sel.removeAllRanges();
    if (!r.collapsed && r._backward) {
      try { sel.setBaseAndExtent(r.endContainer, r.endOffset, r.startContainer, r.startOffset); return; }
      catch (e) {}
    }
    sel.addRange(r);
  }
  function captureSelectionRange() {
    var r = sel.getRangeAt(0).cloneRange();
    r._backward = selectionIsBackward();
    return r;
  }

  function eachCaret(fn) {
    sortCarets();
    for (var i = carets.length - 1; i >= 0; i--) {
      applyCaretToSelection(carets[i]);
      fn(sel, i);
      if (sel.rangeCount) carets[i] = captureSelectionRange();
    }
  }

  function commitEdit(fn) {
    suppressInput = true;
    eachCaret(fn);
    dedupe();
    // execCommand already fired native 'input' events (editor.js saved); one more
    // guarantees a save even if a path was a no-op. Kept inside the suppress window
    // so our own input-guard below doesn't tear multi-cursor down.
    editor.dispatchEvent(new Event('input', { bubbles: true }));
    if (maybeExit()) { suppressInput = false; return; }
    syncNativeToPrimary();
    suppressInput = false;
    render();
  }

  function typeText(text) {
    commitEdit(function () { document.execCommand('insertText', false, text); });
  }
  function delBackward() {
    commitEdit(function () { document.execCommand('delete', false, null); });
  }
  function delForward() {
    commitEdit(function () { document.execCommand('forwardDelete', false, null); });
  }
  function splitLines() {
    commitEdit(function () { document.execCommand('insertParagraph', false, null); });
  }
  // Copy every selection's text (document order, newline-joined) to the clipboard; cut also
  // deletes them. A hidden textarea + execCommand('copy') works inside the Standard Notes
  // iframe where the async Clipboard API is blocked; navigator.clipboard is the fallback.
  function copyRanges(cut) {
    if (!carets.length) return;
    sortCarets();
    var parts = [];
    for (var i = 0; i < carets.length; i++) parts.push(carets[i].toString());
    var text = parts.join('\n');
    var ta = document.createElement('textarea');
    ta.value = text;
    ta.setAttribute('readonly', '');
    ta.style.cssText = 'position:fixed;left:-9999px;top:0;opacity:0';
    document.body.appendChild(ta);
    ta.select();
    var ok = false;
    try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    document.body.removeChild(ta);
    if (!ok && navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).catch(function () {});
    }
    editor.focus();
    if (cut) { delBackward(); }              // remove every selected range
    else { syncNativeToPrimary(); render(); }
  }

  /* ---------- Movement across every caret ---------- */
  function moveAll(alter, direction, granularity) {
    eachCaret(function (s) {
      try { s.modify(alter, direction, granularity); } catch (e) {}
    });
    dedupe();
    if (maybeExit()) return;
    syncNativeToPrimary();
    render();
  }

  /* ---------- Adding carets ---------- */
  function caretRangeFromPoint(x, y) {
    var r = null;
    if (document.caretRangeFromPoint) {
      r = document.caretRangeFromPoint(x, y);
    } else if (document.caretPositionFromPoint) {
      var p = document.caretPositionFromPoint(x, y);
      if (p) { r = document.createRange(); r.setStart(p.offsetNode, p.offset); r.collapse(true); }
    }
    if (r && editor.contains(r.startContainer)) return r;
    return null;
  }

  function initFromSelection() {
    if (!sel.rangeCount || !editor.contains(sel.anchorNode)) { editor.focus(); }
    if (sel.rangeCount && editor.contains(sel.getRangeAt(0).startContainer)) {
      carets = [sel.getRangeAt(0).cloneRange()];
    } else {
      carets = [];
    }
    goalX = null;
  }

  function addCaretVertical(dir) {
    if (!mcMode) initFromSelection();
    if (!carets.length) return;
    sortCarets();
    var anchor = dir < 0 ? carets[0] : carets[carets.length - 1];
    var rect = collapsedEndRect(anchor);
    if (!rect) return;
    if (goalX === null) goalX = rect.left;
    var lh = rect.height || 20;
    var y = dir < 0 ? rect.top - lh * 0.5 : rect.top + lh + lh * 0.5;
    var next = caretRangeFromPoint(goalX, y);
    if (next) {
      // Don't add a duplicate at an existing caret
      for (var i = 0; i < carets.length; i++) {
        if (cmp(carets[i], next) === 0 && carets[i].collapsed) { next = null; break; }
      }
    }
    if (next) { carets.push(next); enterMode(); }
    else if (!mcMode) { carets = []; }   // single line, nothing to add
  }

  function addCaretAtPoint(x, y) {
    if (!mcMode) initFromSelection();
    var next = caretRangeFromPoint(x, y);
    if (!next) return;
    // Toggle: Alt+click on an existing caret removes it
    for (var i = 0; i < carets.length; i++) {
      if (carets[i].collapsed && cmp(carets[i], next) === 0) {
        carets.splice(i, 1);
        if (carets.length < 2) leaveMode(true); else render();
        return;
      }
    }
    carets.push(next);
    goalX = null;
    enterMode();
  }

  /* ============================================================
     FIND BAR
     ============================================================ */
  var findBar = null, findInput = null, findCount = null, replaceInput = null;
  var matches = [];       // [{start:{node,off}, end:{node,off}}]
  var matchIndex = -1;
  var optCase = false, optWord = false, optRegex = false;   // Find options (Aa / whole word / .*)

  /* All-match highlighting via the CSS Custom Highlight API (paints every match as you
     type — stays visible even while focus is in the Find input). Degrades gracefully:
     when the API is missing, only the current match's native selection shows. */
  var findHi = null, findHiCur = null;
  if (window.CSS && CSS.highlights && window.Highlight) {
    try {
      findHi = new Highlight();
      findHiCur = new Highlight();
      CSS.highlights.set('rn-find', findHi);
      CSS.highlights.set('rn-find-current', findHiCur);
    } catch (e) { findHi = findHiCur = null; }
  }
  function clearHighlights() {
    if (findHi) findHi.clear();
    if (findHiCur) findHiCur.clear();
  }
  function paintHighlights() {
    if (!findHi) return;
    findHi.clear();
    for (var i = 0; i < matches.length; i++) findHi.add(rangeOf(matches[i]));
  }
  function paintCurrent() {
    if (!findHiCur) return;
    findHiCur.clear();
    if (matchIndex >= 0 && matchIndex < matches.length) findHiCur.add(rangeOf(matches[matchIndex]));
  }

  function buildFindBar() {
    findBar = document.createElement('div');
    findBar.className = 'mc-find';
    findBar.innerHTML =
      '<button type="button" class="mc-find-toggle" data-act="toggle" title="Toggle Replace"><svg class="ico" aria-hidden="true" focusable="false" viewBox="0 0 16 16"><path d="M6 5l3 3-3 3"/></svg></button>' +
      '<div class="mc-find-cols">' +
        '<div class="mc-find-row">' +
          '<span class="mc-find-ico"><svg class="ico" aria-hidden="true" focusable="false" viewBox="0 0 16 16"><circle cx="7" cy="7" r="4.3"/><path d="M10.2 10.2 14 14"/></svg></span>' +
          '<input type="text" class="mc-find-input" placeholder="Find in note…" spellcheck="false" />' +
          '<span class="mc-find-count">0/0</span>' +
          '<button type="button" class="mc-find-opt" data-act="optCase" title="Match case">Aa</button>' +
          '<button type="button" class="mc-find-opt" data-act="optWord" title="Whole word"><span class="mc-opt-word">ab</span></button>' +
          '<button type="button" class="mc-find-opt" data-act="optRegex" title="Use regular expression">.*</button>' +
          '<span class="mc-find-sep"></span>' +
          '<button type="button" class="mc-find-btn" data-act="prev" title="Previous (Shift+Enter)"><svg class="ico" aria-hidden="true" focusable="false" viewBox="0 0 16 16"><path d="M4 10l4-4 4 4"/></svg></button>' +
          '<button type="button" class="mc-find-btn" data-act="next" title="Next (Enter)"><svg class="ico" aria-hidden="true" focusable="false" viewBox="0 0 16 16"><path d="M4 6l4 4 4-4"/></svg></button>' +
          '<button type="button" class="mc-find-btn mc-find-all" data-act="all" title="Select all matches (Alt+Enter)">Select all</button>' +
          '<button type="button" class="mc-find-btn mc-find-close" data-act="close" title="Close (Esc)"><svg class="ico" aria-hidden="true" focusable="false" viewBox="0 0 16 16"><path d="M4 4l8 8M12 4l-8 8"/></svg></button>' +
        '</div>' +
        '<div class="mc-find-row mc-replace-row">' +
          '<span class="mc-find-ico"><svg class="ico" aria-hidden="true" focusable="false" viewBox="0 0 16 16"><path d="M3 6.5A4.5 4.5 0 0 1 11 4l1.6 1.6"/><path d="M12.8 3v3h-3"/><path d="M13 9.5A4.5 4.5 0 0 1 5 12l-1.6-1.6"/><path d="M3.2 13v-3h3"/></svg></span>' +
          '<input type="text" class="mc-replace-input" placeholder="Replace with…" spellcheck="false" />' +
          '<span class="mc-find-sep"></span>' +
          '<button type="button" class="mc-find-btn" data-act="replace" title="Replace (Enter)">Replace</button>' +
          '<button type="button" class="mc-find-btn mc-replace-all" data-act="replaceAll" title="Replace all (Alt+Enter)">All</button>' +
        '</div>' +
      '</div>';
    (editorArea || document.body).appendChild(findBar);
    findInput = findBar.querySelector('.mc-find-input');
    findCount = findBar.querySelector('.mc-find-count');
    replaceInput = findBar.querySelector('.mc-replace-input');

    findInput.addEventListener('input', function () { computeMatches(); if (matches.length) gotoMatch(0); else paintCurrent(); });
    findInput.addEventListener('keydown', function (ev) {
      if (ev.key === 'Enter' && ev.altKey) { ev.preventDefault(); selectAllMatches(); }
      else if (ev.key === 'Enter') { ev.preventDefault(); gotoMatch(matchIndex + (ev.shiftKey ? -1 : 1)); }
      else if (ev.key === 'Escape') { ev.preventDefault(); closeFind(); }
    });
    replaceInput.addEventListener('keydown', function (ev) {
      if (ev.key === 'Enter' && ev.altKey) { ev.preventDefault(); replaceAll(); }
      else if (ev.key === 'Enter') { ev.preventDefault(); replaceOne(); }
      else if (ev.key === 'Escape') { ev.preventDefault(); closeFind(); }
    });
    findBar.addEventListener('mousedown', function (ev) {
      // keep the editor selection/caret when clicking buttons
      if (ev.target.closest('.mc-find-btn, .mc-find-toggle, .mc-find-opt')) ev.preventDefault();
    });
    findBar.addEventListener('click', function (ev) {
      var opt = ev.target.closest('.mc-find-opt');
      if (opt) {
        if (opt.dataset.act === 'optCase') optCase = !optCase;
        else if (opt.dataset.act === 'optWord') optWord = !optWord;
        else if (opt.dataset.act === 'optRegex') optRegex = !optRegex;
        syncOptButtons();
        computeMatches();
        if (matches.length) gotoMatch(0); else paintCurrent();
        findInput.focus();
        return;
      }
      var btn = ev.target.closest('.mc-find-btn, .mc-find-toggle'); if (!btn) return;
      var act = btn.dataset.act;
      if (act === 'next') gotoMatch(matchIndex + 1);
      else if (act === 'prev') gotoMatch(matchIndex - 1);
      else if (act === 'all') selectAllMatches();
      else if (act === 'replace') replaceOne();
      else if (act === 'replaceAll') replaceAll();
      else if (act === 'toggle') toggleReplace();
      else if (act === 'close') closeFind();
    });
    syncOptButtons();
  }
  function syncOptButtons() {
    if (!findBar) return;
    var map = { optCase: optCase, optWord: optWord, optRegex: optRegex };
    var btns = findBar.querySelectorAll('.mc-find-opt');
    for (var i = 0; i < btns.length; i++) btns[i].classList.toggle('active', !!map[btns[i].dataset.act]);
  }

  function toggleReplace(force) {
    if (!findBar) return;
    var on = (force === undefined) ? !findBar.classList.contains('mc-find-expanded') : force;
    findBar.classList.toggle('mc-find-expanded', on);
    if (on && replaceInput) { replaceInput.focus(); replaceInput.select(); }
  }

  function openFind() {
    if (!findBar) buildFindBar();
    findBar.classList.add('open');
    var seed = (sel.rangeCount && editor.contains(sel.anchorNode)) ? sel.toString() : '';
    if (seed && seed.indexOf('\n') < 0) findInput.value = seed;
    computeMatches();
    findInput.focus();
    findInput.select();
    if (matches.length) gotoMatch(0);
  }
  function closeFind() {
    if (!findBar) return;
    findBar.classList.remove('open');
    matches = []; matchIndex = -1;
    clearHighlights();
    editor.focus();
  }
  function findOpen() { return findBar && findBar.classList.contains('open'); }

  // Exposed so the menu (Edit → Find / Find & Replace) can open the bar on devices
  // without a keyboard (mobile). openReplace also expands the Replace row.
  window.__richnoteFind = {
    open: function () { openFind(); },
    openReplace: function () { openFind(); if (findBar) findBar.classList.add('mc-find-expanded'); }
  };

  // Exposed so the toolbar / shortcuts can apply a formatting command to EVERY range of
  // a discontiguous (Ctrl+drag / multi-cursor) selection, not just the primary one.
  window.__richnoteMC = {
    active: function () { return mcMode; },
    run: function (fn) { if (mcMode) commitEdit(fn); }
  };

  // Flatten editor text nodes so we can map a global string index back to (node, offset)
  function textIndex() {
    var walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT, null);
    var nodes = [], text = '', node;
    while ((node = walker.nextNode())) {
      nodes.push({ node: node, start: text.length });
      text += node.nodeValue;
    }
    return { text: text, nodes: nodes };
  }
  // Map a global string index to a (node, offset). A boundary index resolves to the
  // LATER node for a match start, and the EARLIER node for a match end — so a match that
  // touches a block boundary stays inside a single text node instead of spanning it.
  function locateStart(idx, nodes) {
    for (var i = 0; i < nodes.length; i++) {
      var n = nodes[i], len = n.node.nodeValue.length;
      if (idx >= n.start && idx < n.start + len) return { node: n.node, off: idx - n.start };
    }
    var last = nodes[nodes.length - 1];
    return last ? { node: last.node, off: last.node.nodeValue.length } : null;
  }
  function locateEnd(idx, nodes) {
    for (var i = 0; i < nodes.length; i++) {
      var n = nodes[i], len = n.node.nodeValue.length;
      if (idx > n.start && idx <= n.start + len) return { node: n.node, off: idx - n.start };
    }
    var first = nodes[0];
    return first ? { node: first.node, off: 0 } : null;
  }
  function topBlockOf(node) {
    var n = node;
    while (n && n.parentNode !== editor) n = n.parentNode;
    return (n && n.parentNode === editor) ? n : null;
  }
  // Build a RegExp from the query + the active options. Returns null (empty query) or
  // the string 'error' (invalid regex the user is still typing).
  function buildMatcher() {
    var q = findInput ? findInput.value : '';
    if (!q) return null;
    var flags = optCase ? 'g' : 'gi';
    if (optRegex) { try { return new RegExp(q, flags); } catch (e) { return 'error'; } }
    var esc = q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    if (optWord) esc = '(?:^|\\W)(' + esc + ')(?=\\W|$)';   // whole-word (no \b — matches non-ASCII too)
    return { re: new RegExp(esc, flags), group: optWord ? 1 : 0 };
  }
  // Scan the flat text for [start, end] index pairs, honouring the options. null / 'error' pass through.
  function scanMatches(text) {
    var mk = buildMatcher();
    if (!mk || mk === 'error') return mk;
    var re = mk.re || mk, group = mk.group || 0;
    var out = [], m, guard = 0;
    re.lastIndex = 0;
    while ((m = re.exec(text))) {
      var whole = m[0], part = group ? m[group] : whole;
      var start = m.index + (group ? whole.indexOf(part) : 0);
      var end = start + part.length;
      if (part.length === 0) { re.lastIndex = m.index + 1; if (re.lastIndex > text.length) break; }
      else { out.push([start, end]); if (re.lastIndex <= start) re.lastIndex = end; }
      if (++guard > 100000) break;
    }
    return out;
  }
  function setFindError(on) { if (findInput) findInput.classList.toggle('error', !!on); }

  function computeMatches() {
    matches = []; matchIndex = -1;
    var q = findInput ? findInput.value : '';
    if (!q) { if (findCount) findCount.textContent = '0/0'; setFindError(false); clearHighlights(); return; }
    var idx = textIndex();
    var found = scanMatches(idx.text);
    if (found === 'error') { setFindError(true); if (findCount) findCount.textContent = '0/0'; clearHighlights(); return; }
    setFindError(false);
    for (var i = 0; i < found.length; i++) {
      var s = locateStart(found[i][0], idx.nodes), e = locateEnd(found[i][1], idx.nodes);
      // Skip matches that would span two blocks (editing them could merge the blocks)
      if (s && e && topBlockOf(s.node) === topBlockOf(e.node)) matches.push({ s: s, e: e });
    }
    if (findCount) findCount.textContent = (matches.length ? 1 : 0) + '/' + matches.length;
    paintHighlights();   // highlight every match live as the user types
  }
  function rangeOf(m) {
    var r = document.createRange();
    r.setStart(m.s.node, m.s.off);
    r.setEnd(m.e.node, m.e.off);
    return r;
  }
  function gotoMatch(i) {
    if (!matches.length) return;
    matchIndex = ((i % matches.length) + matches.length) % matches.length;
    var r = rangeOf(matches[matchIndex]);
    // Setting a selection inside the contenteditable steals focus from the Find/Replace
    // input (so the next keystroke would land in the editor). Remember the focused field
    // and restore it — the current match stays visible via the Highlight API.
    var active = document.activeElement;
    var keepFocus = findBar && findBar.contains(active) ? active : null;
    sel.removeAllRanges(); sel.addRange(r);
    var rect = r.getBoundingClientRect();
    var host = editor.getBoundingClientRect();
    if (rect.top < host.top || rect.bottom > host.bottom) {
      editor.scrollTop += rect.top - host.top - host.height / 2;
    }
    paintCurrent();
    if (keepFocus) keepFocus.focus();
    if (findCount) findCount.textContent = (matchIndex + 1) + '/' + matches.length;
  }

  /* ---------- Replace ---------- */
  // insertText keeps the browser undo stack and fires 'input' (editor.js auto-saves).
  // It targets the active editable, so we focus the editor and set the selection first.
  function replaceOne() {
    if (!matches.length) return;
    if (matchIndex < 0) { gotoMatch(0); return; }
    var repl = replaceInput ? replaceInput.value : '';
    var keep = matchIndex;
    var r = rangeOf(matches[matchIndex]);
    editor.focus();
    sel.removeAllRanges(); sel.addRange(r);
    document.execCommand('insertText', false, repl);
    computeMatches();
    if (matches.length) gotoMatch(keep <= matches.length - 1 ? keep : 0);
    else { matchIndex = -1; clearHighlights(); if (findCount) findCount.textContent = '0/0'; }
    if (replaceInput) replaceInput.focus();
  }
  function replaceAll() {
    var q = findInput ? findInput.value : '';
    if (!q) return;
    var repl = replaceInput ? replaceInput.value : '';
    // Collect every match's GLOBAL [start,end] up-front, then replace LAST → FIRST so
    // earlier indices stay accurate; re-resolve (node, offset) from a fresh text map each
    // time so node splits/merges from execCommand can't invalidate a stale Range.
    var idx = textIndex();
    var found = scanMatches(idx.text);
    if (!found || found === 'error' || !found.length) return;
    var spans = [];
    for (var i = 0; i < found.length; i++) {
      var s = locateStart(found[i][0], idx.nodes), e = locateEnd(found[i][1], idx.nodes);
      if (s && e && topBlockOf(s.node) === topBlockOf(e.node)) spans.push(found[i]);
    }
    if (!spans.length) return;
    editor.focus();
    for (var j = spans.length - 1; j >= 0; j--) {
      var map = textIndex();
      var ss = locateStart(spans[j][0], map.nodes);
      var ee = locateEnd(spans[j][1], map.nodes);
      if (!ss || !ee) continue;
      var rr = document.createRange();
      rr.setStart(ss.node, ss.off); rr.setEnd(ee.node, ee.off);
      sel.removeAllRanges(); sel.addRange(rr);
      document.execCommand('insertText', false, repl);
    }
    computeMatches();
    if (matches.length) gotoMatch(0);
    else { matchIndex = -1; clearHighlights(); if (findCount) findCount.textContent = '0/0'; }
    if (replaceInput) replaceInput.focus();
  }

  function selectAllMatches() {
    if (!matches.length) return;
    carets = matches.map(rangeOf);
    goalX = null;
    closeFind();
    enterMode();
  }

  /* ============================================================
     KEY & MOUSE HANDLING  (capture phase → runs before editor.js)
     ============================================================ */
  // Attached on document in the CAPTURE phase so it runs before editor.js's own
  // keydown handler (which is on the editor element, registered earlier).
  document.addEventListener('keydown', function (ev) {
    if (document.activeElement !== editor) return;   // only when the editor is focused
    var mod = ev.ctrlKey || ev.metaKey;

    // Ctrl+F — open Find (works whether or not multi-cursor is active)
    if (mod && !ev.altKey && !ev.shiftKey && (ev.key === 'f' || ev.key === 'F')) {
      ev.preventDefault(); ev.stopImmediatePropagation();
      openFind();
      return;
    }

    // Shift+Alt+Up / Down — add a caret above / below
    if (ev.altKey && ev.shiftKey && !mod && (ev.key === 'ArrowUp' || ev.key === 'ArrowDown')) {
      ev.preventDefault(); ev.stopImmediatePropagation();
      addCaretVertical(ev.key === 'ArrowUp' ? -1 : 1);
      return;
    }

    if (!mcMode) return;   // below here: only while multi-cursor is active

    // Escape collapses to a single caret
    if (ev.key === 'Escape') {
      ev.preventDefault(); ev.stopImmediatePropagation();
      leaveMode(true);
      return;
    }

    // Arrows / Home / End — move (or, with Shift, extend) every caret. Honour the same
    // modifiers native single-caret editing does: Ctrl/Alt+Left/Right jumps by word, and
    // Ctrl+Home/End jumps to the document boundary (else per line).
    var alter = ev.shiftKey ? 'extend' : 'move';
    var hGran = (ev.ctrlKey || ev.altKey) ? 'word' : 'character';
    var vBound = (ev.ctrlKey || ev.metaKey) ? 'documentboundary' : 'lineboundary';
    if (ev.key === 'ArrowLeft')  { ev.preventDefault(); ev.stopImmediatePropagation(); goalX = null; moveAll(alter, 'left', hGran); return; }
    if (ev.key === 'ArrowRight') { ev.preventDefault(); ev.stopImmediatePropagation(); goalX = null; moveAll(alter, 'right', hGran); return; }
    if (ev.key === 'ArrowUp')    { ev.preventDefault(); ev.stopImmediatePropagation(); moveAll(alter, 'backward', 'line'); return; }
    if (ev.key === 'ArrowDown')  { ev.preventDefault(); ev.stopImmediatePropagation(); moveAll(alter, 'forward', 'line'); return; }
    if (ev.key === 'Home')       { ev.preventDefault(); ev.stopImmediatePropagation(); goalX = null; moveAll(alter, 'backward', vBound); return; }
    if (ev.key === 'End')        { ev.preventDefault(); ev.stopImmediatePropagation(); goalX = null; moveAll(alter, 'forward', vBound); return; }

    // Editing
    if (ev.key === 'Backspace') { ev.preventDefault(); ev.stopImmediatePropagation(); delBackward(); return; }
    if (ev.key === 'Delete')    { ev.preventDefault(); ev.stopImmediatePropagation(); delForward(); return; }
    if (ev.key === 'Enter')     { ev.preventDefault(); ev.stopImmediatePropagation(); splitLines(); return; }
    if (ev.key === 'Tab')       { ev.preventDefault(); ev.stopImmediatePropagation(); typeText('\t'); return; }

    // Ctrl+C / Ctrl+X — copy (and cut) every selected range, joined by newlines (Sublime-style)
    if (mod && !ev.shiftKey && !ev.altKey && (ev.key === 'c' || ev.key === 'C')) {
      ev.preventDefault(); ev.stopImmediatePropagation(); copyRanges(false); return;
    }
    if (mod && !ev.shiftKey && !ev.altKey && (ev.key === 'x' || ev.key === 'X')) {
      ev.preventDefault(); ev.stopImmediatePropagation(); copyRanges(true); return;
    }

    // A modifier key pressed on its own (Ctrl / Alt / Shift / Meta) fires a keydown whose
    // own modifier flag is already set — ignore it. Otherwise merely holding Ctrl to reach a
    // shortcut like Ctrl+B would collapse every caret before the letter key even arrives.
    if (ev.key === 'Control' || ev.key === 'Shift' || ev.key === 'Alt' || ev.key === 'Meta') return;

    // Ctrl/Cmd + A / V / Z / Y — select-all, paste, undo and redo don't map onto multiple
    // carets: drop to normal editing and let the browser handle it once.
    if (mod && (ev.code === 'KeyA' || ev.code === 'KeyV' || ev.code === 'KeyZ' || ev.code === 'KeyY')) {
      leaveMode(false); return;
    }

    // Any other Ctrl/Cmd combo (Ctrl+B/I/U, alignment, headings, sub/super, clear, …) stays in
    // multi-cursor: fall through so editor.js applies it to EVERY caret via its mcActive() routing.
    if (mod) return;

    // Printable character
    if (ev.key.length === 1 && !ev.ctrlKey && !ev.metaKey && !ev.altKey) {
      ev.preventDefault(); ev.stopImmediatePropagation();
      typeText(ev.key);
      return;
    }
  }, true);

  /* ---------- Sublime-style mouse multi-selection ----------
     • Ctrl/Cmd + drag  → add another highlighted range (existing ones stay painted — no flicker)
     • Ctrl/Cmd + click → add a caret; clicking on an existing caret/selection removes it (toggle)
     • Ctrl/Cmd + click on a link is left to editor.js (open link)
     We drive the new range ourselves (caretRangeFromPoint) instead of the browser's native
     selection, so the other selections never disappear mid-drag. */
  var ctrlDrag = null;

  // Finish a Ctrl selection edit: 2+ ranges → multi-cursor; exactly 1 → a normal selection.
  function finalizeCarets() {
    goalX = null;
    dedupe();
    if (carets.length >= 2) {
      mcMode = true;
      editor.classList.add('mc-active');
      editor.focus();
      syncNativeToPrimary();
      render();
    } else {
      var one = carets.length ? carets[0].cloneRange() : null;
      carets = [];
      mcMode = false;
      editor.classList.remove('mc-active');
      clearOverlay();
      editor.focus();
      if (one) applyCaretToSelection(one);
    }
  }
  function rangeHitsPoint(range, node, offset) {
    try {
      if (range.collapsed) return range.startContainer === node && range.startOffset === offset;
      return range.comparePoint(node, offset) === 0;
    } catch (e) { return false; }
  }

  document.addEventListener('mousedown', function (ev) {
    if (!editor.contains(ev.target)) return;
    if (ev.altKey && ev.button === 0) {
      ev.preventDefault();
      addCaretAtPoint(ev.clientX, ev.clientY);
      return;
    }
    if ((ev.ctrlKey || ev.metaKey) && ev.button === 0) {
      if (ev.target.closest && ev.target.closest('a')) return;          // Ctrl+click a link → open it
      if (ev.target.closest && ev.target.closest('td,th')) return;      // tables own their own drags
      var anchor = caretRangeFromPoint(ev.clientX, ev.clientY);
      if (!anchor) return;
      ev.preventDefault();
      // Preserve the current selection/caret (collapsed or not) as the base set.
      var base = mcMode ? carets.slice()
        : (sel.rangeCount && editor.contains(sel.anchorNode) ? [sel.getRangeAt(0).cloneRange()] : []);
      ctrlDrag = { base: base, anchor: anchor.cloneRange(), pending: anchor.cloneRange(), moved: false };
      sel.removeAllRanges();                                            // overlay owns the visuals now
      editor.focus();
      paintRanges(base.concat([ctrlDrag.pending]));                     // show base + the new caret at once
      window.addEventListener('mousemove', onCtrlDragMove, true);
      window.addEventListener('mouseup', onCtrlDragUp, true);
      return;
    }
    if (mcMode) leaveMode(false);   // plain click → let the browser place a single caret
  }, true);

  function onCtrlDragMove(ev) {
    if (!ctrlDrag) return;
    var pt = caretRangeFromPoint(ev.clientX, ev.clientY);
    if (!pt) return;
    ctrlDrag.moved = true;
    var a = ctrlDrag.anchor, rng = document.createRange();
    if (a.compareBoundaryPoints(Range.START_TO_START, pt) <= 0) {
      rng.setStart(a.startContainer, a.startOffset); rng.setEnd(pt.startContainer, pt.startOffset); rng._backward = false;
    } else {
      rng.setStart(pt.startContainer, pt.startOffset); rng.setEnd(a.startContainer, a.startOffset); rng._backward = true;
    }
    ctrlDrag.pending = rng;
    paintRanges(ctrlDrag.base.concat([rng]));
  }
  function onCtrlDragUp() {
    window.removeEventListener('mousemove', onCtrlDragMove, true);
    window.removeEventListener('mouseup', onCtrlDragUp, true);
    var cd = ctrlDrag; ctrlDrag = null;
    if (!cd) return;
    var base = cd.base, pending = cd.pending;
    // A click (no drag) on an existing caret/selection toggles it off.
    if (pending.collapsed && !cd.moved) {
      for (var i = 0; i < base.length; i++) {
        if (rangeHitsPoint(base[i], pending.startContainer, pending.startOffset)) {
          base.splice(i, 1); carets = base; finalizeCarets(); return;
        }
      }
    }
    base.push(pending);
    carets = base;
    finalizeCarets();
  }

  // Keep the overlay glued to the carets as the view changes
  editor.addEventListener('scroll', render);
  window.addEventListener('resize', render);
  window.addEventListener('scroll', render, true);

  // If the note is reloaded / heavily edited from outside, drop multi-cursor state.
  // Never tear down mid-composition — the IME's intermediate 'input' events are expected.
  editor.addEventListener('input', function () { if (suppressInput || composing) return; if (mcMode) { leaveMode(false); } });

  /* ---------- IME / preedit (e.g. Vietnamese Telex on Ubuntu IBus/fcitx) ----------
     The browser runs the composition on the primary caret (native selection), so the
     preedit popup shows there. We just hold multi-cursor open during composition, then
     on commit replay the composed string onto every OTHER caret. */
  editor.addEventListener('compositionstart', function () {
    if (mcMode) composing = true;
  });
  editor.addEventListener('compositionend', function (e) {
    if (!composing) return;
    composing = false;
    if (!mcMode) return;
    var data = e.data || '';
    // Hold suppression across this whole handler AND the composition's trailing 'input'
    // event (which fires right after in the same tick) so multi-cursor isn't torn down.
    suppressInput = true;
    sortCarets();
    // The IME committed at the primary (bottom-most) caret — refresh it from the selection.
    if (sel.rangeCount && editor.contains(sel.anchorNode)) {
      carets[carets.length - 1] = sel.getRangeAt(0).cloneRange();
    }
    // Replay the committed text onto every other caret.
    if (data) {
      for (var i = carets.length - 2; i >= 0; i--) {
        sel.removeAllRanges();
        sel.addRange(carets[i]);
        document.execCommand('insertText', false, data);
        carets[i] = sel.getRangeAt(0).cloneRange();
      }
    }
    dedupe();
    if (!maybeExit()) { syncNativeToPrimary(); render(); }
    setTimeout(function () { suppressInput = false; }, 0);
  });
})();
