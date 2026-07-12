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

  function render() {
    if (!mcMode || !carets.length) { clearOverlay(); return; }
    // Drop stale carets if the document was replaced underneath us (external reload)
    for (var d = 0; d < carets.length; d++) {
      if (!editor.contains(carets[d].startContainer)) { leaveMode(false); return; }
    }
    var host = editor.getBoundingClientRect();
    overlay.style.display = 'block';
    overlay.style.left = host.left + 'px';
    overlay.style.top = host.top + 'px';
    overlay.style.width = host.width + 'px';
    overlay.style.height = host.height + 'px';
    var html = [];
    for (var i = 0; i < carets.length; i++) {
      var range = carets[i];
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
  var findBar = null, findInput = null, findCount = null;
  var matches = [];       // [{start:{node,off}, end:{node,off}}]
  var matchIndex = -1;

  function buildFindBar() {
    findBar = document.createElement('div');
    findBar.className = 'mc-find';
    findBar.innerHTML =
      '<span class="mc-find-ico"><svg class="ico" viewBox="0 0 16 16"><circle cx="7" cy="7" r="4.3"/><path d="M10.2 10.2 14 14"/></svg></span>' +
      '<input type="text" class="mc-find-input" placeholder="Find in note…" spellcheck="false" />' +
      '<span class="mc-find-count">0/0</span>' +
      '<span class="mc-find-sep"></span>' +
      '<button type="button" class="mc-find-btn" data-act="prev" title="Previous (Shift+Enter)"><svg class="ico" viewBox="0 0 16 16"><path d="M4 10l4-4 4 4"/></svg></button>' +
      '<button type="button" class="mc-find-btn" data-act="next" title="Next (Enter)"><svg class="ico" viewBox="0 0 16 16"><path d="M4 6l4 4 4-4"/></svg></button>' +
      '<button type="button" class="mc-find-btn mc-find-all" data-act="all" title="Select all matches (Alt+Enter)">Select all</button>' +
      '<button type="button" class="mc-find-btn mc-find-close" data-act="close" title="Close (Esc)"><svg class="ico" viewBox="0 0 16 16"><path d="M4 4l8 8M12 4l-8 8"/></svg></button>';
    (editorArea || document.body).appendChild(findBar);
    findInput = findBar.querySelector('.mc-find-input');
    findCount = findBar.querySelector('.mc-find-count');

    findInput.addEventListener('input', function () { computeMatches(); if (matches.length) gotoMatch(0); });
    findInput.addEventListener('keydown', function (ev) {
      if (ev.key === 'Enter' && ev.altKey) { ev.preventDefault(); selectAllMatches(); }
      else if (ev.key === 'Enter') { ev.preventDefault(); gotoMatch(matchIndex + (ev.shiftKey ? -1 : 1)); }
      else if (ev.key === 'Escape') { ev.preventDefault(); closeFind(); }
    });
    findBar.addEventListener('mousedown', function (ev) {
      // keep the editor selection when clicking buttons
      if (ev.target.closest('.mc-find-btn')) ev.preventDefault();
    });
    findBar.addEventListener('click', function (ev) {
      var btn = ev.target.closest('.mc-find-btn'); if (!btn) return;
      var act = btn.dataset.act;
      if (act === 'next') gotoMatch(matchIndex + 1);
      else if (act === 'prev') gotoMatch(matchIndex - 1);
      else if (act === 'all') selectAllMatches();
      else if (act === 'close') closeFind();
    });
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
    editor.focus();
  }
  function findOpen() { return findBar && findBar.classList.contains('open'); }

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
  function computeMatches() {
    matches = []; matchIndex = -1;
    var q = findInput ? findInput.value : '';
    if (!q) { if (findCount) findCount.textContent = '0/0'; return; }
    var idx = textIndex();
    var hay = idx.text.toLowerCase(), needle = q.toLowerCase();
    var from = 0, at;
    while ((at = hay.indexOf(needle, from)) !== -1) {
      var s = locateStart(at, idx.nodes), e = locateEnd(at + q.length, idx.nodes);
      // Skip matches that would span two blocks (editing them could merge the blocks)
      if (s && e && topBlockOf(s.node) === topBlockOf(e.node)) matches.push({ s: s, e: e });
      from = at + Math.max(1, q.length);
    }
    if (findCount) findCount.textContent = (matches.length ? 1 : 0) + '/' + matches.length;
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
    sel.removeAllRanges(); sel.addRange(r);
    var rect = r.getBoundingClientRect();
    var host = editor.getBoundingClientRect();
    if (rect.top < host.top || rect.bottom > host.bottom) {
      editor.scrollTop += rect.top - host.top - host.height / 2;
    }
    if (findCount) findCount.textContent = (matchIndex + 1) + '/' + matches.length;
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

    // Arrows / Home / End — move (or, with Shift, extend) every caret
    var alter = ev.shiftKey ? 'extend' : 'move';
    if (ev.key === 'ArrowLeft')  { ev.preventDefault(); ev.stopImmediatePropagation(); goalX = null; moveAll(alter, 'left', 'character'); return; }
    if (ev.key === 'ArrowRight') { ev.preventDefault(); ev.stopImmediatePropagation(); goalX = null; moveAll(alter, 'right', 'character'); return; }
    if (ev.key === 'ArrowUp')    { ev.preventDefault(); ev.stopImmediatePropagation(); moveAll(alter, 'backward', 'line'); return; }
    if (ev.key === 'ArrowDown')  { ev.preventDefault(); ev.stopImmediatePropagation(); moveAll(alter, 'forward', 'line'); return; }
    if (ev.key === 'Home')       { ev.preventDefault(); ev.stopImmediatePropagation(); goalX = null; moveAll(alter, 'backward', 'lineboundary'); return; }
    if (ev.key === 'End')        { ev.preventDefault(); ev.stopImmediatePropagation(); goalX = null; moveAll(alter, 'forward', 'lineboundary'); return; }

    // Editing
    if (ev.key === 'Backspace') { ev.preventDefault(); ev.stopImmediatePropagation(); delBackward(); return; }
    if (ev.key === 'Delete')    { ev.preventDefault(); ev.stopImmediatePropagation(); delForward(); return; }
    if (ev.key === 'Enter')     { ev.preventDefault(); ev.stopImmediatePropagation(); splitLines(); return; }
    if (ev.key === 'Tab')       { ev.preventDefault(); ev.stopImmediatePropagation(); typeText('\t'); return; }

    // Ctrl+A / Ctrl+C / Ctrl+X / Ctrl+V / Ctrl+Z with many carets → drop to normal editing
    if (mod) { leaveMode(false); return; }   // let the browser handle it once

    // Printable character
    if (ev.key.length === 1 && !ev.ctrlKey && !ev.metaKey && !ev.altKey) {
      ev.preventDefault(); ev.stopImmediatePropagation();
      typeText(ev.key);
      return;
    }
  }, true);

  // Alt+Click adds/removes a caret; a plain click collapses multi-cursor
  document.addEventListener('mousedown', function (ev) {
    if (!editor.contains(ev.target)) return;
    if (ev.altKey && ev.button === 0) {
      ev.preventDefault();
      addCaretAtPoint(ev.clientX, ev.clientY);
      return;
    }
    if (mcMode) leaveMode(false);   // plain click → let the browser place a single caret
  }, true);

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
