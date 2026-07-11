/*!
 * RichNote — Minimap (document overview, VS Code / Sublime style).
 *
 * Renders a scaled-down live clone of the note on the right edge, with a viewport
 * box marking the visible region. Click or drag the minimap to scroll. Toggle via
 * View → Minimap (state is remembered). Hidden on small screens.
 *
 * @author  Nguyen Huu Duc <nguyenhuuduc.it.91@gmail.com> (VIETIS)
 * @license MIT
 */
(function () {
  'use strict';

  var editor = document.getElementById('editor');
  var area = editor && editor.closest ? editor.closest('.editor-area') : null;
  if (!editor || !area) return;

  /* ---------- DOM ---------- */
  var minimap = document.createElement('div');
  minimap.className = 'minimap';
  minimap.id = 'minimap';
  minimap.setAttribute('aria-hidden', 'true');
  var content = document.createElement('div');
  var viewport = document.createElement('div');
  viewport.className = 'minimap-viewport';
  minimap.appendChild(content);
  minimap.appendChild(viewport);
  area.appendChild(minimap);

  var menuItem = document.getElementById('minimapItem');
  var scale = 1, contentTop = 0;

  var on = true;
  try { var saved = localStorage.getItem('richnote-minimap'); if (saved !== null) on = saved === '1'; } catch (e) {}

  function usable() { return on && minimap.clientWidth > 0; }

  /* ---------- Rendering ---------- */
  // Re-clone the note into the minimap and recompute the width scale.
  function renderContent() {
    if (!usable()) return;
    content.className = 'doc minimap-content' + (editor.classList.contains('wrap') ? ' wrap' : '');
    content.innerHTML = editor.innerHTML;
    var cur = content.querySelectorAll('.current-line');
    for (var i = 0; i < cur.length; i++) cur[i].classList.remove('current-line');
    var cw = editor.clientWidth || 1;
    content.style.width = cw + 'px';
    scale = (minimap.clientWidth || 1) / cw;
    updateViewport();
  }

  // Cheap: reposition the scaled clone + the viewport box as the editor scrolls.
  function updateViewport() {
    if (!usable()) return;
    var mmH = minimap.clientHeight || 1;
    var scaledH = editor.scrollHeight * scale;
    var maxScroll = editor.scrollHeight - editor.clientHeight;
    var ratio = maxScroll > 0 ? editor.scrollTop / maxScroll : 0;
    var slide = Math.max(0, scaledH - mmH);         // how far the thumbnail slides
    contentTop = -ratio * slide;
    content.style.transform = 'translateY(' + contentTop + 'px) scale(' + scale + ')';
    viewport.style.top = (editor.scrollTop * scale + contentTop) + 'px';
    viewport.style.height = (editor.clientHeight * scale) + 'px';
  }

  var renderTimer = null;
  function scheduleRender() {
    if (!on) return;
    if (renderTimer) return;
    renderTimer = setTimeout(function () { renderTimer = null; renderContent(); }, 150);
  }

  // Throttle viewport updates to one per frame for smooth scrolling
  var vpRaf = false;
  editor.addEventListener('scroll', function () {
    if (vpRaf) return;
    vpRaf = true;
    requestAnimationFrame(function () { vpRaf = false; updateViewport(); });
  }, { passive: true });
  window.addEventListener('resize', scheduleRender);
  // Catch typing, formatting and programmatic note loads (which fire no 'input')
  new MutationObserver(scheduleRender).observe(editor, {
    childList: true, subtree: true, characterData: true, attributes: true
  });

  /* ---------- Click / drag to scroll ---------- */
  var dragging = false;
  function clampScroll(v) { return Math.max(0, Math.min(v, editor.scrollHeight - editor.clientHeight)); }
  // Map a pointer y (viewport coords) to the scrollTop that centres the view there.
  // Instant (no smooth) so dragging tracks the cursor exactly, like a scrollbar.
  function scrollToPointer(clientY) {
    var py = clientY - minimap.getBoundingClientRect().top;
    editor.scrollTop = clampScroll((py - contentTop) / scale - editor.clientHeight / 2);
  }
  minimap.addEventListener('mousedown', function (ev) {
    ev.preventDefault();
    dragging = true;
    scrollToPointer(ev.clientY);
  });
  window.addEventListener('mousemove', function (ev) {
    if (!dragging) return;
    scrollToPointer(ev.clientY);
  });
  window.addEventListener('mouseup', function () { dragging = false; });
  // Mouse wheel over the minimap scrolls the editor (it has no scroll of its own)
  minimap.addEventListener('wheel', function (ev) {
    ev.preventDefault();
    editor.scrollTop = clampScroll(editor.scrollTop + ev.deltaY);
  }, { passive: false });

  /* ---------- Toggle ---------- */
  function apply() {
    minimap.style.display = on ? '' : 'none';   // '' lets the mobile media query hide it
    editor.classList.toggle('mm-on', on);       // reserves right padding so text clears the minimap
    if (menuItem) menuItem.classList.toggle('checked', on);
    if (on) renderContent();
  }
  function toggle() {
    on = !on;
    try { localStorage.setItem('richnote-minimap', on ? '1' : '0'); } catch (e) {}
    apply();
  }
  window.__richnoteMinimap = { toggle: toggle, refresh: scheduleRender };

  apply();
})();
