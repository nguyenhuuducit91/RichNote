<div align="center">

<img src="Images/icon.png" alt="RichNote" width="96" />

# RichNote

### The rich‑text editor Standard Notes was missing.

A fast, beautiful **WYSIWYG editor** for [Standard Notes](https://standardnotes.com) — write with **live formatting**, a **Word‑like toolbar**, headings, colors, lists, line numbers and keyboard shortcuts. No Markdown syntax to remember. What you see is what you get.

![Standard Notes](https://img.shields.io/badge/Standard%20Notes-editor-2f6fdb)
![No build](https://img.shields.io/badge/build-none%20(static)-success)
![License](https://img.shields.io/badge/license-MIT-blue)
![Made in Vietnam](https://img.shields.io/badge/made%20in-Vietnam%20%F0%9F%87%BB%F0%9F%87%B3-red)

<img src="Images/screenshot.png" alt="RichNote screenshot" width="820" />

</div>

---

## ✨ Why you'll love it

- 🪄 **True WYSIWYG** — bold looks bold, links look like links. No raw `**markdown**`.
- 🧰 **Word‑like toolbar** — style dropdown (Normal / **H1–H6** / Quote), font, size, **B / I / U / S**, inline code, and more, all on **one clean line**.
- 🎨 **Text & highlight colors** — a **Google‑Sheets‑style palette** plus any custom color.
- 🔢 **Line numbers + current‑line highlight** — a code‑editor feel for your notes.
- ⌨️ **Keyboard shortcuts everyone knows** — `Ctrl+B/I/U`, `Ctrl+Alt+1…6` for headings, `Ctrl+L/E/R` to align, `Ctrl+K` for links, `F4` to repeat the last action.
- 🔗 **Smart links** — insert/edit with a popup; **Ctrl/Cmd‑click** to open.
- ↹ **Tab indentation** (single line or a whole block) and a **Word Wrap** toggle.
- 🌗 **Clean light theme**, thoughtfully designed and responsive.
- 🔒 **Private by design** — 100% local, static HTML/CSS/JS. No trackers, no network calls, no build step.

> Your notes stay yours. The editor runs entirely inside Standard Notes' sandbox.

---

## 🚀 Install in 30 seconds

Standard Notes loads editors from a URL, so the fastest way is to use the hosted version.

1. Open **Standard Notes** → **Preferences / Settings → Plugins (or Extensions) → Import Extension**.
2. Paste this link and confirm:

   ```
   https://YOUR_GITHUB_USERNAME.github.io/richnote/ext.json
   ```

3. Open any note → click the **editor switcher** → choose **RichNote**. Done! 🎉

> Custom editors may require a Standard Notes plan that supports extensions, or the self‑hosted / desktop app.

### Prefer to host it yourself? (free, no server needed)

You don't need your own server — **GitHub Pages** hosts static files over HTTPS for free:

1. **Fork / create** a repo, e.g. `richnote`, and push these files.
2. In the repo: **Settings → Pages → Build from branch → `main` / root → Save.**
3. Your files are now live at `https://<username>.github.io/richnote/`.
4. Edit [`ext.json`](ext.json) — set `url` to `…/index.html` and `latest_url` to `…/ext.json` (replace `YOUR_GITHUB_USERNAME`).
5. Import the `ext.json` URL in Standard Notes (step 2 above).

> ⚠️ Standard Notes runs over **HTTPS**, so the editor must be HTTPS too. GitHub Pages is HTTPS out of the box — `http://localhost` will be blocked by the web app (use the desktop app for local testing).

> **App icon:** the editor ships a custom icon via `dock_icon` (SVG) in [`ext.json`](ext.json). If your Standard Notes version doesn't render the SVG, replace the `dock_icon` block with the native circle style:
> ```json
> "dock_icon": { "type": "circle", "background_color": "#2f6fdb", "foreground_color": "#ffffff", "border_color": "#2f6fdb" }
> ```

---

## ⌨️ Keyboard shortcuts

| Shortcut | Action | Shortcut | Action |
|---|---|---|---|
| `Ctrl+Z` / `Ctrl+Y` | Undo / Redo | `Ctrl+B / I / U` | Bold / Italic / Underline |
| `Ctrl+Shift+X` | Strikethrough | `Ctrl+Alt+1…6` | Heading 1–6 |
| `Ctrl+Alt+0` | Normal text | `Ctrl+L / E / R` | Align left / center / right |
| `Ctrl+Shift+8 / 7` | Bulleted / Numbered list | `Ctrl+K` | Insert link |
| `Tab` / `Shift+Tab` | Indent / Outdent | `Ctrl+\` | Clear formatting |
| `Enter` / `Shift+Enter` | New line (new block) | `F4` | **Repeat last format** |

`Ctrl` = `Cmd` on macOS.

---

## 💛 Support the developer

RichNote is **free and open‑source**. If it makes your notes nicer, please consider buying me a coffee — it directly funds new features and fixes. 🙏

<div align="center">

<img src="Images/QR%20Code%20bank%20donate.png" alt="Bank donate QR code" width="240" />

**Scan to donate** · Nguyễn Hữu Đức

</div>

You can also support by ⭐ **starring the repo** and sharing it with fellow Standard Notes users!

---

## 👤 Author

**Nguyễn Hữu Đức** — Software Developer @ **VIETIS**

- 📧 Email: [nguyenhuuduc.it.91@gmail.com](mailto:nguyenhuuduc.it.91@gmail.com)
- 📱 Phone/Zalo: **0964 589 910**

Found a bug or have an idea? [Open an issue](https://github.com/YOUR_GITHUB_USERNAME/richnote/issues) or reach out — feedback is always welcome.

---

## 🗂️ Project structure

```
richnote/
├── ext.json                  # Standard Notes component manifest
├── index.html                # Editor page (menu, toolbar, WYSIWYG area)
├── styles/editor.css         # All styling (light theme, toolbar, dialogs)
├── js/editor.js              # Editor logic + ComponentRelay integration
├── vendor/component-relay.js # Standard Notes bridge library
├── Images/                   # Screenshot + donate QR
└── LICENSE                   # MIT
```

## 🛠️ Tech

Plain **HTML / CSS / JavaScript** — no framework, no build. Formatting is powered by the browser's `contenteditable` and the official [`@standardnotes/component-relay`](https://www.npmjs.com/package/@standardnotes/component-relay) bridge for reading/saving the note.

Run locally (desktop app):

```bash
npm start           # serves at http://localhost:8080
# or: python3 -m http.server 8080
```

---

<div align="center">

Made with ♥ in Vietnam · MIT License · © 2026 Nguyễn Hữu Đức

**If this saved you time, a ⭐ and a ☕ go a long way!**

</div>
