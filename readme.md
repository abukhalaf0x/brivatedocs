# 📝 brivatedocs

**100% Client-Side, Free, Private, and Unlimited Document Suite for Word (`.docx`), PowerPoint (`.pptx`), and PDF (`.pdf`).**

🌐 **Live App:** [brivatedocs.vercel.app](https://brivatedocs.vercel.app/)

---

## 💡 Why This Project Exists

Most online document editors today lock basic tools behind paywalls, enforce daily file limits, require account sign-ups, or upload your sensitive documents to remote servers.

**brivatedocs** was built to solve that: a fast, lightweight, privacy-first document suite that runs **100% locally inside your browser**. Your files never leave your device, there are zero backend servers, zero accounts, and zero artificial limits.

## 🤝 Open Source & Transparency

* **Free for Everyone:** Completely open-source under the [MIT License](LICENSE). Feel free to use it, fork it, customize it, self-host it, or build upon it however you like.
* **Built with AI:** A large portion of this project was built and refined with the assistance of AI. I am **not** a full-time software developer, so please don't expect enterprise-grade perfection or zero bugs!
* **Feedback & Maintenance:** If you run into any bugs, edge cases, or have feature ideas, feel free to reach out anytime — I always do my best to reply and keep improving the project.

---

## ✨ What You Can Do

### 🔒 Privacy, Offline & Cross-Platform
* **100% Local Processing:** Parsing, rendering, editing, signing, and exporting happen entirely in browser memory.
* **Installable Offline PWA:** Install **brivatedocs** as a standalone desktop or mobile app that works even without an internet connection.
* **Auto-Save Drafts (`IndexedDB`):** Automatically saves up to 5 recent projects locally in your browser so you never lose progress.
* **Bilingual & Theme Ready:** Instant switching between **English (`LTR`)** and **Arabic (`RTL`)**, plus **Light Mode** and **Dark Mode**, optimized for both desktop and mobile screens.

### 📕 PDF Editor (`.pdf`)
* **Surgical Native Text Editing:** Click directly on original PDF text to edit or delete it in place without messy whiteout boxes.
* **Digital & Electronic Signatures:** Sign documents by drawing smoothly, typing in signature fonts, or uploading a signature/stamp image (with automatic white-background removal), local signature saving, multi-page placement, and an optional **`SHA-256` cryptographic verification seal**.
* **Markup & Page Tools:** Translucent highlighter, whiteout/redaction rectangles, freehand pen, draggable text boxes, and image/logo insertion.
* **Page Management & Search:** Merge multiple PDFs (with drag-and-drop ordering), insert blank pages, reorder, rotate, delete, extract specific page ranges, full-text PDF search, and high-DPI PDF / PNG export.

### 📄 Word Editor (`.docx`, `.docm`)
* **Rich WYSIWYG Editing:** Create from built-in templates (Blank, Resume/CV, Formal Letter, Academic Report, Invoice) or edit existing `.docx` files.
* **Complete Formatting:** Headings, fonts, colors, highlights, format painter, RTL/LTR paragraph direction, lists, interactive checklists (`☐`/`☑`), tables (with a smart sticky contextual toolbar), images (with `Ctrl+V` clipboard paste & resize), hyperlinks, footnotes, comments, auto Table of Contents, and Markdown import.
* **Page Setup & Print:** A4 / Letter sizes, Portrait / Landscape, custom margins, page colors, borders, watermarks, headers/footers with page numbering, and clean direct printing (`Ctrl+P`).

### 📊 PowerPoint Editor (`.pptx`, `.pptm`)
* **Full Slide & Canvas Control:** Add, duplicate, reorder, hide/show slides, switch 16:9 / 4:3 aspect ratios, and apply one-click global deck themes or gradient/image backgrounds.
* **Rich Slide Elements:** Text boxes, Smart Slides, SmartArt diagrams, SVG icon library, editable tables, multi-series charts (`Bar`, `Line`, `Pie`), shapes, images (including batch **Images → Slides** & `Ctrl+V` paste), and audio/video media.
* **Presenting & Exports:** Fullscreen slideshow with slide transitions, object entrance animations, autoplay timer, touch swipe navigation, laser pointer & live pen, Presenter View with speaker notes & timer, and `.pptx` / `PDF` / `PNG` / `ZIP` exports.

### 🔄 Cross-Format Conversions & Quick Utilities
* **PDF ↔ Word (`.docx`):** Convert PDFs into editable Word documents or send Word documents straight into the PDF editor.
* **Word ↔ PowerPoint (`.pptx`):** Turn structured Word headings into presentation slides or export slide text to Word.
* **Quick Home Tools:** Standalone one-click **Images → PDF**, **Split PDF**, **Compress PDF**, **Protect / Unlock PDF**, and **Fill PDF Form** tools (also available from the PDF editor toolbar).
* **Compress PDF:** Four levels (Light / Medium / Strong / Maximum) that re-compress embedded images on your device while keeping text sharp and metadata intact; *Maximum* turns pages into images for the smallest size.
* **Password Protect & Unlock:** AES-256 encryption and decryption done locally; password-protected PDFs ask for the password when opened in the editor.
* **Fill PDF Forms:** Fill text fields, checkboxes, dropdowns and radio buttons of ready-made forms (including Arabic text) and optionally flatten the result.

---

## 🛠️ Tech Stack & Architecture

Built with **zero build step**, **zero npm dependencies**, and **zero backend** — modular `'use strict'` ES6+ scripts under `js/` (~2.8 MB total including vendored engines):

* **Frontend Core:** Vanilla HTML5, CSS3 (logical RTL/LTR properties & CSS custom properties), and modular vanilla JavaScript (`js/`).
* **Office OOXML Engine (`.docx` & `.pptx`):** [JSZip](https://stuk.github.io/jszip/) + native browser `DOMParser` & `XMLSerializer` for direct Open Packaging Conventions (OPC) XML manipulation.
* **PDF Rendering & Search:** Mozilla [PDF.js](https://mozilla.github.io/pdf.js/) (with Web Worker) for fast canvas rendering, thumbnails, and text extraction.
* **PDF Manipulation & Signing:** [pdf-lib](https://pdf-lib.js.org/) + native `WebCrypto` (`SHA-256`) for structural edits, content-stream patching, annotations, and digital signature seals.
* **PDF Encryption:** [qpdf](https://github.com/qpdf/qpdf) compiled to WebAssembly ([qpdf-wasm](https://github.com/neslinesli93/qpdf-wasm), ISC license — see qpdf-wasm-LICENSE.txt), loaded only when you protect or unlock a PDF.
* **Client-Side Storage & Offline:** Native `IndexedDB`, `localStorage`, and Progressive Web App `Service Worker`.

---

## 🚀 Getting Started Locally

No Node.js build pipeline or package installation is required:

1. Clone or download this repository:
   ```bash
   git clone https://github.com/abukhalaf0x/brivatedocs.git
   cd brivatedocs
   ```
2. Serve the folder over any local HTTP server (so `pdf.worker.min.js` and the PWA service worker run cleanly):
   ```bash
   npx serve .
   # or
   python -m http.server 8000
   ```
3. Or deploy the static root directory directly to **Vercel**, **Netlify**, or **GitHub Pages**.

---

## 📬 Contact

Found an issue or want to share feedback? Reach out anytime via the contact email listed on the website footer or through [abukhalaf0x.vercel.app](https://abukhalaf0x.vercel.app).

---

## 🇸🇦 Made in The Kingdom of Saudi Arabia 

لَا إِلَهَ إِلَّا أَنْتَ سُبْحَانَكَ إِنِّي كُنْتُ مِنَ الظَّالِمِينَ