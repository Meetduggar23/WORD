<p align="center">
  <img src="frontend/react/public/newlogo.png" width="98" alt="WORD logo" />
</p>

<h1 align="center">WORD</h1>

<p align="center">
  <strong>A modern word processor built for the way people actually write.</strong>
</p>

<p align="center">
  <a href="#what-is-word">What is WORD</a> &nbsp;&middot;&nbsp;
  <a href="#how-it-works">How it Works</a> &nbsp;&middot;&nbsp;
  <a href="#features">Features</a> &nbsp;&middot;&nbsp;
  <a href="#get-started">Get Started</a> &nbsp;&middot;&nbsp;
  <a href="#roadmap">Roadmap</a>
</p>

---

## What is WORD

WORD is a full-featured word processor designed for everyday writing — from letters and essays to reports and documentation. It gives you everything you'd expect from a professional document editor: rich text formatting, page-based layout, tables, headers and footers, and a familiar ribbon-based interface.

What sets WORD apart is its built-in AI intelligence. It doesn't just store your words — it understands them. WORD can proofread your writing, suggest improvements, analyze document structure, answer questions about your content, and even respond to voice commands. It works offline with on-device AI or connects to cloud services like OpenAI for more powerful capabilities.

Whether you're a student writing an essay, a professional drafting a report, or a developer documenting a project, WORD adapts to your workflow.

---

## How it Works

### Writing and Editing

WORD uses a page-based canvas that renders your document exactly as it will appear when printed. You write in a clean, distraction-free editor with a floating toolbar for quick formatting. The ribbon at the top provides access to all tools organized into tabs — File, Home, Insert, Draw, Design, Layout, References, Mailings, Developer, and View.

Everything you type is processed through a document engine that maintains your content's structure. This means WORD understands the difference between a heading, a paragraph, a list item, and a table — not just the raw text.

### AI Intelligence

WORD's AI runs in multiple modes:

- **On-device (offline):** A rule-based engine that works without an internet connection. It handles proofreading, formatting analysis, text simplification, and document health checks using local heuristics.
- **Cloud-powered:** Connect to OpenAI, Ollama, or any compatible API for advanced rewriting, translation, tone adjustment, and natural language commands.
- **Privacy-first:** A status indicator shows whether AI is running locally or in the cloud. You control which mode is active.

When you select text and choose an AI action — like "make this more formal" or "fix grammar" — WORD sends your text to the active AI provider and presents the suggestion as a card. You can replace the original, insert below, copy, or dismiss. Nothing changes without your approval.

### Command Palette

Press `Ctrl+K` to open the command palette. It fuzzy-searches through 50+ commands — bold, insert table, find and replace, export PDF, and more. You can also type a natural language request like "make this paragraph shorter" and WORD will route it to the appropriate AI action.

### Voice Control

WORD listens for voice commands through your browser's speech recognition. Say things like "insert a table with four columns," "go to the introduction heading," or "bold the selected text" to control the editor hands-free.

### Document Analysis

WORD continuously evaluates your document's health across six dimensions: structure, style, clarity, consistency, formatting, and length. It gives you a score from 0 to 100 and suggests one-click fixes. There's also a cleanup planner that previews structural changes before applying them, and a design inspector that checks typography consistency.

### Smart Features

- **Smart Paste** detects when you paste content from the web, a PDF, or a spreadsheet and suggests the best format.
- **Table Intelligence** analyzes tables for patterns, anomalies, and data insights.
- **Smart References** auto-renumber figures and tables and update in-text citations.
- **Document Test** runs an automated check for common writing problems.
- **Time Machine** saves snapshots of your document so you can restore previous versions and see exactly what changed.

### Privacy and Control

You decide where your data goes. WORD shows a Lock or Cloud badge in the status bar indicating the current AI mode. You can toggle any AI feature on or off per device, use a local Ollama server, or bring your own API key. No data leaves your machine unless you explicitly choose a cloud provider.

---

## Features

### Editing and Formatting

- Rich text editing with page-based layout
- Character and paragraph formatting
- 50+ keyboard shortcuts via the command palette
- Find and replace with regex support
- Tables with merge, split, and alignment
- Headers, footers, and page numbers
- Style system with heading hierarchy

### AI-Powered Writing

- **AI Inline Actions** — Rewrite, shorten, formalize, proofread, translate, simplify, expand, or change tone
- **Suggestion Cards** — Review AI suggestions before applying them
- **Ask Document** — Ask questions about your document with source citations
- **Semantic Search** — Search your document by meaning, not just keywords
- **Document Health** — Score and improve your writing across six quality metrics
- **Voice Commands** — Control the editor with your voice

### Developer Tools

- **Code Block Dialog** — Insert syntax-highlighted code in 16 languages
- **JSON Tools** — Validate, format, minify, and inspect JSON
- **Markdown Import/Export** — Convert between Markdown and WORD format
- **Word-Level Diff** — Visual comparison between any two versions of your text

### History and Focus

- **Time Machine** — Snapshot history with visual diff
- **Focus Mode** — Distraction-free writing with session timer and daily goals
- **Analytics** — Reading time, word count, paragraph count, and document structure

---

## Get Started

### Prerequisites

- **Node.js 18+** and npm
- A modern browser (Chrome, Edge, Firefox)

### Install and Run

```bash
git clone https://github.com/Meetduggar23/WORD.git
cd WORD/frontend/react

npm install

npm run dev
```

The app opens at `http://localhost:5173`.

### Commands

```bash
npm run dev          # Start dev server
npm run build        # Build for production
npm run lint         # Run linter
npm run test         # Run tests
```

---

## Roadmap

- [ ] DOCX / PDF import and export
- [ ] Real-time collaboration
- [ ] Plugin system
- [ ] Macro recording and playback
- [ ] Advanced rendering (math equations, footnotes)
- [ ] Cloud storage integration
- [ ] Mobile and tablet optimization

---

## Contributing

Contributions are welcome. Run `npm run lint` and `npm run test -- --run` before pushing.

---

## Made By Meet Duggar 