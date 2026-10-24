# Changelog

All notable changes to **ARAOffice** are documented in this file.

The release note format below (`###` headings, `- **NEW**:`, `- **FIX**:`, `**bold**`, and `` `code` ``) is parsed directly by the updater / Changelog modal when pasted into a GitHub Release body.

---

## [0.1.0] — 2026-10-20

Initial release of **ARAOffice**, a local-first desktop office suite with six editors and an optional AI assistant.

### Unified Workspace

- **NEW**: **One workspace for six editors** — Open Docs, Sheets, Slides, PDF, Notes, and Diagrams in a shared desktop shell with tabs, a launcher, and recent files.
- **NEW**: **Command Center (`Ctrl+K` / `⌘K`)** — Find actions, describe a task, and draft content for supported editors from one search bar.
- **NEW**: **Cross-editor handoff** — Send supported AI drafts from the Command Center into an editor to continue working on them.

### Six Editors

- **NEW**: **Docs** — Create and edit DOCX documents with pagination, comments, tracked changes, and targeted paragraph saves.
- **NEW**: **Sheets** — Work with XLSX spreadsheets, formulas, charts, and pivot tables, supported by a native Rust calculation engine.
- **NEW**: **Slides** — Edit PPTX presentations with slide layouts, themes, speaker notes, and visual design tools.
- **NEW**: **PDF** — Annotate and fill documents, reorganize pages, merge or split files, and extract embedded images.
- **NEW**: **Notes** — Write Markdown with live preview, GitHub-flavored Markdown, an outline, and DOCX/HTML export.
- **NEW**: **Diagrams** — Create diagrams on a Draw.io canvas or write Mermaid diagrams with automatic layout.

### Optional ARA Assistant

- **NEW**: **In-editor AI assistance** — Ask for help with document text, spreadsheet formulas, presentations, PDFs, Markdown, and diagrams.
- **NEW**: **Choose your AI endpoint** — Connect Google Gemini or an OpenAI-compatible service, including local Ollama and LM Studio models. AI requests go directly from the app to the endpoint you configure.

### Local-First by Design

- **NEW**: **Work without an account** — Core editing is designed to work offline; ARAOffice has no sign-in or telemetry service.
- **NEW**: **Keep control of AI data** — AI is optional. When used, prompts and selected document context are sent to the provider you configured; external providers have their own data policies.
- **NEW**: **Desktop update support** — Windows and macOS builds can check for GitHub Releases updates. Linux builds provide a manual download link.
