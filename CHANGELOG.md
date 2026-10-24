# Changelog

All notable changes to **ARAOffice** are documented in this file.

The release note format below (`###` headings, `- **NEW**:`, `- **FIX**:`, `**bold**`, and `` `code` ``) is parsed directly by the updater / Changelog modal when pasted into a GitHub Release body.

---

## [0.1.0] — 2026-10-20

Major release of **ARAOffice** — fully transitioning to a 100% offline, privacy-first, pure open-source desktop office suite with an integrated client-direct AI copilot (**ARA Assistant**).

### Core Platform & Privacy

- **NEW**: **100% Offline-First & Zero Cloud** — Completely eliminated cloud dependencies, user sign-in requirements, telemetry, and background tracking. The office suite launches instantly and operates entirely offline.
- **NEW**: **Client-Direct AI Provider Engine** — Direct client connection to **Google Gemini** (via Google AI Studio) and any **OpenAI-compatible** endpoint (including local **Ollama** and **LM Studio**). API keys are stored solely in local `ai-settings.json` on the user's machine and are never transmitted over telemetry.
- **NEW**: **ARA Assistant Persona & Copilot** — Unified AI persona across all six editors with multi-step tool execution, undo snapshot timelines, and context-aware assistance.

### Omnibar & Workspace Orchestration

- **NEW**: **Command Center Omnibar (`Ctrl+K` / `⌘K`)** — Global intent detection and drafting. Ask to outline a proposal, analyze numbers, build presentation slides, or design an architecture diagram, and ARA Assistant drafts the artifact directly in chat.
- **NEW**: **1-Click Editor Handover** — Seamlessly transfer agreed drafts into target editors (`Open in Docs & Write`, `Open in Sheets & Write`, `Open in Diagram & Write`) with automated Mermaid syntax detection and document creation.
- **NEW**: **Indonesian Handover Routing** — The Command Center understands Indonesian creation verbs and artifact nouns when routing handover cards.

### Document Editors & Toolkits

- **NEW**: **PDF Document Toolkit** — Split, combine, and reorder PDF pages; extract embedded images to disk; sanitize documents (remove metadata) with a keyboard-friendly dialog; batch-delete annotations; write nested outlines; export text.
- **NEW**: **PDF Comments & Accessibility** — Comments panel with annotation management, squiggly underline markup, automated accessibility report checker, and outline tree renaming.
- **NEW**: **Docs Typography & Pagination Overhaul** — Real font metrics and UAX#14 break rules drive pagination with Word-parity visual rendering, tracked changes, comments, and paragraph byte patching.
- **NEW**: **Markdown Local Images & Diagram JPEG Export** — ARA Notes resolves local image assets through one-time per-window tickets; ARA Diagrams gains JPEG export alongside PNG.
- **NEW**: **Diagram Dual-Engine Editor** — Full-featured editor supporting interactive Draw.io visual canvases and live Mermaid.js code rendering with automatic hierarchical layouts.
- **FIX**: **Docs Canvas Centering** — Smoothly recentres the document canvas upon AI dock toggling without layout shifts or visual jitter.
- **FIX**: **Fork-Safe Environment Defaults** — Service URLs, update feeds, and branding resolve safely from environment with safe offline defaults.

---