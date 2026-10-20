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

## [1.3.0] — 2026-08-15

- **FIX**: **Table Formatting & Border Preservation in Docs** — Fixed table borders vanishing, columns collapsing, and awkward line wrapping when translating or rewriting documents with AI by preserving original borders, column widths, styles, cell fills, and OOXML byte-patch anchors.
- **FIX**: **Document Typography & Spacing Preservation** — Resolved issues where headings and paragraphs lost their custom fonts, line spacings, margins, and alignments during multi-block AI edits.
- **FIX**: **Formula & Math Display in Table Cells** — Fixed mathematical formulas inside table cells degrading to broken plain text during AI edits, ensuring equation structures and inline formatting stay intact.
- **FIX**: **AI Process Execution Indicator** — Fixed the execution card prematurely displaying "Finished execution" while the AI was still streaming responses, and removed auto-collapsing to give users full manual review control.
- **FIX**: **High-Contrast Dark Mode in Ribbon** — Fixed low-contrast text on active alignment buttons in dark mode for clean WCAG AA readability.
- **NEW**: **Redesigned Table Design & Layout Ribbon** — Overhauled the Table Design and Table Layout ribbon tabs into a snapped, compact 56px layout with quick style galleries, color swatches, 2×2 border tools, and intuitive column/row controls.
- **NEW**: **AI Presentation Preferences & Survey Card** — Redesigned the interactive survey card in Slides with a clean badge header, refined typography, and full dark-theme consistency.
- **NEW**: **Presentation Speaker Notes Management** — Added AI capabilities to read, write, and manage slide speaker notes and presenter talking points in Slides.
- **NEW**: **Smarter Slide Layouts & Collision Auditing** — Introduced 1280×720 spatial coordinate bounds and automated collision detection to prevent text and shapes from overlapping.
- **NEW**: **Enriched AI Ribbon Tabs Across All Apps** — Added dedicated AI action groups and preset tools across Docs, Sheets, Slides, PDF, and Markdown.

---

## [1.2.2] — 2026-07-01

- **FIX**: **Idle Tab Blank Screen** — Fixed documents turning blank or black when left open or idle for a long time, ensuring tabs stay visible and responsive without having to switch tabs.
- **FIX**: **Page Gap Display in Docs** — Fixed dark visual artifacts appearing between page gaps in Docs, providing a clean and seamless layout in dark mode.
- **FIX**: **No Extra Blank Pages on PDF Export** — Fixed an issue where an unnecessary empty page was added to the end of Docs files when exporting or printing to PDF.
- **FIX**: **Spreadsheet Stability** — Improved background engine reliability in Sheets to prevent unexpected freezes or disconnects when working on complex workbooks.
- **FIX**: **Smoother Saving on Windows** — Resolved occasional file saving errors on Windows caused by temporary background system locks.
