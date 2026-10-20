# ARAOffice — Design System & UI/UX Specification (`DESIGN.md`)

**Visual Direction**: **Clean Office Productivity + Dynamic AI Copilot Precision**  
**Surface Mode**: **Operate** (High-density desktop office suite where readability, document canvas fidelity, keyboard speed, and zero-distraction layout outrank decorative chrome).

---

## 1. Design Philosophy

ARAOffice harmonizes desktop office productivity with modern AI copilot assistance:

1. **Native Windowing & Multi-Tab Workspace (`apps/shell`)**:
   - Universal `40px` tab strip hosting per-document `WebContentsView` renderers.
   - Clean Home launcher dashboard with quick actions (New Doc, Spreadsheet, Presentation, PDF, Markdown, Diagram) and recent files carousel.
   - Centered Command Center Omnibar (`Ctrl+K` / `⌘K`) providing instant keyboard-driven search, multi-app intent detection, and 1-click artifact drafting.

2. **Consistent Document Editor Chrome**:
   - Every editor (`docs`, `sheets`, `slides`, `pdf`, `markdown`, `diagram`) shares common ribbon ergonomics, toolbar geometry, and high-contrast dark/light theme tokens.
   - Distinct, recognizable document brand accents:
     - **Docs**: Word Blue (`#2b579a`)
     - **Sheets**: Excel Green (`#107c41`)
     - **Slides**: PowerPoint Vermilion (`#c43e1c`)
     - **PDF**: Acrobat Crimson (`#b30b00`)
     - **Notes / Markdown**: Markdown Teal (`#0e707e`)
     - **Diagram**: Diagram Purple (`#7738a6`)

3. **Integrated AI Assistant Dock (`ARA Assistant`)**:
   - Collapsible 340px right sidebar with subtle 1px divider border.
   - Unified AI composer across all editors with multi-step tool timelines, interactive diff reviews, and 1-click undo.
   - Unconfigured states offer clear, guidance-first provider setup cards (Google AI Studio Gemini or OpenAI-compatible endpoints).

---

## 2. Color & Surface Token Architecture

All components across `@araoffice/ui`, shell, and individual editors reference standardized CSS custom properties from `:root` / `[data-theme="light"]` and `[data-theme="dark"]`.

### 2.1 Surface Hierarchy

| Token                               | Light Mode | Dark Mode | Usage                                                      |
| :---------------------------------- | :--------- | :-------- | :--------------------------------------------------------- |
| `--chrome` / `--chrome-bg`          | `#f1f3f6`  | `#16181a` | Tab strip (`40px`), status bar (`26px`), frame chrome      |
| `--canvas` / `--bg-content`         | `#fafafa`  | `#1a1a1a` | App background canvas, document backdrop                   |
| `--surface`                         | `#ffffff`  | `#23272e` | Primary editor panes, document cards, modal dialogs, menus |
| `--surface-subtle`                  | `#f6f7f9`  | `#1e2227` | Ribbon strips, secondary toolbars, side rails              |
| `--hover` / `--bg-hover`            | `#f5f5f5`  | `#2d3139` | Hover states on ribbon buttons, menu items, tab items      |
| `--active-bg` / `--bg-hover-strong` | `#eaeaea`  | `#353b45` | Pressed/active buttons, active tab indicators              |

### 2.2 Brand & Semantic Tokens

| Token                        | Light Mode | Dark Mode | Usage                                                          |
| :--------------------------- | :--------- | :-------- | :------------------------------------------------------------- |
| `--accent`                   | `#2367ec`  | `#4a9eff` | Primary action buttons, active tab border, focus rings (`2px`) |
| `--accent-hover`             | `#1b58cf`  | `#60a9ff` | Hover state for accent buttons                                 |
| `--text` / `--text-primary`  | `#242424`  | `#e6edf3` | Primary body text and headings                                 |
| `--text-muted`               | `#73777c`  | `#8b949e` | Secondary labels, timestamps, hotkey badges                    |
| `--border`                   | `#e3e6ea`  | `#30363d` | Pane borders, card separators, table rules                     |
| `--border-strong`            | `#c9cdd3`  | `#484f58` | Active input borders, modal outlines                           |
| `--color-error` / `--danger` | `#ef4444`  | `#f87171` | Error notices, invalid inputs, destructive dialogs             |

---

## 3. Typography & Numerical Discipline

- **Primary UI Font (`--font-sans`)**: `'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif`
  - Base body: `13px` (`line-height: 1.45`)
  - Compact toolbar controls / ribbon labels: `11px - 12px` (`font-weight: 500`)
  - Modal titles: `15px` (`font-weight: 600`)
- **Code & Markdown Monospace Font (`--font-mono`)**: `'Cascadia Code', 'Consolas', 'Fira Code', monospace`
  - Used in technical markdown codeblocks, formula editor bars, and Mermaid diagram source code.
- **Strict Tabular Numerals (`font-variant-numeric: tabular-nums`)**:
  - Enforced in spreadsheet cells, row/column indicators, pagination counters, and file-size labels.

---

## 4. Modal & Overlay Guidelines

Because Electron `WebContentsView` renders as native child views that paint directly over the web DOM, modals in `apps/shell` follow a strict snapshot architecture:

1. When a modal opens (e.g. AI Provider Setup, Account Dialog, Confirm Dialog), the `TabManager` takes an overlay snapshot of the active view (`setOverlayVisible(true)`).
2. The modal renders in the shell window on top of the snapshot backdrop.
3. Closing the modal re-enables the interactive native view (`setOverlayVisible(false)`).
