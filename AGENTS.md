# Repository Guidelines

## Project Overview

ARAOffice is an npm-workspaces monorepo for an AI-native desktop office suite: six Electron editors (`docs`, `sheets`, `slides`, `pdf`, `markdown`, `diagram`) hosted by a unified `shell` (multi-tab windowing, launcher, Command Center omnibar `Ctrl+K`, recents, auto-updater).

**Core Invariant: File-Format Fidelity.** Opening/editing/saving DOCX, XLSX, PPTX, PDF must preserve untouched content, styling, metadata, and raw bytes. Never rebuild ZIP/OOXML from scratch or drop unknown XML; use narrow byte-anchored patches and minimal diffs.

**Core Invariant: Zero Cloud, Zero Accounts.** The app launches fully offline: no accounts, no sign-in, no telemetry, no heartbeat, no install ID, no remote feedback service. AI is optional and client-direct — the main process calls the model endpoint the user configured (Google AI Studio Gemini, or any OpenAI-compatible endpoint). Never add a service dependency, an auth flow, or a usage-reporting client back.

## Architecture & Data Flow

### Three-layer process separation

Every app (`apps/*/src/`) follows `main` / `preload` / `renderer`:

1. **Main (`src/main`):** Node-only. Window lifecycle, filesystem I/O, sidecars, AI bridge. No DOM.
2. **Preload (`src/preload`):** Context-isolated bridge via `contextBridge.exposeInMainWorld`. Wraps `ipcRenderer.invoke`/`on`, never exposes raw `ipcRenderer`.
3. **Renderer (`src/renderer`):** Sandboxed React 19 + Vite. Consumes `window.desktop*` / `window.aiOffice*` only. Enforce `contextIsolation: true, nodeIntegration: false, sandbox: true`.

### Typed IPC contract

1. Declare channels in `src/shared/*.ts` (e.g. `apps/shell/src/shared/tabs-api.ts`, `apps/docs/src/shared/ipc.ts`).
2. Preload wraps invoke + subscription; subscriptions return `() => void` teardown for `useEffect` cleanup.
3. Main registers `ipcMain.handle` with WebContents/origin + path validation; returns `{ ok: boolean; data?: T; error?: string }`, never throws across IPC.
4. Renderer calls typed `window.*` API or action hooks.

Shell domains: `tabs:*` (`tab-manager.ts`), `home:*` + `window:*` (`shell/src/main/index.ts`), `ai:*` (stream proxy, registered once by `docs-main.ts::registerAiIpc`).

### Engines and shell hosting

- **`docx-engine`:** OOXML → block tree (`p`/`tbl`/`drawing`, anchored by `docxIndex`); save splices OOXML fragments for dirty paragraphs only (`packages/docx-engine/src/text-patch.ts`).
- **`xlsx-engine` Rust sidecar:** `apps/sheets/native/xlsx-engine` (Rust edition 2024). `xlsx-sidecar-client.ts` speaks newline-delimited JSON over stdio for recalc, streaming parse, dependency graph.
- **`pptx-engine` + `pptx-render`:** Master → layout → slide inheritance; compiles EMU→px draw lists for Konva; untouched slide parts stay byte-identical.
- **`agent-core`:** ReAct loop (`src/loop.ts`), `composeSkills` (`src/skill.ts`), shared prompts/tools + `createIpcTransport` (`src/shared-skills.ts`).
- **`electron-utils`:** `safeExternalUrl`, `atomicWriteFile` + `renameWithRetry` (Windows `EBUSY`/`EPERM` backoff).
- **Shell:** Single `BrowserWindow` + per-app `WebContentsView` (`apps/shell/src/main/tab-manager.ts`, tab strip `y:0-40px`). Modals use `setOverlayVisible` snapshot trick because native views paint over DOM. Updater in `src/main/updater.ts`.

### AI skill wiring (unified)

Per app `apps/<app>/src/renderer/ai/prompts/system.md` (domain copy only) + `prompts/follow-ups.ts`, wired by `<app>-skill.ts` as `create<App>Skill` (id = folder name, tools `<APP>_TOOLS`, executor `execute<App>Tool(deps, call)`, context `build<App>Context`). Shared identity/policy/files/search live in `shared-skills.ts` (`buildAppSystemPrompt`, `CONTINUE_INSTRUCTION`, `FILES_PROMPT`, `SEARCH_PROMPT`). `ai/transport.ts` is a thin `createIpcTransport` wrapper. Tool names (`read_attachment`, `propose_operations`, ...) are model-facing — never rename. Omnibar has no tools; routes via `TabKind` + `dispatchAction` + `menu:command/ai-prompt`. Keep `.md` = copy, `.ts` = glue.
Ready-gate: each renderer calls `window.<api>.reportEditorReady()` once its real editor mounts → `ipcRenderer.send('editor:ready')` → `TabManager` delivers the create-doc handover payload (`insertText` + `menu:command`/`ai-prompt`) behind `EDITOR_READY_KINDS` (docs/markdown/sheets) with `EDITOR_READY_TIMEOUT_MS` (2500ms) fallback; kinds without an emitter resolve immediately (defined in `apps/shell/src/main/tab-manager.ts`).

## Key Directories

| Directory                                       | Purpose                                                           |
| ----------------------------------------------- | ----------------------------------------------------------------- |
| `apps/shell`                                    | Tab host, launcher, omnibar orchestrator, settings, updater       |
| `apps/docs`                                     | DOCX editor: TipTap/ProseMirror, pagination, paragraph patch save |
| `apps/sheets`                                   | Univer grid + Rust `xlsx-engine` sidecar, charts/pivots in-house  |
| `apps/pdf`                                      | PDF.js viewer + pdf-lib annotations/forms/pages                   |
| `apps/markdown`                                 | Live preview, GFM, outline, export pipelines                      |
| `apps/diagram`                                  | Draw.io canvas + live Mermaid renderer                            |
| `packages/docx-engine`                          | OOXML parser, block tree, byte patcher                            |
| `packages/pptx-engine` / `packages/pptx-render` | PPTX model/serializer, EMU→px draw lists                          |
| `packages/agent-core`                           | ReAct loop, skills, shared prompts, IPC transport                 |
| `packages/ai-provider`                          | Provider settings + direct Gemini / OpenAI-compatible clients     |
| `packages/ai-search`                            | Web/image search drivers                                          |
| `packages/file-parse`                           | Attachment text extraction                                        |
| `packages/project-store`                        | SQLite/JSON recents + conversations                               |
| `packages/electron-utils`                       | URL gates, atomic writes                                          |
| `packages/i18n`, `packages/ui`                  | Locales (en, zh-CN, ja, ko), shared React kit                     |
| `scripts/`, `tools/`                            | Fidelity runners, license/notice gates, formatters                |
| `e2e/`                                          | Serial Playwright + Electron specs                                |

## Development Commands

Run from repo root:

```bash
npm install                  # also lazy-downloads Electron via postinstall
npm run dev                  # 6 renderers (:5173-5178) + shell (:5199)
npm run dev:docs             # single-app renderer (same pattern per app)
npm run dev:docs-shell       # docs + shell subset
npm run build:all            # electron-vite build all apps into out/
npm run shell                # build:all + electron apps/shell
npm run fixtures             # generate docx fixtures (after docx-engine changes)
```

Quality gates:

```bash
npm run format               # Prettier --write, changed files only
npm run format:check         # Prettier --check (CI passes --base automatically)
npm run lint                 # eslint . (0 errors required)
npm run typecheck            # tsc --noEmit per workspace
npm test                     # vitest run per workspace + Rust sidecar tests
npm run test:e2e             # playwright vs built app (run build:all first)
npm run licenses             # allowlist gate from package-lock.json
```

Workspace-targeted:

```bash
npm test -w @araoffice/docs
npm run typecheck -w @araoffice/sheets
npm run build -w @araoffice/shell
npm run native:test -w @araoffice/sheets    # cargo test
npm run native:build -w @araoffice/sheets   # cargo build --release
npm run gate -w @araoffice/sheets           # typecheck + fixtures + test + compat
```

Packaging (uses the codebase GitHub Releases default; `ARAOFFICE_UPDATE_URL` optionally selects a generic update feed; unsigned without signing creds):

```bash
npm run dist:win
npm run dist:mac
npm run dist:linux
```

## Code Conventions & Common Patterns

- **Formatting:** Prettier `semi: false, singleQuote, trailingComma: all, printWidth: 100, endOfLine: lf`. Incremental only — format changed files via `tools/format-changed.mjs`, never whole-repo reformat.
- **Lint:** Flat `eslint.config.mjs`. `any` allowed at vendor/IPC boundaries; unused vars must be `_`-prefixed; empty `catch` allowed, other empty blocks flagged; `no-undef` off for TS.
- **Naming:** `kebab-case` files/dirs (`tab-manager.ts`, `atomic-write.ts`); `PascalCase` React components and types (`ExcelShell.tsx`, `DesktopApi`); tests `*.test.ts` under `tests/`, e2e `*.spec.ts` under `e2e/`.
- **Error handling:** Structured `{ ok, data?, error? }` across IPC. Validate paths/args/origins in main. Wrap layout surfaces in `<ErrorBoundary>`. Use `atomicWriteFile` for all writes (Windows AV/sync locks). English only in code/comments; user strings via i18n (`src/renderer/i18n/` + main-process dicts).
- **Async/events:** Preload subscriptions return unsubscribe; long AI/parse ops take `AbortController` keyed by `requestId`; main uses `node:fs/promises` or streams, never sync I/O in request paths.
- **DI/state:** `packages/*` stay pure — no Electron/Node imports except `electron-utils`; pass config bags. No Redux/Zustand/MobX; local `useState`/`useRef`/action modules, document state lives in the editor engine (TipTap, Univer, Konva).
- Renderer-shared modules (`apps/*/src/shared/*.ts`) are bundled into the sandboxed renderer: never import `node:`/`electron` there; Node-only helpers live in `src/main/` (precedent: `normalizePathKey` in `apps/shell/src/main/tab-manager.ts`).
- **Security:** All external URLs through `safeExternalUrl` (`http:`/`https:`/`mailto:` for user-activated document links — a dead mailto hyperlink in a DOCX/PPTX is a fidelity break; `file:`, `javascript:`, custom schemes rejected). One gate: `packages/electron-utils/src/safe-external-url.ts` — no per-app `allowedProtocols` overrides, no preload regexes. AI layout scripts run via constrained AST interpreter (Acorn), never `eval`/`Function`. Treat doc content as `<untrusted_context>`. Keep new files small — split substantial new concerns into new modules.

## System Security Hardening

AI credentials: the API key lives only in the local `ai-settings.json` inside the app's userData folder and is never sent to the renderer — reads are redacted (`redactAiSettings`) and only `hasApiKey` crosses the bridge. Auto-update: electron-updater on win/mac; Linux and other platforms get an explicit manual-download notice with the GitHub releases link in the Changelog modal (`updater.ts::manualDownloadNotice`) instead of a silent no-op.

## Important Files

| File                                                                      | Role                                                                        |
| ------------------------------------------------------------------------- | --------------------------------------------------------------------------- |
| `package.json`                                                            | Workspaces, engines, root fan-out scripts                                   |
| `tsconfig.base.json`                                                      | Shared TS base (ES2022, bundler, strict, noEmit)                            |
| `eslint.config.mjs`, `.prettierrc.json`                                   | Lint + format                                                               |
| `vitest.config.ts`                                                        | Root aggregation (`apps/*/vitest.config.ts`, `packages/*/vitest.config.ts`) |
| `apps/shell/src/main/index.ts`                                            | Shell entry, window lifecycle, home/window IPC                              |
| `apps/shell/src/main/tab-manager.ts`                                      | WebContentsView host, bounds, overlay snapshots                             |
| `apps/shell/src/main/updater.ts`                                          | electron-updater channels                                                   |
| `apps/shell/electron-builder.cjs`                                         | Packaging, file associations, update feed                                   |
| `apps/docs/src/main/docs-main.ts`                                         | Docs file ops, paragraph patching                                           |
| `apps/sheets/src/main/xlsx-sidecar-client.ts`                             | stdio JSON client to Rust engine                                            |
| `apps/sheets/native/xlsx-engine/src/lib.rs`                               | Calc + streaming archive engine                                             |
| `packages/docx-engine/src/text-patch.ts`                                  | Paragraph byte patcher                                                      |
| `packages/agent-core/src/loop.ts`, `src/shared-skills.ts`                 | Agent loop, shared prompts/tools/transport                                  |
| `packages/electron-utils/src/atomic-write.ts`, `src/safe-external-url.ts` | Atomic writes, URL gate                                                     |
| `packages/ai-provider/src/providers.ts`                                   | ARA model settings, default resolution, streaming/usage clients             |
| `e2e/playwright.config.ts`                                                | Serial Electron E2E config                                                  |

## Runtime/Tooling Preferences

- **Runtime:** Node `>=22.12.0`, npm `>=10` (npm workspaces `apps/*` + `packages/*`, `package-lock.json` v3). Never Bun/pnpm/yarn.
- **Rust:** `cargo` on PATH required only for sheets sidecar; `npm run build -w @araoffice/sheets` compiles it. Windows `dist:win` needs MinGW target `x86_64-pc-windows-gnu` staged (see CONTRIBUTING.md).
- **Stack:** TypeScript strict, React 19, Vite 7/8 + electron-vite 5, Electron 43, electron-builder 26.
- **Ports (`npm run dev` wires `*_RENDERER_URL`):** docs 5173, sheets 5174, slides 5175, pdf 5176, markdown 5177, diagram 5178, shell 5199. Override via `*_DEV_PORT` / `*_RENDERER_URL`.
- **Env (all optional):** `ARAOFFICE_USER_DATA`, `ARAOFFICE_LANG`, `ARAOFFICE_FAKE_UPDATE`, `ARAOFFICE_UPDATE_URL`, `XLSX_SIDECAR_PATH` (+ `XLSX_OPEN_PATH`, `XLSX_DEBUG_PORT`).

## Testing & QA

- **Vitest 4 per workspace:** `npm test` chains `vitest run` across 17 workspaces (i18n → … → diagram; `packages/ui` typecheck-only, no tests). UI tests use `jsdom`; engines/main use `node`; Canvas stubbed (e.g. `apps/slides/tests/setup.ts`); Electron mocked via `vi.mock('electron')`. Single: `npm test -w @araoffice/docs`; watch: `test:watch` in most workspaces. No coverage gate.
- **Rust:** `cargo test` under `apps/sheets/native/xlsx-engine` (formula, streaming XML, serialization); runs as part of sheets `npm test`.
- **Playwright E2E:** `e2e/*.spec.ts` (13 files) launch the real built app serially (`workers: 1`, 90s timeout). Requires `npm run build:all` first. Local-only drivers in `scripts/drivers/` are gitignored, never CI.
- **Fidelity suites:** `scripts/docs-word-fidelity.mjs` (vs Word/LibreOffice), `scripts/pagination-baseline*.mjs` → `apps/docs/tests/pagination-corpus/`, `tools/fidelity-compare.mjs` (PPTX vs LibreOffice). Fixtures: never hand-edit; regen via `npm run fixtures` / `npm run fixtures -w @araoffice/sheets`.
- **Gate before commit:** `npm run format:check && npm run lint && npm run typecheck && npm test && npm run licenses && npm run spill`. Round-trip changes (open/save docx/xlsx/pptx) must add a byte-for-byte survival test.
