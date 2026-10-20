/**
 * electron-builder configuration (moved out of package.json "build" so the
 * update feed can be injected at build time instead of living in the repo).
 *
 * Update feed (two modes):
 * - ARAOFFICE_UPDATE_URL set  -> generic provider, URL prefix serving
 *   latest.yml / latest-mac.yml (CDN). CI provides it as a secret; local
 *   builds use apps/shell/electron-builder.env (gitignored, auto-loaded by
 *   the CLI — keep its ARA_GITHUB_* in sync with the root .env private
 *   section; that file is the single source of truth).
 * - ARAOFFICE_UPDATE_URL unset -> GitHub provider from ARA_GITHUB_OWNER /
 *   ARA_GITHUB_REPO (same vars the app reads at runtime for the changelog
 *   + manual-download + issues links; private .env, gitignored). Empty =
 *   no publish config: electron-builder bakes no app-update.yml and in-app
 *   auto-update stays disabled (fork-safe open-source default).
 */

const { existsSync, mkdirSync, readFileSync, writeFileSync } = require('node:fs')
const { dirname, join } = require('node:path')

const updateUrl = process.env.ARAOFFICE_UPDATE_URL || sourceEnvValue('ARAOFFICE_UPDATE_URL')
const githubOwner = process.env.ARA_GITHUB_OWNER || sourceEnvValue('ARA_GITHUB_OWNER')
const githubRepo = process.env.ARA_GITHUB_REPO || sourceEnvValue('ARA_GITHUB_REPO')
/** One value from the repo-root private .env (build-time identity fallback); undefined when missing. */
function sourceEnvValue(key) {
  const source = join(__dirname, '../../.env')
  if (!existsSync(source)) return undefined
  for (const line of readFileSync(source, 'utf8').split(/\r?\n/)) {
    if (!line.startsWith(`${key}=`)) continue
    const value = line.slice(key.length + 1).trim()
    const quote = value[0]
    if ((quote === '"' || quote === "'") && value.endsWith(quote) && value.length >= 2) {
      return value.slice(1, -1)
    }
    return value
  }
  return undefined
}

// Vendor metadata resolved from build env (electron-builder.env / CI secrets) with a
// repo-root .env fallback; generic defaults keep fork builds buildable without private values.
const buildVendor =
  process.env.ARA_BRAND_VENDOR || sourceEnvValue('ARA_BRAND_VENDOR') || 'Adhi Rahmadian'
const copyright = `Copyright © ${new Date().getFullYear()} ${buildVendor}. All rights reserved.`
const legalTrademarks =
  process.env.ARA_BRAND_TRADEMARKS ||
  sourceEnvValue('ARA_BRAND_TRADEMARKS') ||
  'ARAOffice original work, all rights reserved'

// Runtime .env for installers: ship everything (brand, provider endpoints)
// EXCEPT credential-class keys — an API key/token/secret/cookie or free-form
// headers value never rides in a distributed build (they are gitignored on the
// build machine only). ARAOFFICE_SHIP_PRIVATE_ENV=1 bundles the source file
// verbatim for private builds when a credential must ride along. Never rename —
// these keys are resolved in apps/shell/src/main/service-config.ts.
const SECRET_ENV_RE = /(API_?KEY|TOKEN|SECRET|PASSWORD|COOKIE|HEADERS)/i

/** Writes the filtered runtime .env; returns its path, or null when no source .env exists. */
function buildPackagedEnv() {
  const source = join(__dirname, '../../.env')
  if (!existsSync(source)) return null
  if (process.env.ARAOFFICE_SHIP_PRIVATE_ENV === '1') return source
  const kept = readFileSync(source, 'utf8')
    .split(/\r?\n/)
    .filter((line) => {
      const t = line.trim()
      if (!t || t.startsWith('#')) return false
      return !SECRET_ENV_RE.test(t.split('=')[0] ?? '')
    })
  const out = join(__dirname, 'build', 'packaged.env')
  mkdirSync(dirname(out), { recursive: true })
  writeFileSync(out, `${kept.join('\n')}\n`)
  return out
}

const packagedEnvPath = buildPackagedEnv()

// Google Surf is staged into a production-only resource tree in `beforePack`.
// LICENSES.chromium.html only exists after the Electron binary download —
// since Electron 42 that no longer happens during `npm ci` (the postinstall
// script was replaced by the lazy `install-electron` bin), and electron-builder
// exits 0 on a missing extraResources source, so without this check the
// installer would silently ship without the Chromium license.
for (const rel of ['../../node_modules/electron/dist/LICENSES.chromium.html']) {
  if (!existsSync(join(__dirname, rel))) {
    throw new Error(
      `electron-builder extraResources source missing: ${rel} (npm hoisting changed?)`,
    )
  }
}

// The module trees are electron-vite outputs produced by build:all; a missing
// one means that module's build did not run or failed. electron-builder only
// logs "file source doesn't exist" for an absent extraResources source and
// still exits 0, so without this the installer launches normally and is simply
// missing that editor — it surfaces only when a user opens the tab.
//
// Runs from the beforePack hook, not at module load: gen-third-party-notices
// requires this config to read extraResources, and the dist:* scripts run
// notices before build:all, when the out dirs legitimately don't exist yet.
function assertModuleTreesPresent() {
  for (const rel of [
    '../docs/out',
    '../sheets/out',
    '../slides/out',
    '../pdf/out',
    '../markdown/out',
    '../diagram/out',
  ]) {
    if (!existsSync(join(__dirname, rel))) {
      throw new Error(
        `electron-builder extraResources source missing: ${rel} (run npm run build:all first)`,
      )
    }
  }
}

/** @type {import('electron-builder').Configuration} */
const config = {
  appId: 'com.araoffice.app',
  productName: 'ARAOffice',
  copyright,
  // Resolved from the installed electron package so dependency bumps can
  // never leave a stale hard-coded pin behind (packaging would silently ship
  // the old runtime).
  electronVersion: require('electron/package.json').version,
  directories: {
    output: 'release',
  },
  files: ['out/**'],
  extraResources: [
    {
      from: 'build/THIRD-PARTY-NOTICES.txt',
      to: 'THIRD-PARTY-NOTICES.txt',
    },
    {
      from: '../../node_modules/electron/dist/LICENSES.chromium.html',
      to: 'LICENSES.chromium.html',
    },
    {
      from: '../docs/out',
      to: 'modules/docs',
    },
    {
      from: '../sheets/out',
      to: 'modules/sheets',
    },
    {
      from: '../slides/out',
      to: 'modules/slides',
    },
    {
      from: '../pdf/out',
      to: 'modules/pdf',
    },
    {
      from: '../markdown/out',
      to: 'modules/markdown',
    },
    {
      from: '../diagram/out',
      to: 'modules/diagram',
    },
    // Runtime .env: added below only when a local .env exists (gitignored) —
    // public-key filtered unless ARAOFFICE_SHIP_PRIVATE_ENV=1.
    {
      from: '../../public/ARA_dark.png',
      to: 'ARA_dark.png',
    },
    {
      from: '../../public/ARA_light.png',
      to: 'ARA_light.png',
    },
    {
      from: '../../public/ARAOffice_dark.png',
      to: 'ARAOffice_dark.png',
    },
    {
      from: '../../public/ARAOffice_light.png',
      to: 'ARAOffice_light.png',
    },
    {
      from: 'build/docx.ico',
      to: 'docx.ico',
    },
    {
      from: 'build/xlsx.ico',
      to: 'xlsx.ico',
    },
    {
      from: 'build/pptx.ico',
      to: 'pptx.ico',
    },
    {
      from: 'build/pdf.ico',
      to: 'pdf.ico',
    },
    {
      from: 'build/md.ico',
      to: 'md.ico',
    },
    {
      from: 'build/diagram.ico',
      to: 'diagram.ico',
    },
    {
      from: 'build/diagram-code.ico',
      to: 'diagram-code.ico',
    },
  ],
  // `mimeType` is read only by the Linux target, where it becomes the
  // desktop entry's MimeType= list; associations without it are dropped
  // there. macOS and Windows ignore the field and key off `ext`.
  fileAssociations: [
    {
      ext: ['docx', 'doc'],
      name: 'ARAOffice.docx',
      description: 'ARA Docs',
      icon: 'build/docx.ico',
      role: 'Editor',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    },
    {
      ext: ['xlsx', 'xls', 'csv'],
      name: 'ARAOffice.xlsx',
      description: 'ARA Sheets',
      icon: 'build/xlsx.ico',
      role: 'Editor',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    },
    {
      ext: ['pptx', 'ppt'],
      name: 'ARAOffice.pptx',
      description: 'ARA Slides',
      icon: 'build/pptx.ico',
      role: 'Editor',
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    },
    {
      ext: 'pdf',
      name: 'ARAOffice.pdf',
      description: 'ARA PDF',
      icon: 'build/pdf.ico',
      role: 'Editor',
      mimeType: 'application/pdf',
    },
    {
      ext: ['md', 'markdown'],
      name: 'ARAOffice.md',
      description: 'ARA Notes',
      icon: 'build/md.ico',
      role: 'Editor',
      mimeType: 'text/markdown',
    },
    {
      ext: ['drawio', 'dio'],
      name: 'ARAOffice.drawio',
      description: 'ARA Diagrams',
      icon: 'build/diagram.ico',
      role: 'Editor',
      mimeType: 'application/xml',
    },
    {
      ext: ['mmd', 'mermaid'],
      name: 'ARAOffice.mmd',
      description: 'ARA Diagram Code',
      icon: 'build/diagram-code.ico',
      role: 'Editor',
      mimeType: 'text/plain',
    },
  ],
  npmRebuild: false,
  mac: {
    icon: '../../public/ARA_dark.png',
    target: ['dmg', 'zip'],
    category: 'public.app-category.productivity',
    hardenedRuntime: true,
    gatekeeperAssess: false,
    entitlements: 'build/entitlements.mac.plist',
    entitlementsInherit: 'build/entitlements.mac.plist',
    notarize: false,
    extraResources: [
      {
        from: '../sheets/native/xlsx-engine/target/release/xlsx-sidecar',
        to: 'native/xlsx-sidecar',
      },
    ],
  },
  win: {
    icon: 'build/icon.ico',
    legalTrademarks,
    target: [
      {
        target: 'nsis',
        arch: ['x64'],
      },
    ],
    extraResources: [
      {
        from: existsSync(
          join(__dirname, '../sheets/native/xlsx-engine/target/release/xlsx-sidecar.exe'),
        )
          ? '../sheets/native/xlsx-engine/target/release/xlsx-sidecar.exe'
          : '../sheets/native/xlsx-engine/target/x86_64-pc-windows-msvc/release/xlsx-sidecar.exe',
        to: 'native/xlsx-sidecar.exe',
      },
    ],
  },
  // Unlike win (which cross-compiles the sidecar to an explicit target
  // triple), linux takes it from cargo's host-native target/release/ — the
  // same source mac uses. So no `arch` is pinned here: electron-builder
  // defaults to the build host's architecture, which is the only one the
  // sidecar was actually built for. Packaging arm64 on an x64 host, or the
  // reverse, needs a matching `cargo build --target` first.
  linux: {
    // AppImage only: it needs no packaging identity, whereas deb/rpm would
    // require a Debian maintainer and homepage in the repo metadata.
    target: ['AppImage'],
    category: 'Office',
    icon: '../../public/ARA_dark.png',
    // mac and win name the binary from productName; linux instead derives it
    // from package.json "name", and "@araoffice/shell" sanitizes to the
    // invalid "@araofficeshell". Setting it explicitly also makes the
    // generated araoffice.desktop match the WM_CLASS Electron reports (it
    // takes that from the executable basename), so the running window links
    // back to its launcher entry.
    executableName: 'araoffice',
    // Electron takes its X11 app_id from package.json "desktopName"
    // (araoffice.desktop); syncDesktopName makes electron-builder name the
    // .desktop file and its StartupWMClass from the same value. Without it
    // StartupWMClass falls back to productName ("ARAOffice"), which does not
    // match the "araoffice" WM_CLASS the window actually reports — and X11
    // compares case-sensitively, so the taskbar shows an unlinked window.
    syncDesktopName: true,
    extraResources: [
      {
        from: '../sheets/native/xlsx-engine/target/release/xlsx-sidecar',
        to: 'native/xlsx-sidecar',
      },
    ],
  },
  nsis: {
    oneClick: false,
    perMachine: false,
    allowToChangeInstallationDirectory: true,
    artifactName: 'ARAOffice.Setup.${version}.${ext}',
    include: 'build/installer.nsh',
    uninstallDisplayName: '${productName}',
    installerSidebar: 'build/installerSidebar.bmp',
    uninstallerSidebar: 'build/uninstallerSidebar.bmp',
    installerHeader: 'build/installerHeader.bmp',
    createDesktopShortcut: 'always',
    createStartMenuShortcut: true,
    shortcutName: 'ARAOffice',
  },
  beforePack: async () => {
    assertModuleTreesPresent()
  },
  dmg: {
    sign: false,
  },
}

// Runtime .env resource: only when the source file exists at all; buildPackagedEnv
// already handled the public-key filtering (or private verbatim opt-in).
if (packagedEnvPath) {
  config.extraResources.push({ from: packagedEnvPath, to: '.env' })
}

if (updateUrl) {
  config.publish = [
    {
      provider: 'generic',
      url: updateUrl.replace(/\/+$/, ''),
      channel: 'latest',
    },
  ]
} else if (githubOwner && githubRepo) {
  // Private-release default: serve latest.yml from GitHub Releases
  // (e.g. <owner>/<app>-release). Forks without these vars get no
  // publish config and ship with auto-update disabled.
  config.publish = [
    {
      provider: 'github',
      owner: githubOwner,
      repo: githubRepo,
      releaseType: 'release',
    },
  ]
}
// When neither is set (forks, PR smoke builds, plain local packaging)
// publish is omitted: electron-builder bakes no app-update.yml into the
// app and in-app auto-update stays disabled.

module.exports = config
