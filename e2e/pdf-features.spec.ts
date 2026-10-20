import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { test, expect } from '@playwright/test'
import { PDFDocument } from 'pdf-lib'
import { applyOutline } from '../apps/pdf/src/main/outline-writer'
import { closeAndSaveVideo, launchShell, screenshotPath, waitForPageWithUrl } from './helpers'

/**
 * Real-app smoke of the Phase 2 PDF features: everything that tests cannot cover, because none of
 * it has an Electron runtime. The outline round trip is the most valuable leg — the panel edit
 * goes through the real tab state, the real IPC grant, the real file rewrite and the reload —
 * none of which is reachable from a vitest sandbox.
 *
 * Deliberately out of scope: actions that pop a native folder/save dialog (combine, embedded
 * images, export text/pages) — a modal native dialog would wedge the driver; their wiring is
 * exercised only up to the click.
 */

test.describe('pdf phase 2 features', () => {
  test('outline editor round trip rewrites the file and re-renders the sidebar', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'araoffice-pdf-'))
    const path = join(dir, 'document.pdf')
    const document = await PDFDocument.create()
    document.addPage([612, 792])
    document.addPage([612, 792])
    applyOutline(document, [
      { title: 'Alpha', pageIndex: 0 },
      { title: 'Beta', pageIndex: 1 },
    ])
    await writeFile(path, await document.save())

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'audit-pdf-smoke',
      openFile: path,
    })
    const { app, page } = launched
    // Uncaught renderer exceptions are the smoke signal that matters: everything else either
    // works visibly or fails visibly in a dialog
    const crashes: string[] = []
    page.on('pageerror', (error) => crashes.push(String(error)))

    try {
      const pdf = await waitForPageWithUrl(app, 'pdf/out')
      await expect(pdf.locator('.ribbon')).toBeVisible()

      // Open the bookmarks sidebar and enter edit mode
      await pdf.locator('.rb-big[title="Outline"]').first().click()
      const pane = pdf.locator('.pdf-outline-pane')
      await expect(pane.locator('.pdf-outline-toolbar')).toBeVisible()
      await pane.locator('.pdf-outline-tool', { hasText: 'Edit bookmarks' }).click()
      await expect(pane.locator('.pdf-outline-row')).toHaveCount(2)

      // Rename "Beta" through the prompt: ✎ opens with the existing title prefilled, Enter commits
      await pane.locator('.pdf-outline-row').nth(1).locator('button[title="Rename"]').click()
      const prompt = pdf.locator('.pdf-modal .pdf-modal-input')
      await expect(prompt).toHaveValue('Beta')
      await prompt.fill('Beta Renamed')
      await prompt.press('Enter')
      await expect(pane.locator('.pdf-outline-row').nth(1)).toContainText('Beta Renamed')

      // Delete "Alpha": the row holds the title plus ↑ ↓ ✎ ✕ ＋, addressed by title so button
      // arithmetic can't drift; then apply with "Beta" left
      await pane.locator('.pdf-outline-row').first().locator('button[title="Delete"]').click()
      await expect(pane.locator('.pdf-outline-row')).toHaveCount(1)
      await pane
        .locator('.pdf-outline-toolbar .pdf-outline-tool', { hasText: 'Apply bookmarks' })
        .click()

      // Apply rewrote the file and reloaded the document: edit mode closed, the read view shows
      // the surviving bookmark, and there is exactly one left
      await expect(pane.locator('.pdf-outline-row')).toHaveCount(0)
      await expect(pane.locator('.pdf-outline-item')).toHaveCount(1)
      await expect(pane.locator('.pdf-outline-item', { hasText: 'Beta Renamed' })).toBeVisible()
      await pdf.screenshot({ path: screenshotPath('audit-pdf-outline-applied') })
    } finally {
      await closeAndSaveVideo(launched, 'audit-pdf-smoke')
    }
    // Cross-realm and pdf-lib mismatches show up here first
    expect(crashes).toEqual([])
  })

  test('split and sanitize dialogs open and close without touching the file', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'araoffice-pdf-'))
    const path = join(dir, 'document.pdf')
    const document = await PDFDocument.create()
    for (let i = 0; i < 3; i++) document.addPage([612, 792])
    document.setTitle('Smoke Title')
    await writeFile(path, await document.save())

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'audit-pdf-dialogs',
      openFile: path,
    })
    const { app, page } = launched
    const crashes: string[] = []
    page.on('pageerror', (error) => crashes.push(String(error)))

    try {
      const pdf = await waitForPageWithUrl(app, 'pdf/out')
      await expect(pdf.locator('.ribbon')).toBeVisible()

      // These actions live on the Organize ribbon, not the home tab the app opens with
      await pdf.locator('.ribbon-tab', { hasText: 'Organize' }).click()

      // Split: three modes, then Escape discards without writing anything. The button label is
      // deliberately short ('Split PDF'); the longer wording is the dialog title
      await pdf.locator('.rb-big[title="Split PDF"]').click()
      const split = pdf.locator('.pdf-modal', { hasText: 'Split PDF' })
      const choices = split.locator('.pdf-modal-choices .pdf-modal-choice')
      await expect(choices).toHaveCount(3)
      await pdf.keyboard.press('Escape')
      await expect(split).toHaveCount(0)

      // Sanitize: six categories, Escape again before any checkbox is touched
      await pdf.locator('.rb-big[title="Sanitize Document"]').click()
      const sanitize = pdf.locator('.pdf-modal', { hasText: 'Sanitize Document' })
      await expect(sanitize.locator('.pdf-modal-choices .pdf-modal-choice')).toHaveCount(6)
      await pdf.keyboard.press('Escape')
      await expect(sanitize).toHaveCount(0)

      // Export embedded images on a no-image document surfaces the guard through the toolbar's
      // save-error badge instead of a dialog
      await pdf.locator('.rb-big[title="Export Embedded Images"]').click()
      await expect(pdf.locator('.tb-save-error')).toBeVisible({ timeout: 15_000 })

      // The accessibility report sits on the home ribbon — switch back, since ribbon tabs
      // unmount their groups
      await pdf.locator('.ribbon-tab', { hasText: 'Home' }).click()
      await pdf.locator('.rb-big[title="Accessibility report"]').click()
      const a11y = pdf.locator('.pdf-a11y-rows')
      await expect(a11y.locator('.pdf-a11y-row')).toHaveCount(8)
      await expect(a11y.locator('.pdf-a11y-pass')).toHaveCount(1)
      await expect(a11y.locator('.pdf-a11y-fail')).toHaveCount(2)
      await expect(a11y.locator('.pdf-a11y-manual')).toHaveCount(5)
      await expect(a11y.locator('.pdf-a11y-row', { hasText: 'Document title' })).toBeVisible()
      await pdf.keyboard.press('Escape')
      await pdf.screenshot({ path: screenshotPath('audit-pdf-dialogs') })
    } finally {
      await closeAndSaveVideo(launched, 'audit-pdf-dialogs')
    }
    expect(crashes).toEqual([])
  })
})
