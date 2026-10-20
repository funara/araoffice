import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { test, expect } from '@playwright/test'
import { PDFDocument } from 'pdf-lib'
import { closeAndSaveVideo, launchShell, screenshotPath, waitForPageWithUrl } from './helpers'

test.describe('editor AI dock', () => {
  test('Slides opens with ARA Assistant visible', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'audit-slides' })
    const { app, page } = launched
    try {
      await page.locator('.quick-card', { hasText: 'ARA Slides' }).click()
      const slides = await waitForPageWithUrl(app, 'slides/out')
      await expect(slides.locator('.stage-wrap')).toBeVisible()
      await expect(slides.locator('.ai-dock')).not.toHaveClass(/collapsed/)
      await slides.screenshot({ path: screenshotPath('audit-slides-ai-open') })
    } finally {
      await closeAndSaveVideo(launched, 'audit-slides')
    }
  })

  test('PDF opens with ARA Assistant visible', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'araoffice-pdf-'))
    const path = join(dir, 'document.pdf')
    const document = await PDFDocument.create()
    document.addPage([612, 792])
    await writeFile(path, await document.save())
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'audit-pdf',
      openFile: path,
    })
    const { app } = launched
    try {
      const pdf = await waitForPageWithUrl(app, 'pdf/out')
      await expect(pdf.locator('.ribbon')).toBeVisible()
      await expect(pdf.locator('.ai-dock')).not.toHaveClass(/collapsed/)
      await pdf.screenshot({ path: screenshotPath('audit-pdf-ai-open') })
    } finally {
      await closeAndSaveVideo(launched, 'audit-pdf')
    }
  })

  test('Mermaid opens with ARA Assistant visible', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'araoffice-diagram-'))
    const path = join(dir, 'flow.mmd')
    await writeFile(path, 'flowchart LR\n  Start --> Finish\n')
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'audit-diagram',
      openFile: path,
    })
    const { app } = launched
    try {
      const diagram = await waitForPageWithUrl(app, 'diagram/out')
      await expect(diagram.locator('.mermaid-svg-wrapper')).toBeVisible({ timeout: 30_000 })
      await expect(diagram.locator('.rb-group-caption', { hasText: 'Insert' })).toBeVisible()
      await expect(diagram.locator('.mermaid-preview-actions .mermaid-action-btn')).toHaveCount(1)
      await expect(diagram.locator('.mermaid-preview-actions .mermaid-action-btn')).toHaveText(
        'Copy',
      )
      await expect(diagram.locator('.ai-dock')).not.toHaveClass(/collapsed/)
      await diagram.screenshot({ path: screenshotPath('audit-diagram-ai-open') })
    } finally {
      await closeAndSaveVideo(launched, 'audit-diagram')
    }
  })

  test('Diagram tab icon shows D badge', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'audit-diagram-shapes' })
    const { app, page } = launched
    try {
      await page.locator('.quick-card', { hasText: 'ARA Diagram' }).click()
      await page.getByRole('button', { name: 'ARA Diagrams (.drawio)' }).click()
      await waitForPageWithUrl(app, 'diagram/out')
      await expect(page.locator('.tab-item.active .tab-icon text')).toHaveText('D')
    } finally {
      await closeAndSaveVideo(launched, 'audit-diagram-shapes')
    }
  })

  test('Mermaid preview renders SVG from code', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'araoffice-mmd-render-'))
    const path = join(dir, 'simple.mmd')
    await writeFile(
      path,
      `flowchart TD
  Start([Start]) --> Process[Step 1: Execute Task]
  Process --> Choice{Is Success?}
  Choice -- Yes --> Finish([Finish])
  Choice -- No --> Retry[Step 2: Retry]
  Retry --> Process`,
    )
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'audit-mermaid-render',
      openFile: path,
    })
    const { app } = launched
    try {
      const diagram = await waitForPageWithUrl(app, 'diagram/out')
      const preview = diagram.locator('.mermaid-svg-wrapper svg')
      await expect(preview).toBeVisible({ timeout: 15_000 })
      await expect(preview.getByText('Step 1: Execute Task')).toBeVisible()
      const occupiedWidth = await preview.evaluate((element) => {
        const svg = element as SVGSVGElement
        return svg.getBBox().width / svg.viewBox.baseVal.width
      })
      expect(occupiedWidth).toBeGreaterThan(0.5)
      await diagram.screenshot({ path: screenshotPath('audit-mermaid-render') })
    } finally {
      await closeAndSaveVideo(launched, 'audit-mermaid-render')
    }
  })
})
