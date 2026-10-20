import { test, expect } from '@playwright/test'
import { execSync } from 'node:child_process'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

/**
 * The first save of a blank workbook must persist its edits through Save As.
 */
test.describe('sheets: new blank workbook', () => {
  test('quick-create opens an editable grid and saves its first edit', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'sheets-new-blank' })
    try {
      const { app, page } = launched
      const scratch = await mkdtemp(join(tmpdir(), 'araoffice-sheets-blank-'))
      const workbook = join(scratch, 'first-edit.xlsx')
      await app.evaluate(({ dialog }, targetPath) => {
        dialog.showSaveDialog = async () => ({ canceled: false, filePath: targetPath })
      }, workbook)

      await expect(page.locator('.quick-card').nth(1)).toContainText('ARA Sheets')
      await page.locator('.quick-card').nth(1).click()

      const sheets = await waitForPageWithUrl(app, 'sheets/out')
      await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
        timeout: 30_000,
      })
      await sheets.waitForTimeout(1_500)
      await expect(sheets.locator('.app-shell')).not.toHaveClass(/copilot-collapsed/)

      const grid = await sheets.evaluate(() => {
        for (const canvas of document.querySelectorAll('canvas')) {
          const rect = canvas.getBoundingClientRect()
          if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
        }
        return null
      })
      if (!grid) throw new Error('worksheet canvas not found')
      await sheets.mouse.click(grid.x + 46 + 43, grid.y + 24 + 12)
      await sheets.keyboard.type('42', { delay: 50 })
      await sheets.keyboard.press('Enter')
      await sheets.screenshot({ path: screenshotPath('sheets-new-blank-edited') })

      await app.evaluate(({ webContents }) => {
        webContents
          .getAllWebContents()
          .find((contents) => contents.getURL().includes('sheets/out'))
          ?.send('menu:action', 'save')
      })
      await expect(() => {
        const xml = execSync(`unzip -p "${workbook}" xl/worksheets/sheet1.xml`).toString()
        expect(xml).toContain('<v>42</v>')
      }).toPass({ timeout: 15_000 })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-new-blank')
    }
  })
})
