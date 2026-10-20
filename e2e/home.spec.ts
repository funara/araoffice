import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, screenshotPath } from './helpers'

test.describe('home screen', () => {
  test('shows hero, quick-create cards and tab bar', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'home-basics' })
    const { page } = launched
    try {
      await expect(page.locator('.home-hero')).toBeVisible()
      // one quick-create card for each editor
      await expect(page.locator('.quick-card')).toHaveCount(6)
      await expect(page.locator('.quick-card').first()).toContainText('ARA Docs')
      await expect(page.locator('.quick-card').nth(1)).toContainText('ARA Sheets')
      await expect(page.locator('.quick-card').nth(2)).toContainText('ARA Slides')
      await expect(page.locator('.quick-card').nth(3)).toContainText('ARA PDF')
      await expect(page.locator('.quick-card').nth(4)).toContainText('ARA Notes')
      await expect(page.locator('.quick-card').nth(5)).toContainText('ARA Diagrams')
      const diagramFilter = page.locator('.filter-pills .filter-pill', { hasText: 'Diagrams' })
      await diagramFilter.click()
      await expect(diagramFilter).toHaveClass(/active/)
      await expect(
        page.locator('.tab-bar .tab-item.tab-home:not(.tab-home-placeholder)'),
      ).toBeVisible()
      await page.screenshot({ path: screenshotPath('home-overview') })
    } finally {
      await closeAndSaveVideo(launched, 'home-basics')
    }
  })

  test('renders localized UI when ARAOFFICE_LANG=zh-CN', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      lang: 'zh-CN',
      videoDir: 'home-zh-cn',
    })
    const { page } = launched
    try {
      await expect(page.locator('.nav-item .nav-label').first()).toHaveText('最近')
      await page.screenshot({ path: screenshotPath('home-zh-cn') })
    } finally {
      await closeAndSaveVideo(launched, 'home-zh-cn')
    }
  })
})
