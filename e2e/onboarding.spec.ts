import { test, expect } from '@playwright/test'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, screenshotPath } from './helpers'

test.describe('first-run onboarding', () => {
  test('fresh install walks all slides and persists the seen flag', async () => {
    const launched = await launchShell({ videoDir: 'onboarding-walkthrough' })
    const { page, userDataDir } = launched
    try {
      const overlay = page.locator('.onb-overlay')
      await expect(overlay).toBeVisible()
      // the welcome slide is a hero-style layout: its title lives in .s1-title
      await expect(page.locator('.onb-slide.active .s1-title')).toHaveText('Welcome to ARAOffice')

      // the theme control is a light/dark toggle (no system option on this screen)
      const themeToggle = page.locator('.onb-theme-toggle')
      const pressedOnInit = (await themeToggle.getAttribute('aria-pressed')) === 'true'
      await themeToggle.click()
      const pressedAfter = (await themeToggle.getAttribute('aria-pressed')) === 'true'
      expect(pressedAfter).toBe(!pressedOnInit)
      await expect(page.locator('html')).toHaveAttribute(
        'data-theme',
        pressedAfter ? 'dark' : 'light',
      )
      await page.screenshot({ path: screenshotPath('onboarding-slide-1') })

      await page.locator('.onb-next').click()
      await expect(page.locator('.onb-slide.active .onb-feature-grid')).toBeVisible()
      await page.screenshot({ path: screenshotPath('onboarding-slide-2') })

      await page.locator('.onb-next').click()
      await expect(page.locator('.onb-slide.active .onb-highlights-grid')).toBeVisible()
      await page.screenshot({ path: screenshotPath('onboarding-slide-3') })

      await page.locator('.onb-next').click()
      // slide 4 is the ARA Assistant showcase; the last Next starts the app
      await expect(page.locator('.onb-ai-showcase')).toBeVisible()
      await page.screenshot({ path: screenshotPath('onboarding-slide-4') })

      // the final tour action starts the app; account creation is not part of the tour
      await page.locator('.onb-next').click()
      await expect(overlay).toBeHidden()
      await expect(page.locator('.home-hero')).toBeVisible()
      await page.screenshot({ path: screenshotPath('onboarding-done-home') })

      await expect
        .poll(async () => {
          const raw = await readFile(join(userDataDir, 'app-settings.json'), 'utf8')
          return (JSON.parse(raw) as { onboardingSeen?: boolean }).onboardingSeen
        })
        .toBe(true)
    } finally {
      await closeAndSaveVideo(launched, 'onboarding-walkthrough')
    }

    // second launch with the same userData must not show the onboarding again
    const relaunch = await launchShell({
      userDataDir: launched.userDataDir,
      videoDir: 'onboarding-relaunch',
    })
    try {
      await expect(relaunch.page.locator('.onb-overlay')).toHaveCount(0)
      await expect(relaunch.page.locator('.home-hero')).toBeVisible()
      await relaunch.page.screenshot({ path: screenshotPath('onboarding-relaunch-no-overlay') })
    } finally {
      await closeAndSaveVideo(relaunch, 'onboarding-relaunch')
    }
  })

  test('dot-navigation to the last slide finishes the onboarding and persists the seen flag', async () => {
    const launched = await launchShell({ videoDir: 'onboarding-dots' })
    const { page, userDataDir } = launched
    try {
      const overlay = page.locator('.onb-overlay')
      await expect(overlay).toBeVisible()
      await page.locator('.onb-dot').nth(3).click()
      await page.locator('.onb-next').click()
      await expect(overlay).toBeHidden()
      await expect(page.locator('.home-hero')).toBeVisible()

      await expect
        .poll(async () => {
          const raw = await readFile(join(userDataDir, 'app-settings.json'), 'utf8')
          return (JSON.parse(raw) as { onboardingSeen?: boolean }).onboardingSeen
        })
        .toBe(true)
    } finally {
      await closeAndSaveVideo(launched, 'onboarding-dots')
    }
  })
})
