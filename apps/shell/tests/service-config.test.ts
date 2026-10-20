import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

async function loadConfig() {
  return import('../src/main/service-config')
}

const ENV_KEYS = [
  'ARA_GITHUB_OWNER',
  'ARA_GITHUB_REPO',
  'ARA_WEBSITE_URL',
  'ARA_GITHUB_URL',
  'ARA_INSTAGRAM_URL',
  'ARA_BRAND_VENDOR',
] as const

beforeEach(() => {
  vi.resetModules()
  for (const key of ENV_KEYS) delete process.env[key]
})

afterEach(() => {
  for (const key of ENV_KEYS) delete process.env[key]
})

describe('service-config', () => {
  it('exposes the GitHub identity from env', async () => {
    process.env.ARA_GITHUB_OWNER = 'example-owner'
    process.env.ARA_GITHUB_REPO = 'example-repo'
    const { araGithubOwner, araGithubRepo } = await loadConfig()
    expect(araGithubOwner()).toBe('example-owner')
    expect(araGithubRepo()).toBe('example-repo')
  })

  it('exposes the public profile links and vendor from env', async () => {
    process.env.ARA_WEBSITE_URL = 'https://example.org'
    process.env.ARA_GITHUB_URL = 'https://github.example'
    process.env.ARA_INSTAGRAM_URL = 'https://instagram.example'
    process.env.ARA_BRAND_VENDOR = 'Example Vendor'
    const { araWebsiteUrl, araGithubUrl, araInstagramUrl, araBrandVendor } = await loadConfig()
    expect(araWebsiteUrl()).toBe('https://example.org')
    expect(araGithubUrl()).toBe('https://github.example')
    expect(araInstagramUrl()).toBe('https://instagram.example')
    expect(araBrandVendor()).toBe('Example Vendor')
  })

  it('returns default codebase configuration when unset', async () => {
    const {
      araGithubOwner,
      araGithubRepo,
      araWebsiteUrl,
      araGithubUrl,
      araInstagramUrl,
      araBrandVendor,
      ARA_CONFIG_DEFAULTS,
    } = await loadConfig()
    expect(araGithubOwner()).toBe(ARA_CONFIG_DEFAULTS.githubOwner)
    expect(araGithubRepo()).toBe(ARA_CONFIG_DEFAULTS.githubRepo)
    expect(araWebsiteUrl()).toBe(ARA_CONFIG_DEFAULTS.websiteUrl)
    expect(araGithubUrl()).toBe(ARA_CONFIG_DEFAULTS.githubUrl)
    expect(araInstagramUrl()).toBe(ARA_CONFIG_DEFAULTS.instagramUrl)
    expect(araBrandVendor()).toBe(ARA_CONFIG_DEFAULTS.vendor)
  })

  it('treats an empty value as not set', async () => {
    process.env.ARA_GITHUB_OWNER = ''
    const { araGithubOwner } = await loadConfig()
    expect(araGithubOwner()).toBe('')
  })
})
