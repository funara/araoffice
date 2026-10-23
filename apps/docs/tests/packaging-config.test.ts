import packageJson from '../package.json'
import { describe, expect, it } from 'vitest'

describe('standalone packaging config', () => {
  it('does not bundle the repository .env file', () => {
    expect(JSON.stringify(packageJson.build)).not.toContain('../../.env')
  })
})
