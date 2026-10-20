import { describe, expect, it } from 'vitest'
import { parseBlocks } from '@araoffice/ui'

describe('parseBlocks markdown parser', () => {
  it('parses headings without infinite loop', () => {
    const blocks = parseBlocks('# Heading 1\n## Heading 2\n### Heading 3')
    expect(blocks).toEqual([
      { kind: 'h', level: 1, text: 'Heading 1' },
      { kind: 'h', level: 2, text: 'Heading 2' },
      { kind: 'h', level: 3, text: 'Heading 3' },
    ])
  })

  it('parses horizontal rules without infinite loop', () => {
    const blocks = parseBlocks('---\n***\n___')
    expect(blocks).toEqual([{ kind: 'hr' }, { kind: 'hr' }, { kind: 'hr' }])
  })

  it('parses hashtag without space as a paragraph without infinite loop', () => {
    const blocks = parseBlocks('#tag1 #tag2\n#123')
    expect(blocks).toEqual([{ kind: 'p', lines: ['#tag1 #tag2', '#123'] }])
  })

  it('parses table and adjacent paragraphs', () => {
    const input = `Before table
| Header 1 | Header 2 |
|---|---|
| Cell 1 | Cell 2 |
After table`
    const blocks = parseBlocks(input)
    expect(blocks).toEqual([
      { kind: 'p', lines: ['Before table'] },
      {
        kind: 'table',
        headers: ['Header 1', 'Header 2'],
        rows: [['Cell 1', 'Cell 2']],
      },
      { kind: 'p', lines: ['After table'] },
    ])
  })

  it('handles isolated pipe line without dropping next line', () => {
    const input = '| not a table |\nHello world'
    const blocks = parseBlocks(input)
    expect(blocks).toEqual([{ kind: 'p', lines: ['| not a table |', 'Hello world'] }])
  })

  it('handles empty and whitespace-only text', () => {
    expect(parseBlocks('')).toEqual([])
    expect(parseBlocks('\n\n\n')).toEqual([])
    expect(parseBlocks('   \n  \t  \n')).toEqual([])
  })

  it('parses mixed complex markdown document', () => {
    const input = `# Summary
This document contains #key points from page 1.

---

## Details
> Quoted note
> second line

- Bullet A
- Bullet B

1. First
2. Second`
    const blocks = parseBlocks(input)
    expect(blocks.length).toBe(7)
    expect(blocks[0]).toEqual({ kind: 'h', level: 1, text: 'Summary' })
    expect(blocks[1]).toEqual({
      kind: 'p',
      lines: ['This document contains #key points from page 1.'],
    })
    expect(blocks[2]).toEqual({ kind: 'hr' })
    expect(blocks[3]).toEqual({ kind: 'h', level: 2, text: 'Details' })
    expect(blocks[4]).toEqual({ kind: 'quote', lines: ['Quoted note', 'second line'] })
    expect(blocks[5]).toEqual({ kind: 'ul', items: ['Bullet A', 'Bullet B'] })
    expect(blocks[6]).toEqual({ kind: 'ol', start: 1, items: ['First', 'Second'] })
  })

  it('splits trailing question glued to last ordered item', () => {
    const blocks = parseBlocks(
      '7. Final proofread and prepare for sharing What type of document would you like to draft?',
    )
    expect(blocks).toEqual([
      { kind: 'ol', start: 7, items: ['Final proofread and prepare for sharing'] },
      { kind: 'p', lines: ['What type of document would you like to draft?'] },
    ])
  })

  it('splits trailing question glued to last bullet', () => {
    const blocks = parseBlocks('- Alpha release\n- Beta release When do you want to ship?')
    expect(blocks).toEqual([
      { kind: 'ul', items: ['Alpha release', 'Beta release'] },
      { kind: 'p', lines: ['When do you want to ship?'] },
    ])
  })

  it('splits trailing CTA glued to last bullet', () => {
    const blocks = parseBlocks(
      '- References — formatted in the target style (e.g., APA or IEEE) Click the action button below to open and build the full draft in ARA Docs.',
    )
    expect(blocks).toEqual([
      {
        kind: 'ul',
        items: ['References — formatted in the target style (e.g., APA or IEEE)'],
      },
      {
        kind: 'p',
        lines: ['Click the action button below to open and build the full draft in ARA Docs.'],
      },
    ])
  })

  it('splits glued bullet when intro exceeds 200 characters', () => {
    const longIntro =
      "Here's a ready-to-build outline for a Scopus-indexed journal article on AI and Human Literacy, structured in the standard IMRaD-style format used by most Scopus journals: Proposed article structure:"
    const input = `${longIntro} - Title & Abstract`
    const blocks = parseBlocks(input)
    expect(blocks).toEqual([
      { kind: 'p', lines: [longIntro] },
      { kind: 'ul', items: ['Title & Abstract'] },
    ])
  })

  it('keeps single-sentence items intact', () => {
    const blocks = parseBlocks('1. First item here\n2. Second item here')
    expect(blocks).toEqual([
      { kind: 'ol', start: 1, items: ['First item here', 'Second item here'] },
    ])
  })
})
