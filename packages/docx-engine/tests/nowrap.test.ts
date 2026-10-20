import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/**
 * w:noWrap (ECMA-376 §17.4.30): "the contents of the current table cell shall not be
 * wrapped to multiple lines". Word lays such a cell out on one line per paragraph and
 * lets it overflow its column instead of wrapping, so a cell holding a long identifier
 * stays one line tall. The flag has to survive a save, since the cell may be regenerated.
 */
const cell = (text: string, tcPr: string) =>
  `<w:tc><w:tcPr>${tcPr}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`

const table = (tcPr: string) =>
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  `<w:tr>${cell('a_very_long_identifier_that_would_wrap', tcPr)}${cell('short', '<w:tcW w:w="2000" w:type="dxa"/>')}</w:tr>` +
  '</w:tbl>'

describe('table cell w:noWrap', () => {
  it('records the flag on the cell', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: table('<w:tcW w:w="2000" w:type="dxa"/><w:noWrap/>') }),
    )
    const cells = parsed.blocks[0].table?.rows[0]
    expect(cells?.[0]?.noWrap).toBe(true)
    expect(cells?.[1]?.noWrap).toBeUndefined()
  })

  it('leaves a cell without the flag alone', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: table('<w:tcW w:w="2000" w:type="dxa"/>') }),
    )
    expect(parsed.blocks[0].table?.rows[0]?.[0]?.noWrap).toBeUndefined()
  })

  it('keeps the flag in the saved bytes', async () => {
    const source = await buildDocx({
      bodyXml: table('<w:tcW w:w="2000" w:type="dxa"/><w:noWrap/>'),
    })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(
      parsed,
      parsed.blocks.map((b) => ({ kind: 'original', docxIndex: b.docxIndex! })),
    )
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(xml).toContain('<w:noWrap/>')
  })
})
