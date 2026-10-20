import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { AgentLoop, type AgentStreamCallbacks, type AgentTransport } from '@araoffice/agent-core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { createDocsSkill } from '../src/renderer/ai/docs-skill'
import { buildDocsContext, countWords } from '../src/renderer/ai/protocol'
import { executeTool } from '../src/renderer/ai/tools'

/**
 * End-to-end through the local stack: AgentLoop -> docs skill -> tools ->
 * real tiptap editor. The "model" is scripted, so these tests verify that
 * the capabilities behind typical requests (word count, heading colors,
 * rewrite, insert) actually work when the model picks the right tool.
 */

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const text = (t: string): JsonNode => ({ type: 'text', text: t })
const heading = (t: string, level = 1): JsonNode => ({
  type: 'docHeading',
  attrs: { docxIndex: null, level },
  content: [text(t)],
})
const para = (t: string): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: [text(t)],
})

/** Editors created during the current test; destroyed in afterEach so ProseMirror's
 * DOMObserver timers can't fire after the JSDOM environment is torn down. */
const liveEditors: Editor[] = []

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  liveEditors.push(editor)
  return editor
}

afterEach(() => {
  for (const editor of liveEditors.splice(0)) editor.destroy()
})

/** 0 h1 | 1 p | 2 h2 | 3 p */
const fixture = () => [
  heading('Chapter 1 Overview', 1),
  para('ARAOffice is an AI office suite.'),
  heading('Risk Notes', 2),
  para('This document is for reference only.'),
]

function scriptedTransport(script: Array<(cb: AgentStreamCallbacks) => void>): AgentTransport {
  let turn = 0
  return {
    stream(_request, cb) {
      const step = script[turn++]
      if (step) queueMicrotask(() => step(cb))
      return { cancel: () => queueMicrotask(() => cb.onDone()) }
    },
  }
}

const flush = () => new Promise((r) => setTimeout(r, 0))
const NUM_IDS = { bullet: null, ordered: null }

function makeLoop(editor: Editor, transport: AgentTransport, onDone: (text: string) => void) {
  return new AgentLoop({
    transport,
    skill: createDocsSkill({
      getEditor: () => editor,
      getNumIds: () => NUM_IDS,
    }),
    events: { onDone: (r) => onDone(r.text) },
  })
}

describe('word-count stats (answer-style requests)', () => {
  it('each turn context carries full-text stats matching the status bar so the model can quote them directly', async () => {
    const editor = createEditor(fixture())
    const context = buildDocsContext(editor)
    const expected = countWords(editor.state.doc.textContent)
    expect(context).toContain(`Full-text stats: words ${expected}`)
    expect(context).toContain('0|h1|Chapter 1 Overview')
    expect(context).toContain('3|p|This document is for reference only.')
  })

  it('get_document_context tool returns the same stats and does not modify the document', async () => {
    const editor = createEditor(fixture())
    const before = JSON.stringify(editor.getJSON())
    const exec = await executeTool(
      editor,
      { id: 't1', name: 'get_document_context', input: {} },
      NUM_IDS,
    )
    expect(exec.mutated).toBe(false)
    expect(exec.output).toContain(`words ${countWords(editor.state.doc.textContent)}`)
    expect(JSON.stringify(editor.getJSON())).toBe(before)
  })

  it('answer-only turn: the model replies with plain text and the document is unchanged', async () => {
    const editor = createEditor(fixture())
    const before = JSON.stringify(editor.getJSON())
    let final = ''
    const loop = makeLoop(
      editor,
      scriptedTransport([
        (cb) => {
          cb.onDelta('The document has 17 words.')
          cb.onDone()
        },
      ]),
      (t) => (final = t),
    )
    loop.run('How many words are there now')
    await flush()
    expect(final).toBe('The document has 17 words.')
    expect(JSON.stringify(editor.getJSON())).toBe(before)
    // The context made it into the user message sent to the model
    expect((loop.messages[0] as { text: string }).text).toContain('Full-text stats')
  })
})

describe('changing heading colors (formatting-command requests)', () => {
  it('after the model calls apply_commands, all headings turn red with the aiChanged highlight', async () => {
    const editor = createEditor(fixture())
    let final = ''
    const loop = makeLoop(
      editor,
      scriptedTransport([
        (cb) => {
          cb.onToolCall({
            id: 't1',
            name: 'apply_commands',
            input: {
              commands: [
                {
                  updateTextStyle: {
                    target: { nodeType: 'docHeading' },
                    style: { color: 'FF0000' },
                    fields: ['color'],
                  },
                },
              ],
            },
          })
          cb.onDone()
        },
        (cb) => {
          cb.onDelta('All headings are now red.')
          cb.onDone()
        },
      ]),
      (t) => (final = t),
    )
    loop.run('Turn all headings red')
    await flush()
    await flush()

    expect(final).toBe('All headings are now red.')
    for (const blockIndex of [0, 2]) {
      const block = editor.state.doc.child(blockIndex)
      expect(block.attrs.aiChanged).toBe(true)
      const mark = block.child(0).marks.find((m) => m.type.name === 'docTextStyle')
      expect(mark?.attrs.color).toBe('FF0000')
    }
    // Body text is unaffected
    expect(
      editor.state.doc
        .child(1)
        .child(0)
        .marks.some((m) => m.type.name === 'docTextStyle'),
    ).toBe(false)
    // The tool result was passed back to the model
    const toolMsg = loop.messages[2] as { role: 'tool'; results: Array<{ output: string }> }
    expect(toolMsg.results[0].output).toContain('已更新 2 个段落的文字样式')
  })

  it('invalid commands return an error result, the document is unchanged, and the loop continues', async () => {
    const editor = createEditor(fixture())
    const before = JSON.stringify(editor.getJSON())
    const loop = makeLoop(
      editor,
      scriptedTransport([
        (cb) => {
          cb.onToolCall({
            id: 't1',
            name: 'apply_commands',
            input: { commands: [{ updateTextStyle: { style: {}, fields: [] } }] },
          })
          cb.onDone()
        },
        (cb) => {
          cb.onDelta('Invalid command.')
          cb.onDone()
        },
      ]),
      () => {},
    )
    loop.run('Make random edits')
    await flush()
    await flush()
    const toolMsg = loop.messages[2] as { role: 'tool'; results: Array<{ isError?: boolean }> }
    expect(toolMsg.results[0].isError).toBe(true)
    expect(JSON.stringify(editor.getJSON())).toBe(before)
  })
})

describe('content read/write tools', () => {
  it('read_blocks returns the full restricted HTML', async () => {
    const editor = createEditor(fixture())
    const exec = await executeTool(
      editor,
      { id: 't', name: 'read_blocks', input: { startBlockIndex: 0, endBlockIndex: 1 } },
      NUM_IDS,
    )
    expect(exec.output).toContain('<h1>Chapter 1 Overview</h1>')
    expect(exec.output).toContain('<p>ARAOffice is an AI office suite.</p>')
  })

  it('read_blocks pages oversized content: offset continuation reassembles the full HTML', async () => {
    // one paragraph well past the 24k read cap
    const long = 'A'.repeat(30_000)
    const editor = createEditor([heading('Long chapter', 1), para(long)])
    const first = await executeTool(
      editor,
      { id: 't1', name: 'read_blocks', input: { startBlockIndex: 0, endBlockIndex: 1 } },
      NUM_IDS,
    )
    expect(first.isError).toBeUndefined()
    const match = first.output.match(/offset=(\d+)/)
    expect(first.output).toContain('truncated')
    expect(match).not.toBeNull()
    const offset = Number(match![1])
    const second = await executeTool(
      editor,
      {
        id: 't2',
        name: 'read_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 1, offset },
      },
      NUM_IDS,
    )
    expect(second.isError).toBeUndefined()
    expect(second.output).toContain('end of range')
    const stitched =
      first.output.slice(0, first.output.lastIndexOf('\n…(truncated')) +
      second.output.slice(0, second.output.lastIndexOf('\n(end of range'))
    expect(stitched).toContain('<h1>Long chapter</h1>')
    expect(stitched).toContain(long)
    // an offset beyond the content is an explicit error, not an empty read
    const beyond = await executeTool(
      editor,
      {
        id: 't3',
        name: 'read_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 1, offset: 10_000_000 },
      },
      NUM_IDS,
    )
    expect(beyond.isError).toBe(true)
  })

  it('replace_blocks rewrites the specified blocks', async () => {
    const editor = createEditor(fixture())
    const exec = await executeTool(
      editor,
      {
        id: 't',
        name: 'replace_blocks',
        input: { startBlockIndex: 1, endBlockIndex: 1, html: '<p>Intro rewritten.</p>' },
      },
      NUM_IDS,
    )
    expect(exec.mutated).toBe(true)
    const block = editor.state.doc.child(1)
    expect(block.textContent).toBe('Intro rewritten.')
    expect(block.attrs.aiChanged).toBe(true)
    expect(editor.state.doc.childCount).toBe(4)
  })
  it('replace_blocks preserves table formatting, borders, widths, and paragraph styles', async () => {
    const customBorders = {
      top: { style: 'single', szEighths: 8, color: 'FF0000' },
      bottom: { style: 'single', szEighths: 8, color: 'FF0000' },
      left: { style: 'single', szEighths: 8, color: 'FF0000' },
      right: { style: 'single', szEighths: 8, color: 'FF0000' },
      insideH: { style: 'single', szEighths: 8, color: 'FF0000' },
      insideV: { style: 'single', szEighths: 8, color: 'FF0000' },
    }
    const tableNode: JsonNode = {
      type: 'docTable',
      attrs: {
        docxIndex: 42,
        borders: customBorders,
        colWidthsPct: [40, 60],
        tblStyleId: 'CustomGrid',
      },
      content: [
        {
          type: 'docTableRow',
          content: [
            {
              type: 'docTableHeader',
              attrs: { colwidth: [200], fill: 'E0F2FE' },
              content: [{ type: 'docParagraph', content: [text('Kondisi')] }],
            },
            {
              type: 'docTableHeader',
              attrs: { colwidth: [300], fill: 'E0F2FE' },
              content: [{ type: 'docParagraph', content: [text('Statistik')] }],
            },
          ],
        },
        {
          type: 'docTableRow',
          content: [
            {
              type: 'docTableCell',
              attrs: { colwidth: [200] },
              content: [{ type: 'docParagraph', content: [text('n >= 30')] }],
            },
            {
              type: 'docTableCell',
              attrs: { colwidth: [300] },
              content: [{ type: 'docParagraph', content: [text('z = 1.96')] }],
            },
          ],
        },
      ],
    }
    const styledPara: JsonNode = {
      type: 'docParagraph',
      attrs: {
        docxIndex: 10,
        align: 'center',
        lineSpacing: 360,
        spaceBefore: 120,
      },
      content: [text('Original paragraph text')],
    }
    const editor = createEditor([styledPara, tableNode])

    // Replace both blocks via replace_blocks
    const exec = await executeTool(
      editor,
      {
        id: 't',
        name: 'replace_blocks',
        input: {
          startBlockIndex: 0,
          endBlockIndex: 1,
          html: '<p>Translated paragraph text</p><table><tr><th>Conditions</th><th>Statistics</th></tr><tr><td>n &gt;= 30</td><td>z = 2.58</td></tr></table>',
        },
      },
      NUM_IDS,
    )
    expect(exec.mutated).toBe(true)

    // Verify paragraph formatting is preserved
    const newPara = editor.state.doc.child(0)
    expect(newPara.textContent).toBe('Translated paragraph text')
    expect(newPara.attrs.align).toBe('center')
    expect(newPara.attrs.lineSpacing).toBe(360)
    expect(newPara.attrs.spaceBefore).toBe(120)
    expect(newPara.attrs.docxIndex).toBe(10)

    // Verify table formatting is preserved
    const newTable = editor.state.doc.child(1)
    expect(newTable.attrs.docxIndex).toBe(42)
    expect(newTable.attrs.borders).toEqual(customBorders)
    expect(newTable.attrs.colWidthsPct).toEqual([40, 60])
    expect(newTable.attrs.tblStyleId).toBe('CustomGrid')

    // Verify cell formatting is preserved
    const row0 = newTable.child(0)
    const cell0 = row0.child(0)
    expect(cell0.textContent).toBe('Conditions')
    expect(cell0.attrs.colwidth).toEqual([200])
    expect(cell0.attrs.fill).toBe('E0F2FE')
  })

  it('insert_content inserts after the specified block', async () => {
    const editor = createEditor(fixture())
    const exec = await executeTool(
      editor,
      {
        id: 't',
        name: 'insert_content',
        input: { html: '<h2>New Section</h2><p>New content.</p>', afterBlockIndex: 3 },
      },
      NUM_IDS,
    )
    expect(exec.mutated).toBe(true)
    expect(editor.state.doc.childCount).toBe(6)
    expect(editor.state.doc.child(4).textContent).toBe('New Section')
    expect(editor.state.doc.child(5).textContent).toBe('New content.')
  })

  it('an out-of-range end index is an error, not silently clamped', async () => {
    const editor = createEditor(fixture())
    const before = JSON.stringify(editor.getJSON())
    for (const name of ['read_blocks', 'replace_blocks']) {
      const exec = await executeTool(
        editor,
        { id: 't', name, input: { startBlockIndex: 2, endBlockIndex: 9, html: '<p>x</p>' } },
        NUM_IDS,
      )
      expect(exec.isError).toBe(true)
      expect(exec.output).toContain('4 blocks')
    }
    expect(JSON.stringify(editor.getJSON())).toBe(before)
  })
})

describe('external-edit guard (document freshness baseline)', () => {
  const read = (editor: Editor) =>
    executeTool(editor, { id: 'r', name: 'get_document_context', input: {} }, NUM_IDS)
  const replace = (editor: Editor) =>
    executeTool(
      editor,
      {
        id: 'w',
        name: 'replace_blocks',
        input: { startBlockIndex: 1, endBlockIndex: 1, html: '<p>rewritten</p>' },
      },
      NUM_IDS,
    )

  it('index-addressed writes fail after a user edit, and succeed again after a re-read', async () => {
    const editor = createEditor(fixture())
    await read(editor)
    // simulate a manual user edit between tool calls
    editor.view.dispatch(editor.state.tr.insertText('typed by user ', 2))
    const stale = await executeTool(
      editor,
      {
        id: 'w',
        name: 'apply_commands',
        input: { commands: [{ deleteBlocks: { target: { blockIndexes: [3] } } }] },
      },
      NUM_IDS,
    )
    expect(stale.isError).toBe(true)
    expect(stale.output).toContain('edited by the user')
    expect(editor.state.doc.childCount).toBe(4) // nothing deleted
    await executeTool(
      editor,
      { id: 'r2', name: 'read_blocks', input: { startBlockIndex: 0, endBlockIndex: 3 } },
      NUM_IDS,
    )
    const retry = await replace(editor)
    expect(retry.isError).toBeUndefined()
    expect(editor.state.doc.child(1).textContent).toBe('rewritten')
  })

  it("the AI's own consecutive writes do not trip the guard", async () => {
    const editor = createEditor(fixture())
    await read(editor)
    expect((await replace(editor)).isError).toBeUndefined()
    const second = await executeTool(
      editor,
      { id: 'w2', name: 'insert_content', input: { html: '<p>appendix</p>', afterBlockIndex: 3 } },
      NUM_IDS,
    )
    expect(second.isError).toBeUndefined()
    expect(editor.state.doc.childCount).toBe(5)
  })
})

describe('blank-document detection', () => {
  it('an image-only document is not blank: insert_content appends instead of wiping it', async () => {
    const editor = createEditor([
      {
        type: 'docProtected',
        attrs: { docxIndex: null, blockType: 'image', label: 'Image' },
      },
    ])
    const exec = await executeTool(
      editor,
      { id: 't', name: 'insert_content', input: { html: '<p>caption</p>', afterBlockIndex: 0 } },
      NUM_IDS,
    )
    expect(exec.isError).toBeUndefined()
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.child(0).attrs.blockType).toBe('image')
    expect(buildDocsContext(editor)).not.toContain('blank')
  })

  it('a single empty paragraph is still blank: insert_content replaces the template paragraph', async () => {
    const editor = createEditor([{ type: 'docParagraph', attrs: { docxIndex: null } }])
    expect(buildDocsContext(editor)).toContain('blank')
    const exec = await executeTool(
      editor,
      { id: 't', name: 'insert_content', input: { html: '<p>hello</p>' } },
      NUM_IDS,
    )
    expect(exec.isError).toBeUndefined()
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.state.doc.child(0).textContent).toBe('hello')
  })
})

describe('web_search backend failures', () => {
  it("method 'error' surfaces as a tool error instead of '(no results)'", async () => {
    const editor = createEditor(fixture())
    const w = window as unknown as { desktop?: unknown }
    const saved = w.desktop
    w.desktop = {
      webSearch: async () => ({
        results: [],
        method: 'error',
        error: 'Web search is temporarily unavailable.',
      }),
    }
    try {
      const exec = await executeTool(
        editor,
        { id: 't', name: 'web_search', input: { query: 'office suite' } },
        NUM_IDS,
      )
      expect(exec.isError).toBe(true)
      expect(exec.output).toContain('Web search is temporarily unavailable.')
      expect(exec.output).not.toContain('(no results)')
    } finally {
      w.desktop = saved
    }
  })
})
