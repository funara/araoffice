import { describe, expect, it, vi, type Mock } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { PDF_TOOLS, executePdfTool, type PdfSkillDeps } from '../src/renderer/ai/tools'
import type { SearchIndex } from '../src/renderer/search'
import type { FormValueInput } from '../src/shared/ipc'

/** Two pages: "Hello World" and "foo bar foo" */
const INDEX: SearchIndex = [
  {
    text: 'Hello World',
    lower: 'hello world',
    items: [{ start: 0, end: 11, x: 0, y: 700, w: 110, h: 12 }],
  },
  {
    text: 'foo bar foo',
    lower: 'foo bar foo',
    items: [{ start: 0, end: 11, x: 0, y: 700, w: 110, h: 12 }],
  },
]

interface Widget {
  annotationType?: number
  subtype?: string
  fieldType?: string
  fieldName?: string
  fieldValue?: unknown
  buttonValue?: string
  readOnly?: boolean
  checkBox?: boolean
  radioButton?: boolean
  options?: { exportValue?: unknown; displayValue?: unknown }[]
}

function fakeDoc(pageTexts: string[], annotsByPage: Widget[][] = []): PDFDocumentProxy {
  return {
    numPages: pageTexts.length,
    getPage: async (n: number) => ({
      getTextContent: async () => ({ items: [{ str: pageTexts[n - 1], hasEOL: true }] }),
      getAnnotations: async () => annotsByPage[n - 1] ?? [],
      cleanup: () => {},
    }),
  } as unknown as PDFDocumentProxy
}

type MockedDeps = PdfSkillDeps & { addMarkup: Mock<PdfSkillDeps['addMarkup']> }

function makeDeps(over: Partial<PdfSkillDeps> = {}): MockedDeps {
  const deps: PdfSkillDeps = {
    doc: () => fakeDoc(['Hello World', 'foo bar foo']),
    fileName: () => 'test.pdf',
    pageCount: () => 2,
    currentPage: () => 1,
    readOnly: () => false,
    outline: () => null,
    searchIndex: () => Promise.resolve(INDEX),
    isDeleted: () => false,
    gotoPage: vi.fn(() => true),
    addMarkup: vi.fn(),
    formEdits: () => new Map<string, FormValueInput>(),
    applyFormEdit: vi.fn(),
    rotatePage: vi.fn(),
    deletePage: vi.fn(() => true),
    ...over,
  }
  // SAFETY: the fixture itself always supplies addMarkup as a vi.fn(); only its .mock is read
  return deps as MockedDeps
}

const call = (name: string, input: Record<string, unknown> = {}) => ({ id: 't1', name, input })

describe('PDF_TOOLS definitions', () => {
  it('declares unique names and object input schemas with required fields present', () => {
    const names = PDF_TOOLS.map((t) => t.name)
    expect(new Set(names).size).toBe(names.length)
    for (const tool of PDF_TOOLS) {
      expect(tool.description ?? tool.name).toBeTruthy()
      expect(tool.inputSchema.type).toBe('object')
      const props = tool.inputSchema.properties as Record<string, unknown>
      for (const req of (tool.inputSchema.required as string[] | undefined) ?? []) {
        expect(props).toHaveProperty(req)
      }
    }
  })

  it('every declared tool is handled by executePdfTool', async () => {
    // A tool falling into the default branch would return "Unknown tool"
    const deps = makeDeps({ doc: () => null, searchIndex: () => null })
    for (const tool of PDF_TOOLS) {
      const result = await executePdfTool(deps, call(tool.name, { page: 1, start: 1 }))
      expect(result.output).not.toContain('Unknown tool')
    }
  })
})

describe('read_pages', () => {
  it('reads a page range with [Page N] markers', async () => {
    const result = await executePdfTool(makeDeps(), call('read_pages', { start: 1, end: 2 }))
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('[Page 1]')
    expect(result.output).toContain('Hello World')
    expect(result.output).toContain('[Page 2]')
    expect(result.output).toContain('foo bar foo')
  })

  it('rejects an out-of-range start page', async () => {
    const result = await executePdfTool(makeDeps(), call('read_pages', { start: 5 }))
    expect(result.isError).toBe(true)
    expect(result.output).toContain('Invalid page range')
  })

  it('errors when the document is not loaded', async () => {
    const result = await executePdfTool(
      makeDeps({ doc: () => null }),
      call('read_pages', { start: 1 }),
    )
    expect(result.isError).toBe(true)
  })
})

describe('search_text', () => {
  it('returns page numbers with context excerpts', async () => {
    const result = await executePdfTool(makeDeps(), call('search_text', { query: 'FOO' }))
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('Page 2')
    expect(result.output).toContain('foo bar foo')
  })

  it('rejects an empty query', async () => {
    const result = await executePdfTool(makeDeps(), call('search_text', { query: '  ' }))
    expect(result.isError).toBe(true)
  })

  it('reports no matches without erroring', async () => {
    const result = await executePdfTool(makeDeps(), call('search_text', { query: 'zzz' }))
    expect(result.output).toBe('No matches found')
  })
})

describe('goto_page', () => {
  it('scrolls to a valid page', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(deps, call('goto_page', { page: 2 }))
    expect(result.isError).toBeUndefined()
    expect(deps.gotoPage).toHaveBeenCalledWith(2)
  })

  it('rejects out-of-range and deleted pages', async () => {
    const bad = await executePdfTool(makeDeps(), call('goto_page', { page: 3 }))
    expect(bad.isError).toBe(true)
    expect(bad.output).toContain('out of range')

    const deleted = await executePdfTool(
      makeDeps({ isDeleted: (i) => i === 0 }),
      call('goto_page', { page: 1 }),
    )
    expect(deleted.isError).toBe(true)
    expect(deleted.output).toContain('deleted')
  })
})

describe('readOnly gates the write tools', () => {
  it('refuses markup, form fills, page ops while the document is password-protected', async () => {
    const deps = makeDeps({ readOnly: () => true })
    for (const [name, input] of [
      ['markup_text', { query: 'Hello' }],
      ['fill_form_field', { name: 'name', value: 'A' }],
      ['rotate_page', { page: 1 }],
      ['delete_page', { page: 2 }],
    ] as const) {
      const result = await executePdfTool(deps, call(name, { page: 1, ...input }))
      expect(result.isError, name).toBe(true)
      expect(result.output).toContain('password-protected')
    }
    // Listing stays available: it reads, it does not write
    const listed = await executePdfTool(
      makeDeps({ readOnly: () => true }),
      call('list_form_fields'),
    )
    expect(listed.isError).toBeUndefined()
  })

  /** Per-tool expectations: markup and the form tools read the document; rotate/delete act on
      the App state alone and legitimately work without it */
  it('writes without a loaded document behave per tool', async () => {
    const deps = makeDeps({ doc: () => null })
    const markup = await executePdfTool(
      makeDeps({ doc: () => null, searchIndex: () => null }),
      call('markup_text', { query: 'z' }),
    )
    expect(markup.isError).toBe(true)
    expect(markup.output).toContain('not ready')

    const filled = await executePdfTool(deps, call('fill_form_field', { name: 'x', value: 'y' }))
    expect(filled.isError).toBe(true)
    expect(filled.output).toContain('Unknown field')

    await executePdfTool(deps, call('rotate_page', { page: 1 }))
    expect(deps.rotatePage).toHaveBeenCalled()
    await executePdfTool(deps, call('delete_page', { page: 2 }))
    expect(deps.deletePage).toHaveBeenCalled()
  })
})

describe('list_form_fields', () => {
  const widgets = [
    [{ annotationType: 20, fieldName: 'name', fieldValue: 'Alice' }],
    [
      { annotationType: 20, fieldName: 'agree', checkBox: true, fieldValue: true },
      {
        annotationType: 20,
        fieldName: 'plan',
        radioButton: true,
        options: [{ exportValue: 'basic' }, { exportValue: 'pro' }],
      },
      { annotationType: 20, fieldName: 'locked', fieldValue: 'x', readOnly: true },
    ],
  ]

  it('lists every field with kind and value', async () => {
    const deps = makeDeps({ doc: () => fakeDoc(['Hello World', 'foo bar foo'], widgets) })
    const result = await executePdfTool(deps, call('list_form_fields'))
    expect(result.isError).toBeUndefined()
    const lines = result.output.split('\n')
    expect(lines[0]).toBe('- name (text) = "Alice"')
    expect(lines[1]).toBe('- agree (checkbox) = "true"')
    expect(lines[2]).toBe('- plan (radio) = ""')
    expect(lines[3]).toBe('- locked (text) = "x" [read-only]')
  })

  it('reports when the document has no fields', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(deps, call('list_form_fields'))
    expect(result.output).toContain('no form fields')
  })
})

describe('fill_form_field', () => {
  const widgets = [
    [{ annotationType: 20, fieldName: 'name', fieldValue: '' }],
    [
      { annotationType: 20, fieldName: 'agree', checkBox: true, fieldValue: false },
      {
        annotationType: 20,
        fieldName: 'plan',
        radioButton: true,
        options: [{ exportValue: 'basic' }, { exportValue: 'pro' }],
      },
      { annotationType: 20, fieldName: 'locked', fieldValue: '', readOnly: true },
    ],
  ]

  it('fills a text field through applyFormEdit', async () => {
    const deps = makeDeps({ doc: () => fakeDoc(['Hello World', 'foo bar foo'], widgets) })
    const result = await executePdfTool(
      deps,
      call('fill_form_field', { name: 'name', value: 'Alice' }),
    )
    expect(result.isError).toBeUndefined()
    expect(deps.applyFormEdit).toHaveBeenCalledWith({ kind: 'text', name: 'name', value: 'Alice' })
  })

  it('maps checked values to the checkbox kind', async () => {
    const deps = makeDeps({ doc: () => fakeDoc(['Hello World', 'foo bar foo'], widgets) })
    const yes = await executePdfTool(
      deps,
      call('fill_form_field', { name: 'agree', value: 'true' }),
    )
    expect(deps.applyFormEdit).toHaveBeenCalledWith({
      kind: 'checkbox',
      name: 'agree',
      checked: true,
    })
    expect(yes.output).toContain('agree')

    await executePdfTool(deps, call('fill_form_field', { name: 'agree', value: 'false' }))
    expect(deps.applyFormEdit).toHaveBeenLastCalledWith({
      kind: 'checkbox',
      name: 'agree',
      checked: false,
    })
  })

  it('rejects a value that is not among the radio options', async () => {
    const deps = makeDeps({ doc: () => fakeDoc(['Hello World', 'foo bar foo'], widgets) })
    const result = await executePdfTool(
      deps,
      call('fill_form_field', { name: 'plan', value: 'gold' }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('basic, pro')
  })

  it('rejects unknown fields and read-only fields', async () => {
    const deps = makeDeps({ doc: () => fakeDoc(['Hello World', 'foo bar foo'], widgets) })
    const unknown = await executePdfTool(
      deps,
      call('fill_form_field', { name: 'nope', value: 'x' }),
    )
    expect(unknown.isError).toBe(true)
    expect(unknown.output).toContain('Unknown field')
    expect(unknown.output).toContain('locked')

    // Field-level read-only: 'locked' refuses any fill even with a valid value
    const ro = await executePdfTool(deps, call('fill_form_field', { name: 'locked', value: 'x' }))
    expect(ro.isError).toBe(true)
    expect(ro.output).toContain('read-only')
  })
})

describe('rotate_page', () => {
  it('rotates clockwise by default and counter-clockwise on request', async () => {
    const deps = makeDeps()
    await executePdfTool(deps, call('rotate_page', { page: 2 }))
    expect(deps.rotatePage).toHaveBeenCalledWith(1, 90)

    await executePdfTool(deps, call('rotate_page', { page: 1, direction: 'ccw' }))
    expect(deps.rotatePage).toHaveBeenCalledWith(0, -90)
  })

  it('rejects out-of-range and deleted pages', async () => {
    const bad = await executePdfTool(makeDeps(), call('rotate_page', { page: 9 }))
    expect(bad.isError).toBe(true)
    expect(bad.output).toContain('out of range')

    const deleted = await executePdfTool(
      makeDeps({ isDeleted: (i) => i === 1 }),
      call('rotate_page', { page: 2 }),
    )
    expect(deleted.isError).toBe(true)
  })
})

describe('delete_page', () => {
  it('removes a page through the undoable helper', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(deps, call('delete_page', { page: 2 }))
    expect(result.isError).toBeUndefined()
    expect(deps.deletePage).toHaveBeenCalledWith(1)
    expect(result.output).toContain('Removed page 2')
  })

  it('refuses to remove the last remaining page', async () => {
    const deps = makeDeps({ deletePage: () => false })
    const result = await executePdfTool(deps, call('delete_page', { page: 1 }))
    expect(result.isError).toBe(true)
    expect(result.output).toContain('one page has to remain')
  })
})

describe('markup_text', () => {
  it('marks every match of the query as the requested kind', async () => {
    const deps = makeDeps()
    const result = await executePdfTool(
      deps,
      call('markup_text', { query: 'foo', type: 'squiggly' }),
    )
    expect(result.isError).toBeUndefined()
    expect(deps.addMarkup).toHaveBeenCalledTimes(1)
    const [kind, pageIndex, rects] = deps.addMarkup.mock.calls[0]!
    expect(kind).toBe('squiggly')
    expect(pageIndex).toBe(1)
    // The interpolation floats: compare close to the item rectangle
    expect(rects).toHaveLength(2)
    const [a, b] = rects
    a.forEach((v, i) => expect(v).toBeCloseTo([0, 700, 30, 712][i]!))
    b.forEach((v, i) => expect(v).toBeCloseTo([80, 700, 110, 712][i]!))
    expect(result.output).toContain('Page 2: 2 marks (squiggly)')
  })

  it('marks one page and reports no matches otherwise', async () => {
    const deps = makeDeps()
    await executePdfTool(deps, call('markup_text', { query: 'foo', page: 2 }))
    const [kind, pageIndex] = deps.addMarkup.mock.calls[0]!
    expect(kind).toBe('highlight')
    expect(pageIndex).toBe(1)
    expect(deps.addMarkup.mock.calls[0]![2]).toHaveLength(2)

    const none = await executePdfTool(makeDeps(), call('markup_text', { query: 'zzz' }))
    expect(none.isError).toBe(true)
    // Locale-neutral assertion: the summary quotes the query in every language
    expect(none.output).toContain('"zzz"')
  })

  it('rejects an unknown markup kind and an empty query', async () => {
    const bad = await executePdfTool(makeDeps(), call('markup_text', { query: 'x', type: 'neon' }))
    expect(bad.isError).toBe(true)
    expect(bad.output).toContain('type must be one of')

    const empty = await executePdfTool(makeDeps(), call('markup_text', { query: '  ' }))
    expect(empty.isError).toBe(true)
  })
})

describe('get_outline / unknown tools', () => {
  it('renders the outline tree with indentation', async () => {
    const deps = makeDeps({
      outline: () => [
        { title: 'Chapter 1', items: [{ title: 'Section 1.1' }] },
        { title: 'Chapter 2' },
      ],
    })
    const result = await executePdfTool(deps, call('get_outline'))
    expect(result.output).toBe('Chapter 1\n  Section 1.1\nChapter 2')
  })

  it('reports when the document has no outline', async () => {
    const result = await executePdfTool(makeDeps(), call('get_outline'))
    expect(result.output).toContain('no outline')
  })

  it('errors on unknown tool names', async () => {
    const result = await executePdfTool(makeDeps(), call('nope'))
    expect(result.isError).toBe(true)
    expect(result.output).toContain('Unknown tool')
  })
})
