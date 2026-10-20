/** Native tools must remain usable when creating a deck from a blank slide. */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'
import type { RenderSlide, PlacedBox, ShapeRenderNode } from '@araoffice/pptx-render'
import type { AgentToolCall } from '../src/shared/ipc'

const box = (x: number, y: number, w: number, h: number): PlacedBox => ({
  x,
  y,
  w,
  h,
  rotationDeg: 0,
  flipH: false,
  flipV: false,
  centerX: x + w / 2,
  centerY: y + h / 2,
})
const textNode = (id: string, text: string): ShapeRenderNode =>
  ({
    id,
    sourceId: id,
    type: 'shape',
    box: box(60, 60, 400, 80),
    fill: { kind: 'none' },
    text: {
      lines: [
        {
          runs: [
            {
              text,
              x: 0,
              baselineY: 20,
              fontFamily: 'Arial',
              fontSizePx: 24,
              color: '#000',
              bold: false,
              italic: false,
              underline: false,
              widthPx: 100,
            },
          ],
          top: 0,
          height: 28,
        },
      ],
      insets: { l: 0, t: 0, r: 0, b: 0 },
      anchor: 'top',
      fontScale: 1,
      contentHeight: 28,
    },
  }) as unknown as ShapeRenderNode

// Blank deck (1 empty page, no text content) -> from-scratch scenario
const blankDeck = { widthPx: 1280, heightPx: 720, nodes: [] } as unknown as RenderSlide
// Existing polished deck (multiple elements with text) -> refinement scenario
const richDeck = {
  widthPx: 1280,
  heightPx: 720,
  nodes: [
    textNode('a', 'Title'),
    textNode('b', 'Point 1'),
    textNode('c', 'Point 2'),
    textNode('d', 'Point 3'),
  ],
} as unknown as RenderSlide

function mkAccess(slides: RenderSlide[]): DeckAccess {
  return {
    getSlides: () => slides,
    getCurrent: () => 0,
    getSelectedIds: () => [],
    applySlide: () => {},
    applyDeck: () => {},
    fitWidthPx: 1280,
    generateFromHtml: async () => ({ ok: true, pages: 1 }),
  } as unknown as DeckAccess
}
const call = (name: string): AgentToolCall => ({
  id: 't',
  name,
  input: {
    slideIndex: 0,
    paragraphs: [{ runs: [{ text: 'x' }] }],
    kind: 'rect',
    x: 10,
    y: 10,
    w: 100,
    h: 50,
    layout: 'list',
    items: ['a', 'b'],
  },
})

const addElement = vi.fn(async () => ({ slide: blankDeck, sourceId: 'e1' }))
const addSmartArt = vi.fn(async () => ({ slide: blankDeck, sourceId: 's1' }))
const addBlankSlide = vi.fn(async () => ({ slides: [blankDeck, blankDeck], index: 1 }))
const webSearch = vi.fn(async (): Promise<{ answer: string; results: []; method?: string }> => ({
  answer: '',
  results: [],
}))
let originalSlidesApi: PropertyDescriptor | undefined

beforeEach(() => {
  originalSlidesApi = Object.getOwnPropertyDescriptor(window, 'slidesApi')
  addElement.mockClear()
  addSmartArt.mockClear()
  addBlankSlide.mockClear()
  webSearch.mockClear()
  Object.defineProperty(window, 'slidesApi', {
    configurable: true,
    value: { addElement, addSmartArt, addBlankSlide, webSearch },
  })
})

afterEach(() => {
  if (originalSlidesApi) Object.defineProperty(window, 'slidesApi', originalSlidesApi)
  else Reflect.deleteProperty(window, 'slidesApi')
})

describe('native deck creation', () => {
  it('does not prescribe an unregistered cloud generator', () => {
    expect(createSlidesSkill(mkAccess([blankDeck])).systemPrompt).not.toContain('generate_deck')
  })
  it('routes an approved plan to native creation', async () => {
    const r = await createSlidesSkill(mkAccess([blankDeck])).executeTool!({
      id: 't',
      name: 'plan_deck',
      input: {
        core_hook: 'A focused onboarding story',
        style: 'Warm minimal',
        pages: [{ title: 'Welcome', brief: 'A concise introduction', layout: 'cover' }],
      },
    })
    expect(r.output).not.toContain('generate_deck')
    expect(r.output).toContain('Immediately make the first native page mutation')
    expect(r.output).toContain('Do not search')
  })
  it('directs a zero-result search to continue natively', async () => {
    const r = await createSlidesSkill(mkAccess([blankDeck])).executeTool!({
      id: 't',
      name: 'web_search',
      input: { query: 'employee onboarding statistics' },
    })
    expect(r.output).toContain('Do not retry this search')
    expect(r.isError).toBeUndefined()
  })
  it('reports a provider failure as a tool error', async () => {
    webSearch.mockResolvedValueOnce({ method: 'error', answer: '', results: [] })
    const r = await createSlidesSkill(mkAccess([blankDeck])).executeTool!({
      id: 't',
      name: 'web_search',
      input: { query: 'employee onboarding statistics' },
    })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('service error')
    expect(r.output).not.toContain('Do not retry this search')
  })
  it('creates later AI pages from an empty layout', async () => {
    const r = await createSlidesSkill(mkAccess([blankDeck])).executeTool!({
      id: 't',
      name: 'add_slide',
      input: { sourceIndex: 0 },
    })
    expect(r.isError).toBeUndefined()
    expect(addBlankSlide).toHaveBeenCalledWith({ sourceIndex: 0, fitWidthPx: 1280 })
  })

  it('clamps out-of-range sourceIndex in add_slide to last valid index', async () => {
    const r = await createSlidesSkill(mkAccess([blankDeck])).executeTool!({
      id: 't',
      name: 'add_slide',
      input: { sourceIndex: 99 },
    })
    expect(r.isError).toBeUndefined()
    expect(addBlankSlide).toHaveBeenCalledWith({ sourceIndex: 0, fitWidthPx: 1280 })
  })

  it('scaffolds all planned slides and layout containers in plan_deck', async () => {
    let slidesList = [blankDeck]
    const access = {
      ...mkAccess(slidesList),
      getSlides: () => slidesList,
      applyDeck: (all: RenderSlide[]) => {
        slidesList = all
      },
    } as unknown as DeckAccess

    addBlankSlide.mockImplementation(async () => {
      slidesList = [...slidesList, blankDeck]
      return { slides: slidesList, index: slidesList.length - 1 }
    })

    const r = await createSlidesSkill(access).executeTool!({
      id: 't',
      name: 'plan_deck',
      input: {
        core_hook: 'A scalable cloud architecture',
        style: 'Modern Slate',
        pages: [
          { title: 'Intro', brief: 'Hook and overview', layout: 'hero' },
          { title: 'Features', brief: 'Key capabilities', layout: 'three_column_cards' },
          { title: 'Comparison', brief: 'Before and after', layout: 'two_column' },
        ],
      },
    })
    expect(r.isError).toBeUndefined()
    expect(r.mutated).toBe(true)
    expect(addBlankSlide).toHaveBeenCalledTimes(2)
    expect(addElement).toHaveBeenCalled()
  })
  it('empty deck + add_text_box → allowed', async () => {
    const r = await createSlidesSkill(mkAccess([blankDeck])).executeTool!(call('add_text_box'))
    expect(r.isError).toBeUndefined()
    expect(addElement).toHaveBeenCalledOnce()
  })
  it('empty deck + add_shape → allowed', async () => {
    const r = await createSlidesSkill(mkAccess([blankDeck])).executeTool!(call('add_shape'))
    expect(r.isError).toBeUndefined()
    expect(addElement).toHaveBeenCalledOnce()
  })
  it('empty deck + add_smartart → allowed', async () => {
    const r = await createSlidesSkill(mkAccess([blankDeck])).executeTool!(call('add_smartart'))
    expect(r.isError).toBeUndefined()
    expect(addSmartArt).toHaveBeenCalledOnce()
  })
  it('existing rich deck + add_text_box → allowed', async () => {
    const r = await createSlidesSkill(mkAccess([richDeck])).executeTool!(call('add_text_box'))
    expect(r.isError).toBeUndefined()
    expect(addElement).toHaveBeenCalledOnce()
  })

  it('add_text_box with string text property passes extracted paragraphs to addElement', async () => {
    const r = await createSlidesSkill(mkAccess([blankDeck])).executeTool!({
      id: 't1',
      name: 'add_text_box',
      input: {
        slideIndex: 0,
        x: 100,
        y: 100,
        w: 400,
        h: 100,
        text: 'Nova Launch Deck\nPlan - Build - Ship',
      },
    })
    expect(r.isError).toBeUndefined()
    expect(addElement).toHaveBeenCalledWith(
      expect.objectContaining({
        slideIndex: 0,
        kind: 'textbox',
        paragraphs: [
          { runs: [{ text: 'Nova Launch Deck' }] },
          { runs: [{ text: 'Plan - Build - Ship' }] },
        ],
      }),
    )
  })

  it('add_text_box with string array paragraphs passes correct text runs', async () => {
    const r = await createSlidesSkill(mkAccess([blankDeck])).executeTool!({
      id: 't2',
      name: 'add_text_box',
      input: {
        slideIndex: 0,
        x: 50,
        y: 50,
        w: 300,
        h: 80,
        paragraphs: ['Point 1', 'Point 2'],
      },
    })
    expect(r.isError).toBeUndefined()
    expect(addElement).toHaveBeenCalledWith(
      expect.objectContaining({
        slideIndex: 0,
        kind: 'textbox',
        paragraphs: [{ runs: [{ text: 'Point 1' }] }, { runs: [{ text: 'Point 2' }] }],
      }),
    )
  })

  it('add_shape with string content/text passes text to shape element', async () => {
    const r = await createSlidesSkill(mkAccess([blankDeck])).executeTool!({
      id: 't3',
      name: 'add_shape',
      input: {
        slideIndex: 0,
        kind: 'rect',
        x: 50,
        y: 50,
        w: 300,
        h: 80,
        fillColor: '#1a1a2e',
        text: 'Shape Card Content',
      },
    })
    expect(r.isError).toBeUndefined()
    expect(addElement).toHaveBeenCalledWith(
      expect.objectContaining({
        slideIndex: 0,
        kind: 'rect',
        fillColor: '#1a1a2e',
        paragraphs: [{ runs: [{ text: 'Shape Card Content' }] }],
      }),
    )
  })

  it('add_text_box with object with runs passes through intact', async () => {
    const r = await createSlidesSkill(mkAccess([blankDeck])).executeTool!({
      id: 't4',
      name: 'add_text_box',
      input: {
        slideIndex: 0,
        x: 10,
        y: 10,
        w: 200,
        h: 50,
        paragraphs: [{ runs: [{ text: 'Custom Run', bold: true, color: '#ff0000' }] }],
      },
    })
    expect(r.isError).toBeUndefined()
    expect(addElement).toHaveBeenCalledWith(
      expect.objectContaining({
        slideIndex: 0,
        kind: 'textbox',
        paragraphs: [{ runs: [{ text: 'Custom Run', bold: true, color: '#ff0000' }] }],
      }),
    )
  })
})
