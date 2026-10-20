import { describe, expect, it } from 'vitest'
import type { Block, NumberingDef, ParsedDoc, TableModel } from '@araoffice/docx-engine'
import {
  asianCharCount,
  countDocumentWords,
  countWords,
  nonAsianWordCount,
  type FlowNode,
} from '../src/renderer/word-count'

describe('word count CJK rule', () => {
  it('counts asian chars one by one, punctuation included', () => {
    expect(asianCharCount('深圳市人民政府')).toBe(7)
    expect(asianCharCount('残疾预防，从我做起。')).toBe(10)
    expect(asianCharCount('（市残联、市卫生健康委）')).toBe(12)
    expect(asianCharCount('ひらがなカタカナ')).toBe(8)
    expect(asianCharCount('한국어')).toBe(3)
  })

  it('counts non-asian words by word, not by character', () => {
    expect(nonAsianWordCount('hello world')).toBe(2)
    expect(nonAsianWordCount("it's a state-of-the-art plan")).toBe(4)
    expect(nonAsianWordCount('2026年8月2日')).toBe(3)
    expect(nonAsianWordCount('深圳市')).toBe(0)
  })

  it('Words = asian chars + non-asian words (Word parity)', () => {
    const text = '按照 WHO 的要求，2025 年底前完成。'
    // 12 asian chars (incl. the fullwidth comma and period); non-asian words "WHO" + "2025" = 2
    expect(asianCharCount(text)).toBe(12)
    expect(nonAsianWordCount(text)).toBe(2)
    expect(countWords(text)).toBe(14)
  })

  it('counts a punctuation-bearing run as one word, like Word', () => {
    // Word counts space-delimited runs: these are one word each, not one per fragment
    expect(nonAsianWordCount('1.0')).toBe(1)
    expect(nonAsianWordCount('cld_aiops_source_completeness')).toBe(1)
    expect(nonAsianWordCount('state-of-the-art')).toBe(1)
    expect(nonAsianWordCount('v1.1')).toBe(1)
    expect(nonAsianWordCount('TSLAIOPS-SOP-260904-0009-0101')).toBe(1)
    expect(nonAsianWordCount('https://example.com/a/b')).toBe(1)
  })

  it('still splits a token at the asian characters inside it', () => {
    // "2026年8月2日" = three number runs separated by asian characters
    expect(nonAsianWordCount('2026\u5e748\u67082\u65e5')).toBe(3)
  })

  it('counts a standalone punctuation run as a word, as Word does', () => {
    expect(nonAsianWordCount('a | b')).toBe(3)
  })

  it('half-width digits/latin embedded in CJK are words, full-width punctuation is a char', () => {
    expect(countWords('第1章：ARAOffice 使用指南')).toBe(9)
    // 7 asian chars (incl. the fullwidth colon) + "1" and "ARAOffice" as 2 words
  })
})

// ---- document-level count (flow text + protected display models) ----

function text(value: string): FlowNode {
  return { isText: true, text: value, type: { name: 'text' }, attrs: {}, descendants: () => {} }
}

function flowNode(
  name: string,
  children: FlowNode[],
  attrs: Record<string, unknown> = {},
): FlowNode {
  const walk = (list: FlowNode[], fn: (node: FlowNode) => boolean | void): void => {
    for (const child of list) {
      if (fn(child) === false) continue
      child.descendants((descendant) => fn(descendant))
    }
  }
  return {
    type: { name },
    attrs,
    descendants(fn) {
      walk(children, fn)
    },
  }
}

function parsedDoc(partial: Partial<ParsedDoc>): ParsedDoc {
  return {
    blocks: [],
    numbering: new Map<string, NumberingDef>(),
    footnotes: [],
    endnotes: [],
    ...partial,
  } as unknown as ParsedDoc
}

function numberedList(): Map<string, NumberingDef> {
  const decimal: NumberingDef = {
    numId: '1',
    abstractNumId: '0',
    levels: { 0: { numFmt: 'decimal', lvlText: '%1.', start: 1 } },
    startOverrides: {},
  }
  const bullet: NumberingDef = {
    numId: '2',
    abstractNumId: '1',
    levels: { 0: { numFmt: 'bullet', lvlText: '\u2022', start: 1 } },
    startOverrides: {},
  }
  return new Map([
    ['1', decimal],
    ['2', bullet],
  ])
}

describe('countDocumentWords', () => {
  it('counts flow text, list numbers (not bullets), and inline atoms', () => {
    const doc = flowNode('doc', [
      flowNode('docListItem', [text('alpha')], { numId: '1', ilvl: 0 }),
      flowNode('docListItem', [text('beta')], { numId: '1', ilvl: 0 }),
      flowNode('docListItem', [text('gamma')], { numId: '2', ilvl: 0 }),
      flowNode('docParagraph', [
        text('inline '),
        flowNode('docInlineMath', [], { text: 'E = mc2' }),
        flowNode('docRuby', [], { base: '漢字', rt: 'かんじ' }),
      ]),
    ])
    // alpha/beta/gamma (3 words) + ordered numbers "1." "2." (2; the bullet is not
    // counted) + "inline" (1) + "E = mc2" (3) + ruby base/rt (5 asian chars)
    expect(countDocumentWords(parsedDoc({ numbering: numberedList() }), doc)).toBe(14)
  })

  it('counts protected display content: TOC entries, field results, text boxes, equations, section breaks', () => {
    const blocks = [
      {
        type: 'passthrough',
        fieldDisplay: { kind: 'tocLine', num: '3.1.6', left: 'Kafka', right: '32' },
      },
      // literal number already inside left: the numbering marker must not double count
      {
        type: 'passthrough',
        fieldDisplay: { kind: 'tocLine', num: '2.3.2', left: '2.3.2 Ingestion', right: '15' },
      },
      { type: 'passthrough', fieldDisplay: { kind: 'text', left: 'cached field result' } },
      {
        type: 'passthrough',
        label: 'Text box',
        textboxes: [{ paras: [{ runs: [{ text: 'boxed note' }] }] }],
      },
      { type: 'passthrough', label: 'Equation', previewText: 'a+b=c' },
      {
        type: 'passthrough',
        label: 'Section break paragraph',
        previewText: 'Document Number: SOP-1',
      },
      // text in another part: Word does not count it
      { type: 'passthrough', label: 'Chart', previewText: 'Quarterly chart title' },
      // editable blocks are counted from the flow, not here
      { type: 'paragraph', runs: [{ text: 'ignored here' }] },
    ] as unknown as Block[]
    const doc = flowNode('doc', [])
    // tocLine "3.1.6 Kafka" + "32" (3) + literal-number entry "2.3.2 Ingestion" + "15" (3)
    // + field result (3) + text box (2) + equation (1) + section break (3); chart and
    // editable-paragraph content are not counted here
    expect(countDocumentWords(parsedDoc({ blocks }), doc)).toBe(15)
  })

  it('counts nested table models and notes, excludes headers/footers', () => {
    const nested: TableModel = {
      rows: [[{ paras: ['deep cell'] }]],
    }
    const doc = flowNode('doc', [flowNode('docNestedTable', [], { model: nested })])
    const parsed = parsedDoc({
      footnotes: [{ text: 'footnote words here' }] as ParsedDoc['footnotes'],
      endnotes: [{ text: 'endnote' }] as ParsedDoc['endnotes'],
    })
    expect(countDocumentWords(parsed, doc)).toBe(2 + 3 + 1)
  })
})
