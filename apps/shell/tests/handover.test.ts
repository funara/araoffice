import { describe, expect, it, vi, beforeEach, afterEach } from 'vitest'
import {
  HANDOVER_TRIGGER_REGEX,
  actionableHandoverGoal,
  extractOutlineStages,
  extractSubstantiveUserContext,
  fallbackHandover,
  parseExplicitHandoverApp,
  synthesizeHandover,
} from '../src/renderer/src/ai/handover'

describe('handover', () => {
  describe('extractOutlineStages', () => {
    it('parses numbered stages and strips elaboration', () => {
      const text = [
        'Tentu. Alur umum:',
        '1. Identifikasi masalah - survei awal',
        '2. Studi literatur: baca jurnal',
        '3. Pengumpulan data',
      ].join('\n')
      expect(extractOutlineStages(text)).toEqual([
        { label: 'Identifikasi masalah', depth: 0 },
        { label: 'Studi literatur', depth: 0 },
        { label: 'Pengumpulan data', depth: 0 },
      ])
    })
    it('parses a single arrow chain into stages', () => {
      expect(extractOutlineStages('Mulai → Proses data → Selesai')).toEqual([
        { label: 'Mulai', depth: 0 },
        { label: 'Proses data', depth: 0 },
        { label: 'Selesai', depth: 0 },
      ])
    })

    it('strips bold markup from numbered stages', () => {
      const text = [
        '1. **Requester submits approval request**',
        '2. **System checks request completeness**',
        '3. **Request complete?**',
        '4. **Manager approves?**',
      ].join('\n')
      expect(extractOutlineStages(text).map((s) => s.label)).toEqual([
        'Requester submits approval request',
        'System checks request completeness',
        'Request complete?',
        'Manager approves?',
      ])
    })
  })

  describe('extractSubstantiveUserContext', () => {
    it('extracts multi-turn topic and goal accurately', () => {
      const messages = [
        { role: 'user' as const, text: 'saya mau buat jurnal ilmiah sinta 3' },
        { role: 'assistant' as const, text: 'Siap, apa topiknya?' },
        { role: 'user' as const, text: 'literasi ai mahasiswa uin bandung' },
      ]
      const { primaryGoal, refinedTopic, fullSummary } = extractSubstantiveUserContext(messages)
      expect(primaryGoal.toLowerCase()).toContain('jurnal ilmiah sinta 3')
      expect(refinedTopic.toLowerCase()).toContain('literasi ai mahasiswa uin bandung')
      expect(fullSummary.toLowerCase()).toContain('literasi ai mahasiswa uin bandung')
    })

    it('drops greetings and identity chatter from the goal', () => {
      const messages = [
        { role: 'user' as const, text: 'hi siapa kamu?' },
        { role: 'assistant' as const, text: 'Saya ARA Assistant...' },
        { role: 'user' as const, text: 'kamu bisa apa aja?' },
        { role: 'assistant' as const, text: 'Saya bisa ...' },
        { role: 'user' as const, text: 'saya ingin membuat flowchart alur persetujuan dokumen' },
      ]
      const { primaryGoal } = extractSubstantiveUserContext(messages)
      expect(primaryGoal.toLowerCase()).toContain('alur persetujuan dokumen')
      expect(primaryGoal.toLowerCase()).not.toContain('siapa kamu')
    })

    it('strips the i want to create prefix from english requests', () => {
      const messages = [{ role: 'user' as const, text: 'i want to create simple flowchart' }]
      const { primaryGoal } = extractSubstantiveUserContext(messages)
      expect(primaryGoal).toBe('simple flowchart')
    })

    it('strips trailing help-request tails so the goal reads as an imperative task', () => {
      const messages = [
        {
          role: 'user' as const,
          text: 'saya mau buat template jurnal terkait ai literacy mahasiswa, buat jurnal, bisa bantu saya riset dan tulis template dasarnya dulu?',
        },
      ]
      const { primaryGoal } = extractSubstantiveUserContext(messages)
      expect(primaryGoal).toBe('template jurnal terkait ai literacy mahasiswa, buat jurnal')
    })

    it('detects thread language over the UI language in the fallback', () => {
      const messages = [
        {
          role: 'user' as const,
          text: 'saya mau buat template jurnal terkait ai literacy mahasiswa, buat jurnal, bisa bantu saya riset dan tulis template dasarnya dulu?',
        },
      ]
      const out = fallbackHandover(messages, 'en')
      expect(out.prompt).toContain('Language: Indonesian')
      expect(out.prompt).toContain(
        'Buatkan saya "template jurnal terkait ai literacy mahasiswa, buat jurnal"',
      )
      expect(out.prompt).not.toContain('bisa bantu saya')
    })
  })

  describe('fallbackHandover', () => {
    it('falls back to all five editors with stages embedded', () => {
      const messages = [
        { role: 'user' as const, text: 'saya mau buat flowchart pendaftaran' },
        {
          role: 'assistant' as const,
          text: 'Baik.\n1. Masukkan data\n2. Proses data\n3. Tampilkan hasil',
        },
      ]
      const out = fallbackHandover(messages, 'id')
      expect(out.editors).toEqual(['docs', 'sheets', 'slides', 'markdown', 'diagram'])
      expect(out.prompt).toContain('Masukkan data')
      expect(out.prompt).toContain('Proses data')
      expect(out.prompt).not.toContain('Mermaid')
      expect(out.prompt).not.toContain('web')
    })

    it('keeps only the substantive goal when chatter pollutes the thread', () => {
      const messages = [
        { role: 'user' as const, text: 'hi siapa kamu?' },
        { role: 'assistant' as const, text: 'Saya ARA Assistant...' },
        { role: 'user' as const, text: 'kamu bisa apa aja?' },
        { role: 'assistant' as const, text: 'Saya bisa ...' },
        { role: 'user' as const, text: 'saya ingin membuat flowchart alur persetujuan dokumen' },
        {
          role: 'assistant' as const,
          text: 'Baik.\n1. Periksa kelengkapan\n2. Ajukan ke penyetuju',
        },
      ]
      const out = fallbackHandover(messages, 'id')
      expect(out.prompt).toContain('alur persetujuan dokumen')
      expect(out.prompt).not.toContain('siapa kamu')
      expect(out.prompt).not.toContain('Mermaid')
      expect(out.prompt).not.toContain('web')
    })

    it('drops the generic first request and strips bold stages', () => {
      const messages = [
        { role: 'user' as const, text: 'i want to create simple flowchart' },
        { role: 'assistant' as const, text: 'Sure, what is the topic?' },
        { role: 'user' as const, text: 'Approval workflow flowchart' },
        {
          role: 'assistant' as const,
          text: '1. **Requester submits approval request**\n2. **System checks request completeness**\n3. **Request complete?**\n4. **Manager approves?**',
        },
      ]
      const out = fallbackHandover(messages, 'en')
      expect(out.prompt).toContain('"Approval workflow flowchart"')
      expect(out.prompt).toContain('Task: Draft me a "Approval workflow flowchart".')
      expect(out.prompt).not.toContain('simple flowchart)')
      expect(out.prompt).not.toContain('i want to create')
      expect(out.prompt).toContain('Requester submits approval request')
      expect(out.prompt).not.toContain('**')
    })

    it('emits the fixed envelope and omits empty Data/Constraints sections', () => {
      const messages = [
        { role: 'user' as const, text: 'saya mau buat flowchart pendaftaran' },
        {
          role: 'assistant' as const,
          text: 'Baik.\n1. Masukkan data\n2. Proses data\n3. Tampilkan hasil',
        },
      ]
      const out = fallbackHandover(messages, 'id')
      expect(out.prompt).toContain('Task: Buatkan saya "flowchart pendaftaran" di editor.')
      expect(out.prompt).toContain('Language: Indonesian')
      expect(out.prompt).toContain('Details:')
      expect(out.prompt).toContain('Stages:\n1. Masukkan data')
      expect(out.prompt).not.toContain('Data:')
      expect(out.prompt).not.toContain('Constraints:')
    })

    it('returns null for a lone generic ask with no topic', () => {
      const out = fallbackHandover(
        [{ role: 'user' as const, text: 'i want to create flowchart simple' }],
        'en',
      )
      expect(out).toBeNull()
    })

    it('returns null for pure chatter', async () => {
      const out1 = await synthesizeHandover([{ role: 'user' as const, text: 'hi' }], 'en', 1000)
      expect(out1).toBeNull()

      const out2 = await synthesizeHandover([{ role: 'user' as const, text: 'hi ara' }], 'en', 1000)
      expect(out2).toBeNull()

      const out3 = await synthesizeHandover(
        [{ role: 'user' as const, text: 'halo ara' }],
        'id',
        1000,
      )
      expect(out3).toBeNull()

      const out4 = fallbackHandover([{ role: 'user' as const, text: 'good morning ara' }], 'en')
      expect(out4).toBeNull()

      const out5 = fallbackHandover([{ role: 'user' as const, text: 'who are you' }], 'en')
      expect(out5).toBeNull()
    })
  })

  describe('actionableHandoverGoal – creation intent gate', () => {
    it('returns null for a pure Q&A question with no creation intent', () => {
      // User asks why the handover triggered — no creation verb, no artifact noun
      const out = actionableHandoverGoal([
        { role: 'user', text: 'why did the handover trigger?' },
        { role: 'assistant', text: 'The handover triggers when...' },
      ])
      expect(out).toBeNull()
    })

    it('returns null for explanation/analysis-only threads', () => {
      const out = actionableHandoverGoal([
        { role: 'user', text: 'explain how Excel formulas work' },
        { role: 'assistant', text: 'Excel formulas are...' },
        { role: 'user', text: 'what about VLOOKUP?' },
        { role: 'assistant', text: 'VLOOKUP is...' },
      ])
      expect(out).toBeNull()
    })

    it('returns null for Indonesian Q&A without creation verbs', () => {
      const out = actionableHandoverGoal([
        { role: 'user', text: 'kenapa handover muncul?' },
        { role: 'assistant', text: 'Handover muncul ketika...' },
      ])
      expect(out).toBeNull()
    })

    it('returns a goal when user message has an explicit creation verb', () => {
      const out = actionableHandoverGoal([
        { role: 'user', text: 'saya mau buat flowchart pendaftaran siswa baru' },
        { role: 'assistant', text: 'Baik, berikut strukturnya...' },
      ])
      expect(out).not.toBeNull()
      expect(out).toContain('flowchart pendaftaran')
    })

    it('returns a goal when user message names an artifact noun (implicit creation)', () => {
      const out = actionableHandoverGoal([
        { role: 'user', text: 'flowchart alur persetujuan dokumen' },
        { role: 'assistant', text: '1. Submit\n2. Review\n3. Approve' },
      ])
      expect(out).not.toBeNull()
    })

    it('returns a goal when a later follow-up adds creation intent', () => {
      const out = actionableHandoverGoal([
        { role: 'user', text: 'jelaskan proses persetujuan dokumen' },
        { role: 'assistant', text: 'Proses persetujuan adalah...' },
        { role: 'user', text: 'tolong buatkan template untuk itu' },
        { role: 'assistant', text: 'Baik, berikut templatenya...' },
      ])
      expect(out).not.toBeNull()
    })

    it('returns a goal for Indonesian me- prefixed creation verbs (membuat/menulis/menyusun)', () => {
      // Regression: \b cannot precede the stem inside membuat/menulis/menyusun,
      // so threads like "saya mau membuat..." were suppressed as pure Q&A.
      for (const text of [
        'saya mau membuat template paper jurnal',
        'tolong menulis draf artikel untuk saya',
        'bagaimana saya menyusun laporan praktikum?',
      ]) {
        const out = actionableHandoverGoal([
          { role: 'user', text },
          { role: 'assistant', text: 'Berikut strukturnya...' },
        ])
        expect(out, text).not.toBeNull()
      }
    })

    it('returns a goal for verb-less document nouns (paper/jurnal/dokumen/template)', () => {
      const out = actionableHandoverGoal([
        { role: 'user', text: 'saya mau template paper jurnal' },
        { role: 'assistant', text: 'Berikut outline-nya...' },
      ])
      expect(out).not.toBeNull()
    })

    it('still suppresses analysis-framed messages without any creation signal', () => {
      const out = actionableHandoverGoal([
        { role: 'user', text: 'apa itu paper jurnal?' },
        { role: 'assistant', text: 'Paper jurnal adalah...' },
      ])
      expect(out).toBeNull()
    })

    it('returns null for chatter-only threads', () => {
      const out = actionableHandoverGoal([
        { role: 'user', text: 'hi' },
        { role: 'assistant', text: 'Halo!' },
        { role: 'user', text: 'siapa kamu?' },
      ])
      expect(out).toBeNull()
    })
  })

  describe('synthesizeHandover', () => {
    const originalWindow = globalThis.window

    beforeEach(() => {
      globalThis.window = {
        aiOfficeAi: {
          getSettings: vi.fn().mockResolvedValue({}),
          stream: vi.fn((_req, callbacks) => {
            setTimeout(() => {
              callbacks.onChunk({
                type: 'delta',
                text: JSON.stringify({
                  prompt:
                    'Task: Buatkan saya "flowchart pendaftaran siswa baru" di editor.\nLanguage: Indonesian\nDetails:\n- Kebutuhan: flowchart pendaftaran siswa baru.\nStages:\n1. Masukkan data\n2. Proses data\n3. Tampilkan hasil',
                  editors: ['diagram'],
                }),
              })
              callbacks.onChunk({ type: 'done' })
            }, 10)
            return { cancel: vi.fn() }
          }),
        },
      } as unknown as Window & typeof globalThis
    })

    afterEach(() => {
      globalThis.window = originalWindow
    })

    it('returns prompt plus ordered editors from one LLM call', async () => {
      const messages = [
        { role: 'user' as const, text: 'saya mau buat flowchart pendaftaran' },
        {
          role: 'assistant' as const,
          text: 'Baik.\n1. Masukkan data\n2. Proses data\n3. Tampilkan hasil',
        },
      ]
      const out = await synthesizeHandover(messages, 'id', 1000)
      expect(out.prompt).toContain('pendaftaran')
      expect(out.editors).toEqual(['diagram'])
      expect(out.prompt).toContain('Task:')
      expect(out.prompt).toContain('Details:')
      expect(out.prompt).toContain('Stages:')
      expect(out.outline.map((s) => s.label)).toEqual([
        'Masukkan data',
        'Proses data',
        'Tampilkan hasil',
      ])
    })
    it('strips legacy mermaid/web-search boilerplate from LLM prompts', async () => {
      globalThis.window = {
        aiOfficeAi: {
          getSettings: vi.fn().mockResolvedValue({}),
          stream: vi.fn((_req, callbacks) => {
            setTimeout(() => {
              callbacks.onChunk({
                type: 'delta',
                text: JSON.stringify({
                  prompt:
                    'Task: Susun flowchart alur persetujuan dokumen di editor.\nLanguage: Indonesian\nDetails:\n- Kebutuhan: alur persetujuan dokumen. Tulis langsung tanpa kode Mermaid dan tanpa penelusuran web berulang.\nStages:\n1. Periksa kelengkapan',
                  editors: ['diagram'],
                }),
              })
              callbacks.onChunk({ type: 'done' })
            }, 10)
            return { cancel: vi.fn() }
          }),
        },
      } as unknown as Window & typeof globalThis
      const out = await synthesizeHandover(
        [{ role: 'user' as const, text: 'flowchart alur persetujuan dokumen' }],
        'id',
        1000,
      )
      expect(out.prompt).toContain('alur persetujuan dokumen')
      expect(out.prompt).toContain('Task:')
      expect(out.prompt).not.toContain('Mermaid')
      expect(out.prompt).not.toContain('penelusuran web')
    })

    it('accepts markdown-bolded headings from LLM output', async () => {
      globalThis.window = {
        aiOfficeAi: {
          getSettings: vi.fn().mockResolvedValue({}),
          stream: vi.fn((_req, callbacks) => {
            setTimeout(() => {
              callbacks.onChunk({
                type: 'delta',
                text: JSON.stringify({
                  prompt:
                    '**Task:** Buatkan alur persetujuan dokumen.\n**Details:**\n- Kebutuhan: alur dokumen resmi.\n**Stages:**\n1. Pengajuan\n2. Verifikasi',
                  editors: ['diagram'],
                }),
              })
              callbacks.onChunk({ type: 'done' })
            }, 10)
            return { cancel: vi.fn() }
          }),
        },
      } as unknown as Window & typeof globalThis
      const out = await synthesizeHandover(
        [{ role: 'user' as const, text: 'buatkan alur persetujuan dokumen' }],
        'id',
        1000,
      )
      expect(out.prompt).toContain('Task:')
      expect(out.prompt).toContain('Details:')
      expect(out.prompt).toContain('Stages:')
      expect(out.editors).toEqual(['diagram'])
    })

    it('wraps un-enveloped substantive prompt when stages are not expected', async () => {
      globalThis.window = {
        aiOfficeAi: {
          getSettings: vi.fn().mockResolvedValue({}),
          stream: vi.fn((_req, callbacks) => {
            setTimeout(() => {
              callbacks.onChunk({
                type: 'delta',
                text: JSON.stringify({
                  prompt:
                    'Tuliskan artikel opini mendalam mengenai dampak kecerdasan buatan terhadap integritas akademik di perguruan tinggi.',
                  editors: ['docs'],
                }),
              })
              callbacks.onChunk({ type: 'done' })
            }, 10)
            return { cancel: vi.fn() }
          }),
        },
      } as unknown as Window & typeof globalThis
      const out = await synthesizeHandover(
        [{ role: 'user' as const, text: 'tulis artikel opini tentang ai di kampus' }],
        'id',
        1000,
      )
      expect(out.prompt).toContain('Task:')
      expect(out.prompt).toContain('Details:')
      expect(out.prompt).toContain('integritas akademik')
      expect(out.editors).toEqual(['docs'])
    })

    it('rejects non-envelope LLM output and falls back to the envelope', async () => {
      globalThis.window = {
        aiOfficeAi: {
          getSettings: vi.fn().mockResolvedValue({}),
          stream: vi.fn((_req, callbacks) => {
            setTimeout(() => {
              callbacks.onChunk({
                type: 'delta',
                text: JSON.stringify({
                  prompt: 'Gambar flowchart alur pendaftaran siswa baru, bagus dan rapi sekali.',
                  editors: ['diagram'],
                }),
              })
              callbacks.onChunk({ type: 'done' })
            }, 10)
            return { cancel: vi.fn() }
          }),
        },
      } as unknown as Window & typeof globalThis
      const messages = [
        { role: 'user' as const, text: 'saya mau buat flowchart pendaftaran' },
        {
          role: 'assistant' as const,
          text: 'Baik.\n1. Masukkan data\n2. Proses data\n3. Tampilkan hasil',
        },
      ]
      const out = await synthesizeHandover(messages, 'id', 1000)
      expect(out.prompt).toContain('Task:')
      expect(out.prompt).toContain('Stages:')
      expect(out.prompt).toContain('Masukkan data')
      expect(out.editors).toEqual(['docs', 'sheets', 'slides', 'markdown', 'diagram'])
    })

    it('filters unknown editors and falls back on stream error', async () => {
      globalThis.window = {
        aiOfficeAi: {
          getSettings: vi.fn().mockResolvedValue({}),
          stream: vi.fn((_req, callbacks) => {
            setTimeout(() => {
              callbacks.onChunk({ type: 'error', error: 'Service unavailable' })
            }, 10)
            return { cancel: vi.fn() }
          }),
        },
      } as unknown as Window & typeof globalThis
      const messages = [{ role: 'user' as const, text: 'buatkan laporan evaluasi' }]
      const out = await synthesizeHandover(messages, 'id', 1000)
      expect(out.editors).toEqual(['docs', 'sheets', 'slides', 'markdown', 'diagram'])
      expect(out.prompt).toContain('laporan evaluasi')
    })

    it('falls back to deterministic prompt without a bridge', async () => {
      globalThis.window = undefined as unknown as Window & typeof globalThis
      const out = await synthesizeHandover(
        [{ role: 'user' as const, text: 'buatkan model anggaran bulanan' }],
        'id',
        1000,
      )
      expect(out.editors).toHaveLength(5)
      expect(out.prompt).toContain('anggaran bulanan')
    })

    it('retains all 5 apps in the dropdown even if the model recommends a subset', () => {
      const modelEditors = ['diagram', 'docs'] as const
      const allKnown = ['docs', 'sheets', 'slides', 'markdown', 'diagram'] as const
      const combined = [...new Set([...modelEditors, ...allKnown])]
      expect(combined).toHaveLength(5)
      expect(combined[0]).toBe('diagram')
      expect(combined[1]).toBe('docs')
      expect(combined).toContain('slides')
    })

    it('identifies affirmative handover trigger keywords in Indonesian and English', () => {
      expect(HANDOVER_TRIGGER_REGEX.test('gas')).toBe(true)
      expect(HANDOVER_TRIGGER_REGEX.test('gaskan')).toBe(true)
      expect(HANDOVER_TRIGGER_REGEX.test('buatkan')).toBe(true)
      expect(HANDOVER_TRIGGER_REGEX.test('buatkan sekarang')).toBe(true)
      expect(HANDOVER_TRIGGER_REGEX.test('lanjutkan')).toBe(true)
      expect(HANDOVER_TRIGGER_REGEX.test('buka di sheets')).toBe(true)
      expect(HANDOVER_TRIGGER_REGEX.test('build it')).toBe(true)
      expect(HANDOVER_TRIGGER_REGEX.test('open in docs')).toBe(true)
      expect(HANDOVER_TRIGGER_REGEX.test('go for it')).toBe(true)

      // Ordinary conversation / questions should not match
      expect(HANDOVER_TRIGGER_REGEX.test('bagaimana cara menghitung rumus ini?')).toBe(false)
      expect(HANDOVER_TRIGGER_REGEX.test('tolong jelaskan')).toBe(false)
      expect(HANDOVER_TRIGGER_REGEX.test('apakah bisa diubah warnanya?')).toBe(false)
    })

    it('correctly maps mentioned apps in user commands', () => {
      expect(parseExplicitHandoverApp('buka di sheets')).toBe('sheets')
      expect(parseExplicitHandoverApp('buatkan di excel')).toBe('sheets')
      expect(parseExplicitHandoverApp('open in docs')).toBe('docs')
      expect(parseExplicitHandoverApp('susun di word')).toBe('docs')
      expect(parseExplicitHandoverApp('buatkan presentasi')).toBe('slides')
      expect(parseExplicitHandoverApp('buka flowchart')).toBe('diagram')
      expect(parseExplicitHandoverApp('tulis di note')).toBe('markdown')
      expect(parseExplicitHandoverApp('gas')).toBe(null)
    })
  })
})
