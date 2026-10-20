import { sanitizeModelIdentity } from '@araoffice/agent-core'
import { extractThinking } from '@araoffice/ui'
import type { AiStreamRequest } from '@araoffice/ai-provider'
import type { Lang } from '@araoffice/i18n'
import type { OmnibarMessage } from './omnibar-service'
import { isIndonesianText } from './omnibar-prompt'

const LANGUAGE_NAMES: Record<Lang, string> = {
  zh: 'Chinese',
  en: 'English',
  ja: 'Japanese',
  ko: 'Korean',
  fr: 'French',
  de: 'German',
  es: 'Spanish',
  th: 'Thai',
  id: 'Indonesian',
  ru: 'Russian',
  ar: 'Arabic',
  pt: 'Portuguese',
  it: 'Italian',
  pl: 'Polish',
  nl: 'Dutch',
  ms: 'Malay',
  he: 'Hebrew',
  hi: 'Hindi',
  'zh-TW': 'Traditional Chinese',
}

const BRIEF_LABELS: Record<Lang, string> = {
  zh: '需求',
  en: 'Brief',
  ja: '概要',
  ko: '요약',
  fr: 'Besoin',
  de: 'Anforderung',
  es: 'Necesidad',
  th: 'ความต้องการ',
  id: 'Kebutuhan',
  ru: 'Задача',
  ar: 'المتطلب',
  pt: 'Necessidade',
  it: 'Esigenza',
  pl: 'Potrzeba',
  nl: 'Behoefte',
  ms: 'Keperluan',
  he: 'צורך',
  hi: 'आवश्यकता',
  'zh-TW': '需求',
}

function taskLine(lang: Lang, goal: string): string {
  switch (lang) {
    case 'id':
      return `Buatkan saya "${goal}" di editor.`
    case 'ms':
      return `Buatkan saya "${goal}" dalam editor.`
    case 'zh':
      return `在编辑器中为我创建“${goal}”。`
    case 'zh-TW':
      return `在編輯器中為我建立「${goal}」。`
    case 'ja':
      return `エディターで「${goal}」を作成してください。`
    case 'ko':
      return `에디터에서 "${goal}"을 만들어 주세요.`
    case 'fr':
      return `Rédigez-moi « ${goal} » dans l'éditeur.`
    case 'de':
      return `Erstelle mir „${goal}“ im Editor.`
    case 'es':
      return `Redáctame «${goal}» en el editor.`
    case 'th':
      return `สร้าง "${goal}" ในตัวแก้ไขให้ฉัน`
    case 'ru':
      return `Создайте мне «${goal}» в редакторе.`
    case 'ar':
      return `أنشئ لي "${goal}" في المحرر.`
    case 'pt':
      return `Elabore para mim "${goal}" no editor.`
    case 'it':
      return `Redigimi "${goal}" nell'editor.`
    case 'pl':
      return `Przygotuj mi „${goal}“ w edytorze.`
    case 'nl':
      return `Maak voor mij "${goal}" in de editor.`
    case 'he':
      return `צור עבורי "${goal}" בעורך.`
    case 'hi':
      return `संपादक में मेरे लिए "${goal}" तैयार करें।`
    default:
      return `Draft me a "${goal}".`
  }
}

// Thread language wins over UI: an English UI with an Indonesian thread must
// still synthesize Indonesian. Only Indonesian gets content-sniffed (the
// observed mismatch); every other thread follows the user's UI language.
function resolveHandoverLang(
  chatMessages: Array<{ role: string; text: string }>,
  uiLang: string,
): Lang {
  const base: Lang =
    uiLang === 'id' ||
    uiLang === 'ms' ||
    uiLang === 'zh' ||
    uiLang === 'zh-TW' ||
    uiLang === 'ja' ||
    uiLang === 'ko' ||
    uiLang === 'fr' ||
    uiLang === 'de' ||
    uiLang === 'es' ||
    uiLang === 'th' ||
    uiLang === 'ru' ||
    uiLang === 'ar' ||
    uiLang === 'pt' ||
    uiLang === 'it' ||
    uiLang === 'pl' ||
    uiLang === 'nl' ||
    uiLang === 'he' ||
    uiLang === 'hi'
      ? uiLang
      : 'en'
  if (base !== 'en') return base
  const lastUser = [...chatMessages].reverse().find((m) => m.role === 'user' && m.text?.trim())
    ?.text as string | undefined
  if (lastUser && isIndonesianText(lastUser)) return 'id'
  return base
}
export type TargetAppKind = 'docs' | 'sheets' | 'slides' | 'markdown' | 'diagram'

export interface Handover {
  prompt: string
  editors: TargetAppKind[]
  outline: Array<{ label: string; depth: number }>
}

export const HANDOVER_LABELS: Record<TargetAppKind, string> = {
  docs: 'ARA Docs',
  sheets: 'ARA Sheets',
  slides: 'ARA Slides',
  markdown: 'ARA Notes',
  diagram: 'ARA Diagrams',
}
export const KNOWN_APPS: TargetAppKind[] = ['docs', 'sheets', 'slides', 'markdown', 'diagram']

export const HANDOVER_TRIGGER_REGEX =
  /^(?:gas(?:kan)?|buatkan(?: sekarang)?|buatin|buka(?:kan)?|lanjutkan|eksekusi|oke buatkan|siap buatkan|build it|create it|open it|go for it|do it|proceed|open in (?:sheets|docs|slides|diagram|markdown))\b/i

export function parseExplicitHandoverApp(str: string): TargetAppKind | null {
  const lower = str.toLowerCase()
  if (lower.includes('sheet') || lower.includes('excel')) return 'sheets'
  if (lower.includes('doc') || lower.includes('word')) return 'docs'
  if (lower.includes('slide') || lower.includes('presentasi') || lower.includes('powerpoint'))
    return 'slides'
  if (lower.includes('diagram') || lower.includes('flowchart') || lower.includes('bagan'))
    return 'diagram'
  if (lower.includes('markdown') || lower.includes('note')) return 'markdown'
  return null
}

const CHATTER_PATTERNS = [
  /^(?:hi|halo|hallo|hello|hai|hey|hei|yo|oi|selamat\s+(?:pagi|siang|sore|malam|datang)|good\s+(?:morning|afternoon|evening|day)|pagi|siang|sore|malam|assalamu(?:'|a)?laikum)(?:\s+(?:ara|araoffice|ai|bot|there|friend|kawan|bro|bray|gan|min))?[\s!.,?]*$/i,
  /(?:siapa (?:kamu|kau|dirimu|pembuatmu|penciptamu|yang buat kamu|yang bikin kamu|yang menciptakan kamu)|kamu siapa|who (?:are you|made you|created you|built you)|what are you)/i,
  /(?:(?:kamu|anda)\s+)?bisa apa(?:\s+aja|\s+saja)?|what can you do|apa yang bisa kamu lakukan|fitur apa aja|kamu bisa bantu apa|help me/i,
  /^(?:apa kabar|gimana kabarnya|how are you(?: doing)?|lagi apa|sedang apa|what'?s up)[\s!.,?]*$/i,
  /^(?:thanks|terima kasih|makasih|thank you|thx|ty|ok|oke|okay|baik|siap|mantap|bagus|sip|test|testing|tes|ping|help)[\s!.,]*$/i,
  /^(?:ara|araoffice|ai|ara\s+ai|bot)[\s!.,?]*$/i,
]

// ponytail: naive keyword chatter gate, per-message LLM classifier if miscategorized titles recur.
function isChatterMessage(text: string): boolean {
  return CHATTER_PATTERNS.some((re) => re.test(text.trim()))
}

// A first message that only names the artifact type ("simple flowchart",
// "dokumen baru") adds nothing once the real topic is known — drop it
// instead of titling "Topic (simple flowchart)".
const GENERIC_GOAL_PATTERN =
  /^(?:simple|sederhana|basic|baru|new|general|umum|short|quick|contoh|sample|test|coba|draft)?\s*(flowchart|diagram|alur|susunan|doc|dokumen|document|jurnal|journal|artikel|article|paper|laporan|report|sheet|spreadsheet|slide|presentasi|presentation|markdown|readme|teks|text|tabel|table|gambar|image)\s*(simple|sederhana|basic|baru|new)?$/i

// ponytail: keyword generic gate; LLM synthesis owns the smart merge, this just stops "X (simple flowchart)".
function isGenericGoal(goal: string): boolean {
  return GENERIC_GOAL_PATTERN.test(goal.trim())
}

// Explicit creation/drafting verbs that signal the user wants to produce a document.
// Indonesian me- prefixed verbs are spelled out: word boundary \b cannot precede
// a stem inside membuat/menulis/menyusun/merancang, so bare form\b misses them.
const CREATION_INTENT_RE =
  /\b(?:buat(?:kan|in|lah)?|membuat(?:kan)?|bikin(?:kan)?|tulis(?:kan|lah)?|menulis(?:kan)?|menuliskan|susun(?:kan|lah)?|menyusun(?:kan)?|rancang(?:kan)?|merancang(?:kan)?|rangkum(?:kan)?|merangkum|kembangkan|mengembangkan|isi(?:kan)?|mengisi|create|write|draft|build|generate|make|design|compose|prepare|develop|produce)\b/i

// Noun phrases that implicitly name a document artifact as their subject.
// Only meaningful when NOT paired with an analysis/question pattern.
// paper/jurnal/dokumen/etc. cover verb-less asks ("saya mau template paper jurnal").
const ARTIFACT_NOUN_RE =
  /\b(?:flowchart|diagram|mindmap|roadmap|notulen|minutes|kontrak|contract|sop|invoice|rekap|bagan|anggaran|budget|proposal|paper|jurnal|journal|dokumen|document|template)\b/i

// Analysis / explanation queries that override artifact nouns:
// "explain", "why", "what is", "how does", etc.
const ANALYSIS_INTENT_RE =
  /^(?:explain|describe|what(?:'s| is| are| does)?|why|how(?:\s+does)?|when|where|tell\s+me|show\s+me|is\s+it|are\s+(?:there|you)|jelaskan|kenapa|mengapa|apa\s+(?:itu|ini|yang)?|bagaimana|kapan|di\s+mana|ceritakan|tolong\s+jelaskan|bisakah\s+kamu\s+(?:jelasin|ceritain))/i

/**
 * Returns true when the thread contains at least one user message expressing
 * creation intent (explicit verb OR artifact noun without analysis framing).
 * Pure Q&A / explanation threads return false and suppress the handover card.
 */
function hasCreationIntent(chatMessages: Array<{ role: string; text: string }>): boolean {
  return chatMessages
    .filter((m) => m.role === 'user')
    .some((m) => {
      const t = m.text.trim()
      // Explicit creation verb always counts.
      if (CREATION_INTENT_RE.test(t)) return true
      // Artifact noun only counts when the message is NOT an analysis/question.
      if (ARTIFACT_NOUN_RE.test(t) && !ANALYSIS_INTENT_RE.test(t)) return true
      return false
    })
}

/**
 * Extracts substantive user goal and topic refinements from the multi-turn chat history.
 */
export function extractSubstantiveUserContext(
  chatMessages: Array<{ role: 'user' | 'assistant'; text: string }>,
): {
  primaryGoal: string
  refinedTopic: string
  fullSummary: string
} {
  const userMessages = chatMessages
    .filter((m) => m.role === 'user')
    .map((m) => m.text.trim())
    .filter(Boolean)
    .filter((t) => !isChatterMessage(t))

  if (userMessages.length === 0) {
    return {
      primaryGoal: '',
      refinedTopic: '',
      fullSummary: '',
    }
  }

  const cleanQuery = (q: string) =>
    q
      .replace(
        /^(?:tolong\s+)?(?:(?:saya|aku|gw|gue) (?:mau|ingin|pengen) (?:membuatkan|membuat|buatkan|buatin|buat)|tolong (?:buatkan|buatin|buat)|(?:buatkan|buatin|buat|bikin))\s+/i,
        '',
      )
      .replace(
        /^(?:i (?:want|wanna|would like|'d like) to (?:make|create|build|generate|write|draft)|(?:can|could) you (?:please )?(?:make|create|build|generate|write)|please (?:help me(?: to)? )?(?:make|create|build|generate|write|draft)?|help me (?:to )?(?:make|create|build|generate|write|draft)|i need an?|(?:please )?(?:draft|create|make|generate|write|build))\s+/i,
        '',
      )
      // Trailing request tail ("..., bisa bantu saya ...?", ", can you ...?"): even the
      // raw-quote fallback then reads as an imperative task. Bare "bisa"/"tolong"
      // needs a preceding comma so qualifiers like "yang bisa diedit" survive.
      .replace(
        /(?:[,;]\s*(?:bisa|tolong)\b|\b(?:bisakah|bolehkah|dapatkah|can you|could you|please|terima\s*kasih|makasih|thanks|thank\s*you)\b)[\s\S]*$/i,
        '',
      )
      .replace(/\s*[?!.…]+\s*$/, '')
      .replace(/[,;]\s*$/, '')
      .trim()

  const firstMsg = userMessages[0]
  const lastMsg = userMessages[userMessages.length - 1]

  const primaryGoal = cleanQuery(firstMsg)
  const refinedTopic = userMessages.length > 1 ? cleanQuery(lastMsg) : ''

  let fullSummary = primaryGoal
  if (refinedTopic && refinedTopic.toLowerCase() !== primaryGoal.toLowerCase()) {
    fullSummary = `${primaryGoal} - Topik: ${refinedTopic}`
  }

  return { primaryGoal, refinedTopic, fullSummary }
}

/**
 * Extracts ordered stage labels from an assistant outline. Handles three
 * shapes the omnibar actually emits (regression from screenshots):
 *  1. numbered/bulleted lines ("1. X", "- X", "• X")
 *  2. ASCII tree lines ("├── X", "└── X", "│ ...") — tree glyphs stripped,
 *     depth inferred from indent so hierarchies survive, not flatten
 *  3. a single arrow-chain line ("A → B → C")
 * Strips per-stage elaboration after ' - '/':'/'→', drops chatter and
 * followup lines, dedupes, caps at `max`. Agnostic: no domain assumptions.
 */
export function extractOutlineStages(
  assistantText: string,
  max = 12,
): Array<{ label: string; depth: number }> {
  if (!assistantText) return []
  const stages: Array<{ label: string; depth: number }> = []
  const seen = new Set<string>()
  const push = (label: string, depth: number) => {
    const clean = label
      .replace(/\*\*([^*]+)\*\*/g, '$1')
      .replace(/__([^_]+)__/g, '$1')
      .replace(/^#+\s*/, '')
      .replace(/^["'“”‘’`*_~]+|["'“”‘’`*_~]+$/g, '')
      .trim()
    if (!clean || clean.length < 2 || clean.length > 120) return
    const key = clean.toLowerCase()
    if (seen.has(key)) return
    if (stages.length >= max) return
    seen.add(key)
    stages.push({ label: clean, depth })
  }
  for (const rawLine of assistantText.split('\n')) {
    if (stages.length >= max) break
    const line = rawLine.replace(/\r$/, '')
    if (!line.trim()) continue
    // ASCII tree: measure indent (leading spaces + │) for depth, strip glyphs.
    const tree = /^\s*[│|\s]*(?:├──|└──|├─|└─|├|└)\s*(.*)$/.exec(line)
    if (tree?.[1]?.trim()) {
      const indent = /^\s*(?:[│|\s]*)/.exec(line)?.[0].length ?? 0
      push(tree[1].split(/\s+[-–—]\s+|\s*:\s+/)[0] ?? '', Math.min(3, Math.floor(indent / 2)))
      continue
    }
    const item = /^(?:\d+[.、)]|[-*•])\s+(.*)$/.exec(line.trim())?.[1]?.trim()
    if (item) {
      if (/^(?:open\s+in|buka\s+di)\s+ara/i.test(item)) continue
      if (/^(?:target editor prompt|prompt)\s*:/i.test(item)) continue
      // Strip elaboration: "Tahap - detail ...", "Tahap: detail", "A → B → C".
      const label = item.split(/\s+[-–—]\s+|\s*:\s+|\s*[→⇒]\s*/)[0] ?? ''
      push(label, 0)
      continue
    }
    // Single-line arrow chain fallback ("A → B → C").
    if (/[→⇒]/.test(line) && line.length < 400) {
      for (const part of line.split(/\s*[→⇒]\s*/)) push(part, 0)
    }
  }
  return stages
}

function lastAssistantText(chatMessages: OmnibarMessage[]): string {
  return (
    [...chatMessages].reverse().find((m) => m.role === 'assistant' && m.text?.trim())?.text ?? ''
  )
}

/**
 * Cleans quotation marks, code fences, thinking blocks, or prompt prefixes
 * from the LLM output.
 */
function cleanHandoverPrompt(text: string): string {
  if (!text) return ''
  const { text: withoutThinking } = extractThinking(text)
  const source = withoutThinking || text
  const cleaned = source
    .replace(/:::followups([\s\S]*?)(?:\s:::[ \t]*(?:\r?\n|$)|$)/gi, '')
    .replace(/^```(?:markdown|prompt|text)?\s*/i, '')
    .replace(/\s*```$/i, '')
    .replace(/^["'“”‘’`]|["'“”‘’`]$/g, '')
    .replace(/^(?:Prompt:\s*|Target Editor Prompt:\s*)/i, '')
    // Strip legacy meta-constraint sentences leaked from the old fallback
    // template (Mermaid / web-search boilerplate the target editor never needs).
    .replace(/\s*(?:tulis langsung tanpa[^.\n]*\.|write directly with[^.\n]*\.)/gi, '')
    .replace(/\s*(?:tanpa kode mermaid[^.\n]*\.|no mermaid[^.\n]*\.)/gi, '')
    .replace(/\s*(?:tanpa penelusuran web[^.\n]*\.|no repeated web[^.\n]*\.)/gi, '')
    .trim()
  return sanitizeModelIdentity(cleaned, { allowFallback: false }).trim()
}

function parseHandoverJson(
  raw: string,
  expectStages: boolean,
): { prompt: string; editors: TargetAppKind[] } | null {
  if (!raw || !raw.trim()) return null
  let parsed: unknown
  try {
    const fenced = /```(?:json)?\s*([\s\S]*?)```/i.exec(raw)
    parsed = JSON.parse((fenced?.[1] ?? raw).trim())
  } catch {
    return null
  }
  if (typeof parsed !== 'object' || parsed === null) return null
  const { prompt, editors } = parsed as Record<string, unknown>
  if (typeof prompt !== 'string') return null
  // Heading variants our own prompts emitted historically (fallback Detail: /
  // Tahapan:) or markdown-bolded headers normalized to the envelope anchors before validation.
  let cleaned = cleanHandoverPrompt(prompt)
    .replace(/^[#* \t]*(?:Task|Tugas)[* \t]*:/gim, 'Task:')
    .replace(/^[#* \t]*(?:Language|Bahasa)[* \t]*:/gim, 'Language:')
    .replace(/^[#* \t]*(?:Details|Detail|Rincian)[* \t]*:/gim, 'Details:')
    .replace(/^[#* \t]*(?:Stages|Tahapan|Langkah)[* \t]*:/gim, 'Stages:')
    .replace(/^[#* \t]*(?:Data)[* \t]*:/gim, 'Data:')
    .replace(/^[#* \t]*(?:Constraints|Batasan)[* \t]*:/gim, 'Constraints:')
  if (cleaned.length <= 25) return null

  // If the model emitted a substantive prompt without the explicit Task: or Details: anchors,
  // wrap it gracefully into an envelope rather than throwing away the synthesized output.
  if (!/^Task:/im.test(cleaned)) {
    // If the thread agreed on stages and the model omitted them, reject so fallback can supply the stages
    if (expectStages && !/^Stages:/im.test(cleaned)) return null
    cleaned = `Task: ${cleaned.split('\n')[0]}\nDetails:\n- ${cleaned}`
  } else if (!/^Details:/im.test(cleaned)) {
    cleaned = cleaned.replace(/^Task:([^\n]*)/m, 'Task:$1\nDetails:\n- $1')
  }

  // Envelope gate: Task + Details headings required; Stages required when the
  // thread agreed steps (the fallback always embeds them, so rejection here
  // degrades to a prompt that carries the stages).
  if (!/^Task:/im.test(cleaned) || !/^Details:/im.test(cleaned)) return null
  if (expectStages && !/^Stages:/im.test(cleaned)) return null
  const ordered = Array.isArray(editors)
    ? [
        ...new Set(
          editors.filter((e): e is TargetAppKind =>
            typeof e === 'string' ? (KNOWN_APPS as string[]).includes(e) : false,
          ),
        ),
      ]
    : []
  if (ordered.length === 0) return null
  return { prompt: cleaned, editors: ordered.slice(0, 3) }
}

/**
 * Handover envelope: the fixed section shape every target editor receives as
 * its first user message. English headings are machine-stable anchors across
 * id/en bodies; Data/Stages/Constraints appear only when the thread provides
 * them (never invent values to fill a section).
 *
 *   Task: <one imperative sentence: artifact + topic>
 *   Language: <Indonesian | English>
 *   Details:
 *   - <each agreed constraint>
 *   Data:
 *   - <concrete values stated in chat>
 */
interface HandoverEnvelopeSections {
  goal: string
  lang: Lang
  details: string[]
  data?: string[]
  stages?: string[]
  constraints?: string[]
}

function buildHandoverEnvelope(sections: HandoverEnvelopeSections): string {
  const lines = [
    `Task: ${taskLine(sections.lang, sections.goal)}`,
    `Language: ${LANGUAGE_NAMES[sections.lang]}`,
    'Details:',
  ]
  for (const detail of sections.details) lines.push(`- ${detail}`)
  if (sections.data && sections.data.length > 0) {
    lines.push('Data:')
    for (const datum of sections.data) lines.push(`- ${datum}`)
  }
  if (sections.stages && sections.stages.length > 0) {
    lines.push('Stages:')
    sections.stages.forEach((stage, i) => lines.push(`${i + 1}. ${stage}`))
  }
  if (sections.constraints && sections.constraints.length > 0) {
    lines.push('Constraints:')
    for (const constraint of sections.constraints) lines.push(`- ${constraint}`)
  }
  return lines.join('\n')
}

/**
 * Deterministic fallback when offline, unconfigured, or on timeout: the same
 * envelope the LLM synthesizer emits, built from the thread goal plus outline
 * stages. Data/Constraints are omitted (nothing verified to put in them).
 * Editors fall back to all five — the gate never disappears.
 */
/**
 * Gate: the thread only becomes an editor draft once the user states a topic.
 * A lone generic ask ("simple flowchart", "dokumen baru") gets no handover —
 * the omnibar answer already asks for the topic, and a draft from it can only
 * be hallucinated stages around a bare label (the observed empty-draft bug).
 * Additionally, pure Q&A / explanation threads (no creation intent detected)
 * never show a handover card, so "why did the handover trigger?" stays clean.
 * Returns the merged goal when actionable, null when the thread is chatter,
 * topic-less, or purely analytical (no creation intent).
 */
export function actionableHandoverGoal(chatMessages: OmnibarMessage[]): string | null {
  // Q&A-only threads: suppress the handover card entirely.
  // A thread is creation-intent-free when none of the user messages contain
  // creation verbs or document artifact nouns.
  if (!hasCreationIntent(chatMessages)) return null

  const { primaryGoal, refinedTopic } = extractSubstantiveUserContext(chatMessages)
  if (!primaryGoal || isGenericGoal(primaryGoal) || isChatterMessage(primaryGoal)) {
    if (!refinedTopic || isGenericGoal(refinedTopic) || isChatterMessage(refinedTopic)) return null
    return refinedTopic
  }
  if (!refinedTopic) {
    return primaryGoal
  }
  const normPrimary = primaryGoal.toLowerCase()
  const normRefined = refinedTopic.toLowerCase()
  const restatement =
    normRefined === normPrimary ||
    normRefined.includes(normPrimary) ||
    normPrimary.includes(normRefined)
  if (restatement) {
    const goal = refinedTopic.length >= primaryGoal.length ? refinedTopic : primaryGoal
    return !goal || isGenericGoal(goal) || isChatterMessage(goal) ? null : goal
  }
  // Latest substantive intent wins; only keep the "(first)" suffix when both
  // carry distinct info. A generic first message ("simple flowchart") or a
  // restatement (one contains the other) collapses to the longer form.
  if (isGenericGoal(primaryGoal) || isChatterMessage(primaryGoal)) {
    return isGenericGoal(refinedTopic) || isChatterMessage(refinedTopic) ? null : refinedTopic
  }
  if (isGenericGoal(refinedTopic) || isChatterMessage(refinedTopic)) {
    return primaryGoal
  }
  return `${refinedTopic} (${primaryGoal})`
}

export function fallbackHandover(
  chatMessages: OmnibarMessage[],
  uiLang = 'en',
): Pick<Handover, 'prompt' | 'editors'> | null {
  const goal = actionableHandoverGoal(chatMessages)
  // No topic yet: no prompt, no draft. The caller drops the handover dock.
  if (!goal) return null
  const lang = resolveHandoverLang(chatMessages, uiLang)
  const stages = extractOutlineStages(lastAssistantText(chatMessages)).map((s) => s.label)
  const { refinedTopic } = extractSubstantiveUserContext(chatMessages)
  const details = [`${BRIEF_LABELS[lang]}: ${goal}.`]
  if (refinedTopic && !goal.toLowerCase().includes(refinedTopic.toLowerCase())) {
    details.push(`Topic: ${refinedTopic}`)
  }
  const prompt = buildHandoverEnvelope({
    goal,
    lang,
    details,
    stages: stages.length > 0 ? stages : undefined,
  })
  return { prompt, editors: [...KNOWN_APPS] }
}
function buildHandoverSystemPrompt(lang: Lang): string {
  const display = LANGUAGE_NAMES[lang]
  const isIndo = lang === 'id'
  // Example rendered by the same builder the fallback uses, so the shape the
  // model imitates is byte-identical to the deterministic path.
  const example = JSON.stringify({
    prompt: buildHandoverEnvelope({
      goal: isIndo ? 'flowchart pendaftaran siswa baru' : 'approval workflow flowchart',
      lang,
      details: [
        `${BRIEF_LABELS[lang]}: ${isIndo ? 'flowchart pendaftaran siswa baru.' : 'approval workflow flowchart.'}`,
      ],
      stages: isIndo
        ? [
            'Masukkan data — formulir terisi lengkap',
            'Proses data — validasi kelengkapan',
            'Tampilkan hasil — status diterima atau ditolak',
          ]
        : [
            'Submit request — form complete',
            'Check completeness — all fields valid',
            'Manager decision — approved or rejected',
          ],
    }),
    editors: ['diagram'],
  })
  return `You are the ARAOffice handover synthesizer. Read the chat thread and decide where it executes next. Respond with ONLY one JSON object, no prose, no fences: {"prompt": "...", "editors": [...]}.
- prompt: one mature instruction in ${display}, agnostic to the target app. The target editor never sees this chat, so a bare label is a FAILURE. Use this exact shape — English headings always, body in ${display}:
Task: <one imperative sentence: artifact + topic>
Language: <the thread's language, e.g. ${display}>
Details:
- <each concrete constraint agreed in chat: goal, topic, audience, length, tone>
Data:
- <concrete values ONLY when the thread states them: columns, figures, node lists. OMIT this section when none — never invent numbers>
Stages:
1. <step> — <what done looks like>
Constraints:
- <thread-specific scope fence, e.g. scope: ONLY the named entity. OMIT when none>
Strip markdown (**, __) from stage names; keep decision points as questions in execution order. Never emit code fences or Mermaid.
Example: ${example}
- editors: 1-3 of ["docs", "sheets", "slides", "markdown", "diagram"], best first. Empty array when the thread is pure chatter.
- Never append meta-constraints about Mermaid, code style, or web searching — the target editor owns its own rendering and research policy.`
}

/**
 * Synthesizes the handover payload for the target editors using one LLM call,
 * with a 45s budget matching the 3-model fallback chain, then the deterministic fallback.
 */
export async function synthesizeHandover(
  chatMessages: OmnibarMessage[],
  uiLang = 'en',
  timeoutMs = 45_000,
): Promise<Handover | null> {
  const lang = resolveHandoverLang(chatMessages, uiLang)
  const outline = extractOutlineStages(lastAssistantText(chatMessages))
  // Topic-less thread (generic ask, chatter): no draft at all, not even a
  // fallback envelope — the omnibar answer already asks for the topic.
  if (!actionableHandoverGoal(chatMessages)) return null
  const fallback = (): Handover | null => {
    const fb = fallbackHandover(chatMessages, uiLang)
    return fb ? { ...fb, outline } : null
  }
  // If no AI bridge or no user messages, use the deterministic fallback.
  if (
    typeof window === 'undefined' ||
    !window.aiOfficeAi ||
    !chatMessages.some((m) => m.role === 'user' && m.text?.trim())
  ) {
    return fallback()
  }

  // Format multi-turn chat history into compact context (bounded per-turn so a
  // long thread cannot blow the synthesis request).
  const conversationSummary = chatMessages
    .slice(-8)
    .map((m) => {
      const text = m.text.length > 2000 ? `${m.text.slice(0, 2000)}…` : m.text
      return `[${m.role === 'user' ? 'User' : 'ARA Assistant'}]: ${text}`
    })
    .join('\n\n')

  try {
    const settings = await window.aiOfficeAi.getSettings()

    let accumulated = ''
    const parsed = await new Promise<{ prompt: string; editors: TargetAppKind[] } | null>(
      (resolve) => {
        let resolved = false
        let handle: { cancel: () => void } | null = null
        const settle = (value: { prompt: string; editors: TargetAppKind[] } | null) => {
          if (resolved) return
          resolved = true
          clearTimeout(fallbackTimer)
          resolve(value)
        }

        const fallbackTimer = setTimeout(() => {
          console.warn(
            `[Handover] Synthesis timed out after ${timeoutMs}ms; falling back to deterministic envelope`,
          )
          try {
            handle?.cancel()
          } catch {
            /* ignore */
          }
          settle(null)
        }, timeoutMs)
        // Don't keep the shell alive on a hung synthesis timer.
        fallbackTimer.unref?.()

        try {
          const synthApp: AiStreamRequest['app'] = 'docs'
          handle = window.aiOfficeAi.stream(
            {
              app: synthApp,
              settings,
              system: buildHandoverSystemPrompt(lang),
              messages: [
                {
                  role: 'user',
                  text: `Chat thread:\n${conversationSummary}\n\nTask: Synthesize the handover JSON for this thread.`,
                },
              ],
            },
            {
              onChunk(chunk) {
                if (resolved) return
                if (chunk.type === 'delta' && chunk.text) {
                  accumulated += chunk.text
                } else if (chunk.type === 'done') {
                  const res = parseHandoverJson(accumulated, outline.length > 0)
                  if (!res) {
                    console.warn(
                      '[Handover] LLM output rejected by envelope validator, falling back:',
                      accumulated.slice(0, 200),
                    )
                  }
                  settle(res)
                } else if (chunk.type === 'error') {
                  console.warn(
                    '[Handover] Stream error during synthesis, falling back:',
                    chunk.error,
                  )
                  settle(null)
                }
              },
            },
          )
        } catch (err) {
          console.warn('[Handover] Failed to initiate stream call:', err)
          settle(null)
        }
      },
    )

    if (parsed) return { ...parsed, outline }
    return fallback()
  } catch (err) {
    console.warn('[Handover] Unexpected error in synthesizeHandover:', err)
    return fallback()
  }
}
