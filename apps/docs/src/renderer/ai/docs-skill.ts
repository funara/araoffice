import { buildAppSystemPrompt, type AgentSkill } from '@araoffice/agent-core'
import type { Editor } from '@tiptap/core'
import systemSection from './prompts/system.md?raw'
import { DOCS_TOOLS, executeDocsTool, markDocSeen } from './tools'
import { buildDocsContext, type AiTrack, type NumIds } from './protocol'

const SYSTEM_PROMPT = buildAppSystemPrompt(systemSection)

/** Capability surface the panel provides to the docs skill. */
export interface DocsSkillDeps {
  getEditor: () => Editor
  getNumIds: () => NumIds
  getTrack?: () => AiTrack | undefined
}

/**
 * The docs capability as an AgentSkill: document skeleton context, the five
 * document tools, and the local executor.
 */
export function createDocsSkill(deps: DocsSkillDeps): AgentSkill {
  const { getEditor, getNumIds, getTrack } = deps
  return {
    id: 'docs',
    systemPrompt: SYSTEM_PROMPT,
    tools: DOCS_TOOLS,
    buildContext: () => {
      const editor = getEditor()
      markDocSeen(editor) // the context the model receives is the freshness baseline for index-addressed writes
      return buildDocsContext(editor)
    },
    executeTool: (call, signal) =>
      executeDocsTool(
        { editor: getEditor(), numIds: getNumIds(), track: getTrack?.(), signal },
        call,
      ),
  }
}
