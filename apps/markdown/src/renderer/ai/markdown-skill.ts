import { buildAppSystemPrompt, type AgentSkill } from '@araoffice/agent-core'
import systemSection from './prompts/system.md?raw'
import { MARKDOWN_TOOLS, buildMarkdownContext, executeMarkdownTool, markDocSeen } from './tools'
import type { MarkdownSkillDeps } from './tools'

const SYSTEM_PROMPT = buildAppSystemPrompt(systemSection)

export function createMarkdownSkill(getEditor: () => MarkdownSkillDeps['editor']): AgentSkill {
  return {
    id: 'markdown',
    systemPrompt: SYSTEM_PROMPT,
    tools: MARKDOWN_TOOLS,
    buildContext: () => {
      const editor = getEditor()
      if (!editor) return ''
      markDocSeen(editor)
      return buildMarkdownContext(editor)
    },
    executeTool: (call) => executeMarkdownTool({ editor: getEditor() }, call),
  }
}
