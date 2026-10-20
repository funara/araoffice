import { buildAppSystemPrompt, type AgentSkill } from '@araoffice/agent-core'
import systemSection from './prompts/system.md?raw'
import { DIAGRAM_TOOLS, executeDiagramTool } from './tools'
import type { DiagramSkillDeps } from './tools'

const SYSTEM_PROMPT = buildAppSystemPrompt(systemSection)

export function createDiagramSkill(getDeps: () => DiagramSkillDeps): AgentSkill {
  return {
    id: 'diagram',
    systemPrompt: SYSTEM_PROMPT,
    tools: DIAGRAM_TOOLS,
    buildContext: () => {
      const deps = getDeps()
      const mode = deps.getMode()
      const fileName = deps.getFileName() || 'Untitled'
      return `Current diagram: "${fileName}", active mode: ${mode}. Keep this mode; open a new diagram in another format for a different mode.`
    },
    executeTool: (call) => executeDiagramTool(getDeps(), call),
  }
}
