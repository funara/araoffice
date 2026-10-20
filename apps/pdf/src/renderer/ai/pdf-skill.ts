import { buildAppSystemPrompt, type AgentSkill } from '@araoffice/agent-core'
import systemSection from './prompts/system.md?raw'
import { PDF_TOOLS, executePdfTool } from './tools'
import type { PdfSkillDeps } from './tools'

const SYSTEM_PROMPT = buildAppSystemPrompt(systemSection)
export function createPdfSkill(deps: PdfSkillDeps): AgentSkill {
  return {
    id: 'pdf',
    systemPrompt: SYSTEM_PROMPT,
    tools: PDF_TOOLS,
    buildContext: () => {
      const parts = [
        `Current document: "${deps.fileName()}", ${deps.pageCount()} pages; the user is viewing page ${deps.currentPage()}.`,
        'Mode: the document can be edited from this chat (annotations, form fills, page operations); password-protected files stay read-only.',
      ]
      const outline = deps.outline()
      if (outline && outline.length > 0) {
        parts.push(
          `The document has an outline (${outline.length} top-level entries); use get_outline to view it.`,
        )
      }
      return parts.join('\n')
    },
    executeTool: (call) => executePdfTool(deps, call),
  }
}
