import { buildAppSystemPrompt, type AgentSkill } from '@araoffice/agent-core'
import basePrompt from './prompts/system.md?raw'
import { SHEETS_TOOLS, buildSheetsContext, executeSheetsTool, type SheetsSkillDeps } from './tools'

/**
 * The workbook DSL as an AgentSkill: mirrors createDocsSkill's shape
 * (systemPrompt + tools + buildContext + executeTool) so it plugs into the
 * same packages/agent-core AgentLoop docs uses.
 *
 * Prompt layout: the always-loaded base prompt (prompts/system.md) stays small
 * — workflow, op catalog, cross-cutting discipline — while per-domain field
 * definitions and conventions live in prompts/guides/*.md, loaded on demand
 * via load_guide.
 */
export function createSheetsSkill(deps: SheetsSkillDeps): AgentSkill {
  return {
    id: 'sheets',
    systemPrompt: buildAppSystemPrompt(basePrompt),
    tools: SHEETS_TOOLS,
    buildContext: () => buildSheetsContext(deps),
    executeTool: (call) => executeSheetsTool(deps, call),
  }
}

/** Legacy name kept for App/tests: prefer createSheetsSkill. */
export const createWorkbookSkill = createSheetsSkill
