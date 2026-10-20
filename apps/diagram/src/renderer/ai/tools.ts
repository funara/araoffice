import type { AgentToolCall, AgentToolDef, ToolExecution } from '@araoffice/agent-core'
import {
  layoutNodesWithoutOverlap,
  validateAndSanitizeDrawio,
  validateAndSanitizeMermaid,
} from './diagram-guardrail'

const DIAGRAM_SHAPE_TYPES = [
  'rectangle',
  'rounded',
  'ellipse',
  'doubleEllipse',
  'diamond',
  'cylinder',
  'triangle',
  'hexagon',
  'cloud',
  'text',
] as const
type DiagramAiShapeType = (typeof DIAGRAM_SHAPE_TYPES)[number]

function isDiagramShapeType(value: unknown): value is DiagramAiShapeType {
  return typeof value === 'string' && DIAGRAM_SHAPE_TYPES.includes(value as DiagramAiShapeType)
}

function toolInputError(output: string): ToolExecution {
  return { output, summary: 'Invalid tool input', isError: true }
}

export interface DiagramSkillDeps {
  getMode(): 'visual' | 'mermaid'
  getFileName(): string
  getMermaidCode(): string
  setMermaidCode(code: string): void
  addVisualShape(shape: {
    type: DiagramAiShapeType
    label: string
    x?: number
    y?: number
    width?: number
    height?: number
    fillColor?: string
    strokeColor?: string
  }): { id: string; label: string }
  connectVisualShapes(sourceId: string, targetId: string, label?: string): { id: string }
  layoutVisualDiagram?(direction: 'horizontal' | 'vertical' | 'grid'): void
  styleVisualShapes?(
    cellIds: string[],
    style: { fillColor?: string; strokeColor?: string; strokeWidth?: number },
  ): void
  loadDrawioXml?(xml: string): Promise<void>
  getVisualSummary(): {
    vertexCount: number
    edgeCount: number
    nodes: Array<{ id: string; label: string; type?: string }>
  }
  clearDiagram(): void
  captureSnapshot?(label: string): void
  exportDiagram(format: 'svg' | 'png'): Promise<void>
}

export const DIAGRAM_TOOLS: AgentToolDef[] = [
  {
    name: 'get_diagram_info',
    description:
      'Get current diagram mode (visual Draw.io or mermaid diagram-as-code), file name, node/edge count, and structural summary.',
    inputSchema: {
      type: 'object',
      properties: {},
    },
  },

  {
    name: 'generate_mermaid',
    description:
      'Generate or update a complete Mermaid diagram (flowchart, sequenceDiagram, classDiagram, erDiagram, stateDiagram-v2, mindmap, architecture) only while Mermaid Code mode is active. Never change the editor mode automatically.',
    inputSchema: {
      type: 'object',
      properties: {
        code: {
          type: 'string',
          description: 'The complete valid Mermaid diagram syntax',
        },
      },
      required: ['code'],
    },
  },
  {
    name: 'generate_drawio_diagram',
    description:
      'Generate a complete visual diagram in Draw.io mode from an array of nodes and connections with automatic collision avoidance and clean spacing. Use only while Visual mode is active; never change the editor mode automatically.',
    inputSchema: {
      type: 'object',
      properties: {
        nodes: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              id: {
                type: 'string',
                description: 'Unique identifier for the node',
              },
              label: { type: 'string', description: 'Text label of the shape' },
              type: {
                type: 'string',
                enum: [...DIAGRAM_SHAPE_TYPES],
                description: 'Shape type',
              },
              fillColor: {
                type: 'string',
                description: 'Hex fill color (e.g. #dae8fc)',
              },
              strokeColor: {
                type: 'string',
                description: 'Hex border color (e.g. #6c8ebf)',
              },
            },
            required: ['label'],
          },
          description: 'List of nodes to create',
        },
        connections: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              source: {
                type: 'string',
                description: 'Source node id or label',
              },
              target: {
                type: 'string',
                description: 'Target node id or label',
              },
              label: {
                type: 'string',
                description: 'Optional text label on connector',
              },
            },
            required: ['source', 'target'],
          },
          description: 'List of connectors between nodes',
        },
        layout: {
          type: 'string',
          enum: ['horizontal', 'vertical', 'grid'],
          description:
            'Arrangement layout: horizontal for left-to-right flows, vertical for layered flows, grid for unconnected collections. Connected diagrams default to vertical.',
        },
      },
      required: ['nodes'],
    },
  },
  {
    name: 'add_shape',
    description:
      'In Visual Draw.io mode, add a new shape to the canvas. Never change the editor mode automatically.',
    inputSchema: {
      type: 'object',
      properties: {
        type: {
          type: 'string',
          enum: [...DIAGRAM_SHAPE_TYPES],
          description: 'Shape type',
        },
        label: {
          type: 'string',
          description: 'Label / text inside the shape',
        },
        x: {
          type: 'number',
          description: 'Horizontal position in pixels (e.g. 100)',
        },
        y: {
          type: 'number',
          description: 'Vertical position in pixels (e.g. 100)',
        },
        width: { type: 'number', description: 'Width in pixels (e.g. 140)' },
        height: { type: 'number', description: 'Height in pixels (e.g. 60)' },
        fillColor: {
          type: 'string',
          description: 'Hex fill color (e.g. #dae8fc)',
        },
        strokeColor: {
          type: 'string',
          description: 'Hex border color (e.g. #6c8ebf)',
        },
      },
      required: ['type', 'label'],
    },
  },
  {
    name: 'connect_shapes',
    description:
      'In Visual Draw.io mode, connect two existing shapes by their IDs. Never change the editor mode automatically.',
    inputSchema: {
      type: 'object',
      properties: {
        sourceId: { type: 'string', description: 'ID of the source shape' },
        targetId: { type: 'string', description: 'ID of the target shape' },
        label: {
          type: 'string',
          description: 'Optional text label on the connector',
        },
      },
      required: ['sourceId', 'targetId'],
    },
  },
  {
    name: 'validate_diagram_syntax',
    description:
      'Dry-run validation check for Mermaid or Draw.io code without mutating the canvas.',
    inputSchema: {
      type: 'object',
      properties: {
        format: { type: 'string', enum: ['mermaid', 'drawio'] },
        content: {
          type: 'string',
          description: 'Code or XML string to validate',
        },
      },
      required: ['format', 'content'],
    },
  },
  {
    name: 'clear_canvas',
    description: 'Clear the current diagram so a fresh diagram can be drawn.',
    inputSchema: {
      type: 'object',
      properties: {},
    },
  },
]

function modeMismatch(
  deps: DiagramSkillDeps,
  expected: 'visual' | 'mermaid',
): ToolExecution | null {
  const actual = deps.getMode()
  if (actual === expected) return null
  const expectedLabel = expected === 'mermaid' ? 'ARA Diagram Code' : 'ARA Diagram'
  const actualLabel = actual === 'mermaid' ? 'ARA Diagram Code' : 'ARA Diagram'
  return {
    output: `The active mode is ${actualLabel}. This action requires ${expectedLabel} mode; open a diagram in that mode instead of changing the editor automatically.`,
    summary: `Wrong mode for ${expectedLabel}`,
    isError: true,
  }
}

export async function executeDiagramTool(
  deps: DiagramSkillDeps,
  call: AgentToolCall,
): Promise<ToolExecution> {
  const input = call.input ?? {}
  if (typeof input !== 'object' || input === null || Array.isArray(input)) {
    return toolInputError('Tool input must be an object.')
  }

  switch (call.name) {
    case 'get_diagram_info': {
      const mode = deps.getMode()
      const fileName = deps.getFileName() || 'Untitled'
      if (mode === 'mermaid') {
        const code = deps.getMermaidCode()
        return {
          output: `Mode: Mermaid (Diagram-as-Code)\nFile: ${fileName}\nSource lines: ${code.split('\n').length}\nCode snippet:\n\`\`\`mermaid\n${code.slice(0, 500)}\n\`\`\``,
          summary: 'Read diagram details',
        }
      } else {
        const summary = deps.getVisualSummary()
        return {
          output: `Mode: Visual (Draw.io Native)\nFile: ${fileName}\nVertices (${summary.vertexCount}): ${summary.nodes.map((n) => `[${n.id}: "${n.label}"]`).join(', ')}\nEdges: ${summary.edgeCount}`,
          summary: 'Read diagram details',
        }
      }
    }

    case 'generate_mermaid': {
      const blocked = modeMismatch(deps, 'mermaid')
      if (blocked) return blocked
      if (typeof input.code !== 'string') return toolInputError('Mermaid code must be a string.')
      const rawCode = input.code

      const guardResult = await validateAndSanitizeMermaid(rawCode)

      if (!guardResult.valid) {
        return {
          output: `Mermaid Syntax Error: ${guardResult.error}\nPlease revise the diagram syntax and try again.`,
          summary: 'Mermaid validation failed',
          isError: true,
        }
      }

      deps.captureSnapshot?.('Generate Mermaid diagram')
      deps.setMermaidCode(guardResult.code)
      return {
        output: `Generated valid Mermaid diagram (${guardResult.code.split('\n').length} lines). Preview updated.`,
        summary: 'Generated Mermaid diagram',
        mutated: true,
      }
    }

    case 'generate_drawio_diagram': {
      const blocked = modeMismatch(deps, 'visual')
      if (blocked) return blocked

      const rawNodes = input.nodes
      if (!Array.isArray(rawNodes)) return toolInputError('Nodes must be an array.')
      if (
        rawNodes.some((node) => node === null || typeof node !== 'object' || Array.isArray(node))
      ) {
        return toolInputError('Every node must be an object.')
      }
      const nodes = rawNodes as Array<Record<string, unknown>>

      const rawConnections = input.connections
      if (
        rawConnections !== undefined &&
        (!Array.isArray(rawConnections) ||
          rawConnections.some(
            (connection) =>
              connection === null || typeof connection !== 'object' || Array.isArray(connection),
          ))
      ) {
        return toolInputError('Connections must be an array of objects.')
      }
      const connections =
        rawConnections === undefined ? [] : (rawConnections as Array<Record<string, unknown>>)

      let layoutDirection: 'horizontal' | 'vertical' | 'grid' =
        connections.length > 0 ? 'vertical' : 'grid'
      if (input.layout !== undefined) {
        if (
          input.layout !== 'horizontal' &&
          input.layout !== 'vertical' &&
          input.layout !== 'grid'
        ) {
          return toolInputError('Layout must be horizontal, vertical, or grid.')
        }
        layoutDirection = input.layout
      }

      if (nodes.length === 0) {
        return {
          output: 'No nodes provided to generate diagram.',
          summary: 'Empty node list',
          isError: true,
        }
      }

      const nodeReferences = new Set<string>()
      for (let i = 0; i < nodes.length; i++) {
        const node = nodes[i]
        if (typeof node.label !== 'string') {
          return toolInputError(`Node ${i + 1} must include a string label.`)
        }
        if (node.id !== undefined && typeof node.id !== 'string') {
          return toolInputError(`Node ${i + 1} has an invalid ID.`)
        }
        if (node.type !== undefined && !isDiagramShapeType(node.type)) {
          return toolInputError(`Node ${i + 1} has an invalid shape type.`)
        }
        nodeReferences.add(node.label)
        if (typeof node.id === 'string') nodeReferences.add(node.id)
      }

      for (let i = 0; i < connections.length; i++) {
        const connection = connections[i]
        if (
          typeof connection.source !== 'string' ||
          !connection.source.trim() ||
          typeof connection.target !== 'string' ||
          !connection.target.trim()
        ) {
          return toolInputError(`Connection ${i + 1} must include source and target strings.`)
        }
        if (connection.label !== undefined && typeof connection.label !== 'string') {
          return toolInputError(`Connection ${i + 1} has an invalid label.`)
        }
        if (!nodeReferences.has(connection.source) || !nodeReferences.has(connection.target)) {
          return toolInputError(
            `Connection references an unknown node: ${connection.source} -> ${connection.target}.`,
          )
        }
      }

      deps.captureSnapshot?.('Generate Draw.io diagram')
      const coords = layoutNodesWithoutOverlap(nodes.length, {
        direction: layoutDirection,
        startX: 100,
        startY: 80,
        nodeWidth: 140,
        nodeHeight: 60,
        gapX: 80,
        gapY: 60,
      })

      const idMap = new Map<string, string>()
      const labelMap = new Map<string, string>()

      for (let i = 0; i < nodes.length; i++) {
        const node = nodes[i]
        const coord = coords[i]
        const label = node.label as string
        const type = isDiagramShapeType(node.type) ? node.type : 'rounded'
        const res = deps.addVisualShape({
          type,
          label,
          x: coord.x,
          y: coord.y,
          width: coord.width,
          height: coord.height,
          fillColor: typeof node.fillColor === 'string' ? node.fillColor : '#dae8fc',
          strokeColor: typeof node.strokeColor === 'string' ? node.strokeColor : '#6c8ebf',
        })

        if (!res.id || res.id === '0') {
          return toolInputError('The visual canvas is not ready for drawing.')
        }
        if (typeof node.id === 'string' && node.id) idMap.set(node.id, res.id)
        labelMap.set(label, res.id)
      }

      let connCount = 0
      for (const connection of connections) {
        const source = connection.source as string
        const target = connection.target as string
        const sourceId = idMap.get(source) ?? labelMap.get(source)
        const targetId = idMap.get(target) ?? labelMap.get(target)
        if (!sourceId || !targetId) {
          return toolInputError(`Connection references an unknown node: ${source} -> ${target}.`)
        }
        const edge = deps.connectVisualShapes(
          sourceId,
          targetId,
          connection.label as string | undefined,
        )
        if (!edge.id || edge.id === '0') {
          return toolInputError(`Could not connect ${source} to ${target}.`)
        }
        connCount++
      }
      deps.layoutVisualDiagram?.(layoutDirection)

      return {
        output: `Generated Draw.io diagram with ${nodes.length} nodes and ${connCount} connectors without overlap.`,
        summary: `Created ${nodes.length} shapes & ${connCount} connectors`,
        mutated: true,
      }
    }

    case 'add_shape': {
      const blocked = modeMismatch(deps, 'visual')
      if (blocked) return blocked

      if (!isDiagramShapeType(input.type)) {
        return toolInputError('Shape type is invalid.')
      }
      if (typeof input.label !== 'string') {
        return toolInputError('Shape label must be a string.')
      }

      deps.captureSnapshot?.(`Add ${input.type}`)
      const res = deps.addVisualShape({
        type: input.type,
        label: input.label,
        x: typeof input.x === 'number' ? input.x : undefined,
        y: typeof input.y === 'number' ? input.y : undefined,
        width: typeof input.width === 'number' ? input.width : undefined,
        height: typeof input.height === 'number' ? input.height : undefined,
        fillColor: typeof input.fillColor === 'string' ? input.fillColor : undefined,
        strokeColor: typeof input.strokeColor === 'string' ? input.strokeColor : undefined,
      })
      if (!res.id || res.id === '0') {
        return toolInputError('The visual canvas is not ready for drawing.')
      }
      return {
        output: `Added ${input.type} shape with ID ${res.id} and label "${res.label}".`,
        summary: `Added ${input.type} shape: ${res.label}`,
        mutated: true,
      }
    }

    case 'connect_shapes': {
      const blocked = modeMismatch(deps, 'visual')
      if (blocked) return blocked

      if (
        typeof input.sourceId !== 'string' ||
        !input.sourceId.trim() ||
        typeof input.targetId !== 'string' ||
        !input.targetId.trim()
      ) {
        return toolInputError('Source and target IDs must be non-empty strings.')
      }
      if (input.label !== undefined && typeof input.label !== 'string') {
        return toolInputError('Connection label must be a string.')
      }

      const res = deps.connectVisualShapes(
        input.sourceId,
        input.targetId,
        input.label as string | undefined,
      )
      if (!res.id || res.id === '0') {
        return toolInputError('Could not connect the requested shapes.')
      }
      return {
        output: `Connected shape ${input.sourceId} to shape ${input.targetId} (edge ID ${res.id}).`,
        summary: 'Connected shapes',
        mutated: true,
      }
    }

    case 'validate_diagram_syntax': {
      if (input.format !== 'mermaid' && input.format !== 'drawio') {
        return toolInputError('Format must be mermaid or drawio.')
      }
      if (typeof input.content !== 'string') {
        return toolInputError('Diagram content must be a string.')
      }
      if (input.format === 'drawio') {
        const check = validateAndSanitizeDrawio(input.content)
        return {
          output: check.valid
            ? 'Draw.io XML is valid and secure.'
            : `Draw.io XML Error: ${check.error}`,
          summary: check.valid ? 'XML syntax valid' : 'XML syntax error',
          isError: !check.valid,
        }
      }

      const check = await validateAndSanitizeMermaid(input.content)
      return {
        output: check.valid
          ? 'Mermaid syntax is valid and secure.'
          : `Mermaid Syntax Error: ${check.error}`,
        summary: check.valid ? 'Mermaid syntax valid' : 'Mermaid syntax error',
        isError: !check.valid,
      }
    }

    case 'clear_canvas':
      deps.captureSnapshot?.('Clear canvas')
      deps.clearDiagram()
      return {
        output: 'Diagram cleared.',
        summary: 'Cleared canvas',
        mutated: true,
      }

    default:
      return {
        output: `Unknown tool: ${call.name}`,
        summary: call.name,
        isError: true,
      }
  }
}
