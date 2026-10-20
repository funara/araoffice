import { type Cell, Codec, type Graph } from '@maxgraph/core'

/** Decompress base64 raw DEFLATE bytes used by compressed Draw.io files */
async function decompressDiagramData(base64: string): Promise<string> {
  const binary = atob(base64.trim())
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i)
  }

  // Decompress via standard web Streams API
  const ds = new DecompressionStream('deflate-raw')
  const writer = ds.writable.getWriter()
  await writer.write(bytes)
  await writer.close()

  const text = await new Response(ds.readable).text()
  // URL decode if the inflated data was URL-encoded (standard Draw.io format)
  try {
    return decodeURIComponent(text)
  } catch {
    return text
  }
}

type DrawioSource = {
  document: XMLDocument
  kind: 'mxfile' | 'mxGraphModel'
  diagramIndex: number
  model: Element
}

const drawioSources = new WeakMap<Graph, DrawioSource>()

function directChild(parent: Element, name: string): Element | null {
  for (const child of Array.from(parent.children)) {
    if (child.nodeName === name) return child
  }
  return null
}

function parseModelXml(parser: DOMParser, raw: string): Element | null {
  const parsed = parser.parseFromString(raw, 'text/xml')
  if (parsed.querySelector('parsererror')) return null
  return parsed.documentElement.nodeName === 'mxGraphModel'
    ? parsed.documentElement
    : parsed.querySelector('mxGraphModel')
}

async function parseDrawioSource(xmlText: string): Promise<DrawioSource | null> {
  if (!xmlText || !xmlText.trim()) return null

  const parser = new DOMParser()
  const document = parser.parseFromString(xmlText.trim(), 'text/xml')
  if (document.querySelector('parsererror')) return null

  if (document.documentElement.nodeName === 'mxGraphModel') {
    return {
      document,
      kind: 'mxGraphModel',
      diagramIndex: -1,
      model: document.documentElement,
    }
  }

  if (document.documentElement.nodeName !== 'mxfile') return null

  const diagrams = Array.from(document.documentElement.getElementsByTagName('diagram'))
  for (let diagramIndex = 0; diagramIndex < diagrams.length; diagramIndex += 1) {
    const diagram = diagrams[diagramIndex]
    const directModel = directChild(diagram, 'mxGraphModel')
    if (directModel) {
      return { document, kind: 'mxfile', diagramIndex, model: directModel }
    }

    const rawContent = (diagram.textContent || '').trim()
    if (!rawContent) continue

    try {
      const model = rawContent.startsWith('<')
        ? parseModelXml(parser, rawContent)
        : parseModelXml(parser, await decompressDiagramData(rawContent))
      if (model) return { document, kind: 'mxfile', diagramIndex, model }
    } catch (err) {
      console.error('[diagram] Failed to decompress Draw.io diagram:', err)
    }
  }

  return null
}

/** Extract the <mxGraphModel> XML element from any Draw.io or mxGraph XML string */
export async function extractGraphModelNode(xmlText: string): Promise<Element | null> {
  return (await parseDrawioSource(xmlText))?.model || null
}

/** Convert a style object to Draw.io semicolon-delimited string key=value; */
export function formatStyleString(style: unknown): string {
  if (!style) return ''
  if (typeof style === 'string') return style
  if (typeof style !== 'object') return ''

  const parts: string[] = []
  for (const [k, v] of Object.entries(style as Record<string, unknown>)) {
    if (v === true) {
      parts.push(k)
    } else if (v !== false && v !== null && v !== undefined && v !== '') {
      parts.push(`${k}=${v}`)
    }
  }
  return parts.length > 0 ? `${parts.join(';')};` : ''
}

/**
 * Load Draw.io XML into a maxGraph instance.
 * Preserves nodes, edges, labels, geometry, and styles.
 */
export async function loadDrawioIntoGraph(graph: Graph, xmlText: string): Promise<boolean> {
  const source = await parseDrawioSource(xmlText)
  if (!source) return false

  const decodeDocument = globalThis.document.implementation.createDocument(
    null,
    'mxGraphModel',
    null,
  )
  const modelForDecode = decodeDocument.importNode(source.model, true) as Element
  const decodeRoot = directChild(modelForDecode, 'root')
  if (decodeRoot) {
    for (const child of Array.from(decodeRoot.children)) {
      if (!isGraphNode(child)) decodeRoot.removeChild(child)
    }
  }
  const codec = new Codec(decodeDocument)

  graph.batchUpdate(() => {
    graph.getDataModel().clear()
    codec.decode(modelForDecode, graph.getDataModel())
  })

  drawioSources.set(graph, source)
  graph.refresh()
  return true
}

const CELL_ATTRIBUTES = new Set([
  'id',
  'value',
  'style',
  'vertex',
  'edge',
  'parent',
  'source',
  'target',
  'connectable',
  'visible',
  'collapsed',
])
const GEOMETRY_ATTRIBUTES = new Set(['x', 'y', 'width', 'height', 'relative', 'as', 'offset'])

function graphNodeId(node: Element): string | null {
  if (node.nodeName === 'mxCell') return node.getAttribute('id')
  return directChild(node, 'mxCell')?.getAttribute('id') || node.getAttribute('id')
}

function isGraphNode(node: Element): boolean {
  return node.nodeName === 'mxCell' || directChild(node, 'mxCell') !== null
}

function preserveUnknownAttributes(from: Element, to: Element, known: Set<string>): void {
  for (const attribute of Array.from(from.attributes)) {
    if (!known.has(attribute.name) && !to.hasAttribute(attribute.name)) {
      to.setAttribute(attribute.name, attribute.value)
    }
  }
}

function preserveUnknownChildren(from: Element, to: Element, knownName: string): void {
  for (const child of Array.from(from.children)) {
    if (child.nodeName !== knownName) to.appendChild(to.ownerDocument.importNode(child, true))
  }
}

function mergeModeledElement(existing: Element, encoded: Element, known: Set<string>): Element {
  const next = encoded.cloneNode(true) as Element
  preserveUnknownAttributes(existing, next, known)

  const existingCell = existing.nodeName === 'mxCell' ? existing : directChild(existing, 'mxCell')
  const nextCell = next.nodeName === 'mxCell' ? next : directChild(next, 'mxCell')
  if (existingCell && nextCell) {
    preserveUnknownAttributes(existingCell, nextCell, CELL_ATTRIBUTES)
    const existingGeometry = directChild(existingCell, 'mxGeometry')
    const nextGeometry = directChild(nextCell, 'mxGeometry')
    if (existingGeometry && nextGeometry) {
      preserveUnknownAttributes(existingGeometry, nextGeometry, GEOMETRY_ATTRIBUTES)
      preserveUnknownChildren(existingGeometry, nextGeometry, 'Array')
    }
    preserveUnknownChildren(existingCell, nextCell, 'mxGeometry')
  }

  return next
}

function setCellAttributes(cellNode: Element, cell: Cell): void {
  cellNode.setAttribute('id', String(cell.id))
  if (cell.value != null && typeof cell.value !== 'object' && cell.value !== '') {
    cellNode.setAttribute('value', String(cell.value))
  }
  const style = formatStyleString(cell.style)
  if (style) cellNode.setAttribute('style', style)
  if (cell.vertex) cellNode.setAttribute('vertex', '1')
  if (cell.edge) cellNode.setAttribute('edge', '1')
  if (cell.parent) cellNode.setAttribute('parent', String(cell.parent.id))
  if (cell.source) cellNode.setAttribute('source', String(cell.source.id))
  if (cell.target) cellNode.setAttribute('target', String(cell.target.id))
  if (cell.connectable === false) cellNode.setAttribute('connectable', '0')
  if (cell.visible === false) cellNode.setAttribute('visible', '0')
  if (cell.collapsed) cellNode.setAttribute('collapsed', '1')
}

function appendGeometry(cellNode: Element, cell: Cell): void {
  const geometry = cell.geometry
  if (!geometry) return

  const geoNode = cellNode.ownerDocument.createElement('mxGeometry')
  if (geometry.x != null) geoNode.setAttribute('x', String(geometry.x))
  if (geometry.y != null) geoNode.setAttribute('y', String(geometry.y))
  if (geometry.width != null) geoNode.setAttribute('width', String(geometry.width))
  if (geometry.height != null) geoNode.setAttribute('height', String(geometry.height))
  if (geometry.relative) geoNode.setAttribute('relative', '1')
  if (geometry.offset) {
    const offset = cellNode.ownerDocument.createElement('mxPoint')
    offset.setAttribute('x', String(geometry.offset.x))
    offset.setAttribute('y', String(geometry.offset.y))
    offset.setAttribute('as', 'offset')
    geoNode.appendChild(offset)
  }
  if (geometry.points && geometry.points.length > 0) {
    const arrayNode = cellNode.ownerDocument.createElement('Array')
    arrayNode.setAttribute('as', 'points')
    for (const point of geometry.points) {
      const pointNode = cellNode.ownerDocument.createElement('mxPoint')
      pointNode.setAttribute('x', String(point.x))
      pointNode.setAttribute('y', String(point.y))
      arrayNode.appendChild(pointNode)
    }
    geoNode.appendChild(arrayNode)
  }
  geoNode.setAttribute('as', 'geometry')
  cellNode.appendChild(geoNode)
}

function encodeCell(cell: Cell, doc: XMLDocument): Element {
  const cellNode = doc.createElement('mxCell')
  setCellAttributes(cellNode, cell)
  appendGeometry(cellNode, cell)

  if (cell.value && typeof cell.value === 'object' && cell.value.nodeType === 1) {
    const objectNode = doc.importNode(cell.value, true) as Element
    const innerCell = directChild(objectNode, 'mxCell')
    if (innerCell) objectNode.removeChild(innerCell)
    objectNode.setAttribute('id', String(cell.id))
    objectNode.appendChild(cellNode)
    return objectNode
  }

  return cellNode
}

function updateGraphModel(modelNode: Element, graph: Graph, doc: XMLDocument): void {
  const cells = graph.getDataModel().cells || {}
  const encodedChildren = Object.keys(cells)
    .map((id) => cells[id])
    .filter((cell): cell is Cell => Boolean(cell && cell.id != null))
    .map((cell) => encodeCell(cell, doc))

  let root = directChild(modelNode, 'root')
  if (!root) {
    root = doc.createElement('root')
    modelNode.appendChild(root)
  }

  const existingById = new Map<string, Element>()
  for (const child of Array.from(root.children)) {
    const id = graphNodeId(child)
    if (id && isGraphNode(child)) existingById.set(id, child)
  }

  const encodedIds = new Set<string>()
  for (const encodedChild of encodedChildren) {
    const id = graphNodeId(encodedChild)
    if (!id) continue
    encodedIds.add(id)
    const existing = existingById.get(id)
    const next = existing
      ? mergeModeledElement(existing, encodedChild, CELL_ATTRIBUTES)
      : encodedChild
    if (existing) root.replaceChild(next, existing)
    else root.appendChild(next)
  }

  for (const child of Array.from(root.children)) {
    const id = graphNodeId(child)
    if (id && isGraphNode(child) && !encodedIds.has(id)) root.removeChild(child)
  }
}

function createDefaultDrawioDocument(): { document: XMLDocument; model: Element } {
  const document = globalThis.document.implementation.createDocument(null, 'mxfile', null)
  const mxfile = document.documentElement
  mxfile.setAttribute('host', 'Electron')
  mxfile.setAttribute('agent', 'ARAOffice')
  mxfile.setAttribute('version', '1.0.0')
  mxfile.setAttribute('type', 'device')

  const diagram = document.createElement('diagram')
  diagram.setAttribute('id', 'page-1')
  diagram.setAttribute('name', 'Page-1')
  mxfile.appendChild(diagram)

  const model = document.createElement('mxGraphModel')
  model.setAttribute('dx', '1422')
  model.setAttribute('dy', '798')
  model.setAttribute('grid', '1')
  model.setAttribute('gridSize', '10')
  model.setAttribute('guides', '1')
  model.setAttribute('tooltips', '1')
  model.setAttribute('connect', '1')
  model.setAttribute('arrows', '1')
  model.setAttribute('fold', '1')
  model.setAttribute('page', '1')
  model.setAttribute('pageScale', '1')
  model.setAttribute('pageWidth', '850')
  model.setAttribute('pageHeight', '1100')
  diagram.appendChild(model)
  return { document, model }
}

/**
 * Serialize maxGraph model into a Draw.io XML string.
 * Loaded mxfile documents are cloned and only their modeled graph page is updated.
 */
export function serializeGraphToDrawio(graph: Graph): string {
  const source = drawioSources.get(graph)
  let doc: XMLDocument
  let model: Element

  if (source?.kind === 'mxfile') {
    doc = source.document.cloneNode(true) as XMLDocument
    const diagrams = Array.from(doc.documentElement.getElementsByTagName('diagram'))
    const diagram = diagrams[source.diagramIndex]
    if (!diagram) return ''
    model = directChild(diagram, 'mxGraphModel') || (doc.importNode(source.model, true) as Element)
    if (!model.parentNode) {
      for (const child of Array.from(diagram.childNodes)) {
        if (child.nodeType === Node.TEXT_NODE) diagram.removeChild(child)
      }
      diagram.appendChild(model)
    }
  } else if (source?.kind === 'mxGraphModel') {
    const created = createDefaultDrawioDocument()
    doc = created.document
    model = doc.importNode(source.document.documentElement, true) as Element
    created.document.documentElement.firstElementChild?.replaceChildren(model)
  } else {
    const created = createDefaultDrawioDocument()
    doc = created.document
    model = created.model
  }

  updateGraphModel(model, graph, doc)
  return new XMLSerializer().serializeToString(doc)
}
