import type { ReactElement } from 'react'
import type { OutlineEditNode } from '../shared/ipc'
import { t } from './i18n/locale'

export interface OutlineNode {
  title: string
  bold?: boolean
  italic?: boolean
  dest?: unknown
  url?: string
  items?: OutlineNode[]
}

/** Operations address a node by its path of sibling indices; text ops prompt in App */
export type OutlineEditOp = 'up' | 'down' | 'rename' | 'delete' | 'addChild'

function Item({
  node,
  depth,
  onGo,
}: {
  node: OutlineNode
  depth: number
  onGo: (n: OutlineNode) => void
}): ReactElement {
  return (
    <>
      <button
        className="pdf-outline-item"
        style={{
          paddingLeft: 10 + depth * 14,
          fontWeight: node.bold ? 600 : 400,
          fontStyle: node.italic ? 'italic' : undefined,
        }}
        title={node.title}
        onClick={() => onGo(node)}
      >
        {node.title}
      </button>
      {node.items?.map((c, i) => (
        <Item key={i} node={c} depth={depth + 1} onGo={onGo} />
      ))}
    </>
  )
}

/** One edited row: the title stays the jump target, edits sit beside it */
function EditItem({
  node,
  path,
  onEdit,
  onGo,
}: {
  node: OutlineEditNode
  path: number[]
  onEdit: (path: number[], op: OutlineEditOp) => void
  onGo: (n: OutlineNode) => void
}): ReactElement {
  const actions: [OutlineEditOp, string, string][] = [
    ['up', '↑', 'outlineUp'],
    ['down', '↓', 'outlineDown'],
    ['rename', '✎', 'outlineRename'],
    ['delete', '✕', 'outlineDelete'],
    ['addChild', '+', 'outlineAddChild'],
  ]
  return (
    <>
      <div className="pdf-outline-row" style={{ paddingLeft: 6 + path.length * 12 }}>
        <button className="pdf-outline-item" title={node.title} onClick={() => onGo(node)}>
          {node.title}
        </button>
        <span className="pdf-outline-actions">
          {actions.map(([op, glyph, key]) => (
            <button
              key={op}
              className="pdf-outline-tool"
              title={t(key as 'outlineUp')}
              onClick={() => onEdit(path, op)}
            >
              {glyph}
            </button>
          ))}
        </span>
      </div>
      {node.children?.map((child, i) => (
        <EditItem key={i} node={child} path={[...path, i]} onEdit={onEdit} onGo={onGo} />
      ))}
    </>
  )
}

/**
 * Outline (bookmark) tree: click jumps to internal destinations; url entries open external links.
 * With `draft` and `onEdit` set, the panel renders that tree instead — the edit surface of the
 * bookmark editor, whose state and prompting live in App.
 */
export function OutlinePanel({
  outline,
  onGoToDest,
  draft,
  onEdit,
}: {
  outline: OutlineNode[]
  onGoToDest: (dest: unknown) => void
  /** Present while the bookmark editor is open: the tree being edited */
  draft?: OutlineEditNode[]
  onEdit?: (path: number[], op: OutlineEditOp) => void
}): ReactElement {
  const onGo = (n: OutlineNode) => {
    // The url itself is never opened here: letting the window open reach the main process lets the
    // single suite-wide gate decide (see PDFChannels' setWindowOpenHandler in pdf-main)
    if (n.url) window.open(n.url, '_blank')
    else if (n.dest != null) onGoToDest(n.dest)
  }
  if (draft && onEdit) {
    return (
      <div className="pdf-outline">
        {draft.map((node, i) => (
          <EditItem key={i} node={node} path={[i]} onEdit={onEdit} onGo={onGo} />
        ))}
        {draft.length === 0 && <div className="pdf-outline-empty">{t('outlineAdd')}</div>}
      </div>
    )
  }
  return (
    <div className="pdf-outline">
      {outline.map((n, i) => (
        <Item key={i} node={n} depth={0} onGo={onGo} />
      ))}
    </div>
  )
}
