/**
 * Icon set for the markdown app. Everything that exists in the docs ribbon
 * library is re-exported from there so glyph style stays uniform across the
 * suite; the handful of markdown-only glyphs below are drawn on the same
 * 16-grid / pinned-stroke contract. (Long term this belongs in packages/ui.)
 */

import { Svg, type IconProps } from '../../../../docs/src/renderer/components/icons'

export {
  IconBullets,
  IconNumbered,
  IconIndentDec,
  IconIndentInc,
  IconTable,
  IconPicture,
  IconLink,
  IconHighlight,
  IconUndo,
  IconRedo,
  IconRowInsertAbove,
  IconRowInsertBelow,
  IconColInsertLeft,
  IconColInsertRight,
  IconRowDelete,
  IconColDelete,
  IconTableDelete,
  AssistantMark,
  Svg,
  pinnedStroke,
} from '../../../../docs/src/renderer/components/icons'

export function IconTaskList(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="1.8" y="2.2" width="4.4" height="4.4" rx="1" />
      <path d="M3.2 4.4l1 1 1.7-1.9" />
      <path d="M8.8 4.4h5.4" />
      <rect x="1.8" y="9.4" width="4.4" height="4.4" rx="1" />
      <path d="M8.8 11.6h5.4" />
    </Svg>
  )
}

export function IconHr(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M2 8h12" />
      <path d="M4.5 4.2h7M4.5 11.8h7" opacity="0.45" />
    </Svg>
  )
}

export function IconProperties(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M2.5 4.4h11M2.5 8h11M2.5 11.6h6.2" />
    </Svg>
  )
}

export function IconHeaderRow(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2" y="3" width="12" height="10" rx="1" />
      <path d="M2 6.4h12M6.7 6.4V13M11.3 6.4V13" />
      <path d="M2.6 3.6h10.8v2.3H2.6z" fill="currentColor" stroke="none" opacity="0.35" />
    </Svg>
  )
}

export function IconInlineCode(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M5.4 4.6L2.4 8l3 3.4M10.6 4.6l3 3.4-3 3.4" />
    </Svg>
  )
}

export function IconPrint(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M4.5 5.5V2.5h7v3M4.5 12H3a1.5 1.5 0 01-1.5-1.5V7.5a1.5 1.5 0 011.5-1.5h10a1.5 1.5 0 011.5 1.5v3a1.5 1.5 0 01-1.5 1.5h-1.5" />
      <rect x="4.5" y="9" width="7" height="4.5" rx="0.5" />
    </Svg>
  )
}

export function IconExport(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M8 2.5v7M5.5 7l2.5 2.5L10.5 7M3 10.5v2a1 1 0 001 1h8a1 1 0 001-1v-2" />
    </Svg>
  )
}

export function IconCodeBlock(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M5 5L2 8l3 3M11 5l3 3-3 3M9.5 3.5l-3 9" />
    </Svg>
  )
}

export function IconBlockquote(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3 4v8M6.5 6h6.5M6.5 9h4.5" />
    </Svg>
  )
}
