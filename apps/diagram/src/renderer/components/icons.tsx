import { AraSpark } from '@araoffice/ui'
import type { ReactElement, ReactNode } from 'react'

type IconProps = {
  size?: number
}

function SvgIcon({
  size = 16,
  strokeWidth = 1.6,
  children,
}: IconProps & { strokeWidth?: number; children: ReactNode }): ReactElement {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={strokeWidth}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      {children}
    </svg>
  )
}

export function IconUndo({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size} strokeWidth={1.5}>
      <path d="M5.91026 4L2.5 7.14791L5.91026 10.8205" />
      <path d="M3.96154 7.41028H15.1636C18.5169 7.41028 21.3646 10.1484 21.4953 13.5C21.6334 17.0416 18.707 20.0769 15.1636 20.0769H6.88384" />
    </SvgIcon>
  )
}

export function IconRedo({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size} strokeWidth={1.5}>
      <path d="M18.0897 4L21.5 7.14791L18.0897 10.8205" />
      <path d="M20.0385 7.41028H8.83636C5.4831 7.41028 2.63537 10.1484 2.5047 13.5C2.36657 17.0416 5.29296 20.0769 8.83636 20.0769H17.1162" />
    </SvgIcon>
  )
}

export function IconZoomIn({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <circle cx="11" cy="11" r="7.5" />
      <path d="m20 20-3.7-3.7M11 8v6M8 11h6" />
    </SvgIcon>
  )
}

export function IconZoomOut({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <circle cx="11" cy="11" r="7.5" />
      <path d="m20 20-3.7-3.7M8 11h6" />
    </SvgIcon>
  )
}

export function IconFit({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <path d="M9 3H3v6M15 3h6v6M21 15v6h-6M3 15v6h6" />
    </SvgIcon>
  )
}

export function IconDuplicate({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <rect x="9" y="9" width="12" height="12" rx="2" />
      <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
    </SvgIcon>
  )
}

export function IconDelete({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <path d="M4 7h16M10 11v6M14 11v6M6 7l1 14h10l1-14M9 7V4h6v3" />
    </SvgIcon>
  )
}

export function IconDownload({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <path d="M12 3v12M7 10l5 5 5-5M4 19v2h16v-2" />
    </SvgIcon>
  )
}

export function IconSave({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size} strokeWidth={1.5}>
      <path d="M3 4.5C3 3.67158 3.67158 3 4.5 3H17.1407L21 6.60325V19.5C21 20.3285 20.3285 21 19.5 21H4.5C3.67158 21 3 20.3285 3 19.5V4.5Z" />
      <path d="M12.0042 3L12 6.6923C12 6.86225 11.7761 7 11.5 7H7.5C7.22385 7 7 6.86225 7 6.6923V3H12.0042Z" />
      <path d="M7 13H17" />
      <path d="M7 17H12.0042" />
    </SvgIcon>
  )
}

export function IconCopy({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size} strokeWidth={1.6}>
      <rect x="9" y="9" width="13" height="13" rx="2" ry="2" />
      <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
    </SvgIcon>
  )
}

export function IconCode({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size} strokeWidth={1.6}>
      <polyline points="16 18 22 12 16 6" />
      <polyline points="8 6 2 12 8 18" />
    </SvgIcon>
  )
}

export function IconProperties({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <path d="M4 21v-7M4 10V3M12 21v-9M12 8V3M20 21v-5M20 12V3" />
      <path d="M1 14h6M9 8h6M17 16h6" />
    </SvgIcon>
  )
}

export function IconSelectAll({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <path d="M8 3H3v5M16 3h5v5M21 16v5h-5M3 16v5h5" />
      <rect x="8" y="8" width="8" height="8" rx="1" />
    </SvgIcon>
  )
}

export function IconConnector({ size = 18 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <circle cx="5" cy="6" r="2" />
      <circle cx="19" cy="18" r="2" />
      <path d="M7 7.5c3 1 5 2.8 6.5 5M13.5 15l3.5 2" />
      <path d="m15.5 13.5 1.5 3.5 3.5-1.5" />
    </SvgIcon>
  )
}

export function IconAutoLayout({ size = 22 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <rect x="3" y="3" width="6" height="6" rx="1" />
      <rect x="15" y="3" width="6" height="6" rx="1" />
      <rect x="9" y="15" width="6" height="6" rx="1" />
      <path d="M9 6h6M6 9v3.5l6 2.5M18 9v3.5l-6 2.5" />
    </SvgIcon>
  )
}

function AlignIcon({
  direction,
  size = 16,
}: {
  direction: 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom'
  size?: number
}): ReactElement {
  const horizontal = direction === 'left' || direction === 'center' || direction === 'right'
  const anchor =
    direction === 'left' || direction === 'top'
      ? 4
      : direction === 'right' || direction === 'bottom'
        ? 20
        : 12
  return (
    <SvgIcon size={size}>
      {horizontal ? (
        <>
          <path d={`M${anchor} 3v18`} />
          <path d="M7 7h9M7 12h6M7 17h11" />
        </>
      ) : (
        <>
          <path d={`M3 ${anchor}h18`} />
          <path d="M7 7v9M12 7v6M17 7v11" />
        </>
      )}
    </SvgIcon>
  )
}

export function IconAlignLeft({ size = 16 }: IconProps = {}): ReactElement {
  return <AlignIcon direction="left" size={size} />
}

export function IconAlignCenter({ size = 16 }: IconProps = {}): ReactElement {
  return <AlignIcon direction="center" size={size} />
}

export function IconAlignRight({ size = 16 }: IconProps = {}): ReactElement {
  return <AlignIcon direction="right" size={size} />
}

export function IconAlignTop({ size = 16 }: IconProps = {}): ReactElement {
  return <AlignIcon direction="top" size={size} />
}

export function IconAlignMiddle({ size = 16 }: IconProps = {}): ReactElement {
  return <AlignIcon direction="middle" size={size} />
}

export function IconAlignBottom({ size = 16 }: IconProps = {}): ReactElement {
  return <AlignIcon direction="bottom" size={size} />
}

export function IconDistributeHorizontal({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <path d="M4 4v16M20 4v16M8 7v10M16 7v10M4 12h16" />
    </SvgIcon>
  )
}

export function IconDistributeVertical({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <path d="M4 4h16M4 20h16M7 8h10M7 16h10M12 4v16" />
    </SvgIcon>
  )
}

export function IconBringFront({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <rect x="3" y="7" width="11" height="11" rx="1" />
      <path d="M10 3h10v10M10 3l-3 3M20 3l-3 3" />
    </SvgIcon>
  )
}

export function IconSendBack({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <rect x="10" y="6" width="11" height="11" rx="1" />
      <path d="M4 10V3h7M4 3l3 3M4 3l3-3" />
    </SvgIcon>
  )
}

export function IconGrid({ size = 18 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <path d="M4 4h4v4H4zM10 4h4v4h-4zM16 4h4v4h-4zM4 10h4v4H4zM10 10h4v4h-4zM16 10h4v4h-4zM4 16h4v4H4zM10 16h4v4h-4zM16 16h4v4h-4z" />
    </SvgIcon>
  )
}

export function IconText({ size = 18 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <path d="M5 4h14M12 4v16M8 20h8" />
    </SvgIcon>
  )
}

function ShapeIcon({
  shape,
  size = 18,
}: {
  shape:
    | 'rectangle'
    | 'rounded'
    | 'ellipse'
    | 'doubleEllipse'
    | 'diamond'
    | 'cylinder'
    | 'triangle'
    | 'hexagon'
    | 'cloud'
  size?: number
}): ReactElement {
  return (
    <SvgIcon size={size}>
      {shape === 'rectangle' && <rect x="4" y="5" width="16" height="14" />}
      {shape === 'rounded' && <rect x="4" y="5" width="16" height="14" rx="3" />}
      {shape === 'ellipse' && <ellipse cx="12" cy="12" rx="8" ry="6.5" />}
      {shape === 'doubleEllipse' && (
        <>
          <ellipse cx="12" cy="12" rx="8" ry="6.5" />
          <ellipse cx="12" cy="12" rx="6.5" ry="5" />
        </>
      )}
      {shape === 'diamond' && <path d="m12 3 9 9-9 9-9-9 9-9z" />}
      {shape === 'cylinder' && (
        <>
          <path d="M5 6c0-2 14-2 14 0v12c0 2-14 2-14 0z" />
          <path d="M5 6c0 2 14 2 14 0M5 12c0 2 14 2 14 0" />
        </>
      )}
      {shape === 'triangle' && <path d="m12 4 8 16H4L12 4z" />}
      {shape === 'hexagon' && <path d="m7 4 10 0 4 8-4 8H7l-4-8 4-8z" />}
      {shape === 'cloud' && (
        <path d="M7 18h10a4 4 0 0 0 .6-7.95A6 6 0 0 0 6.2 8.4 4.8 4.8 0 0 0 7 18Z" />
      )}
    </SvgIcon>
  )
}

export function IconRectangle({ size = 18 }: IconProps = {}): ReactElement {
  return <ShapeIcon shape="rectangle" size={size} />
}

export function IconRoundedRectangle({ size = 18 }: IconProps = {}): ReactElement {
  return <ShapeIcon shape="rounded" size={size} />
}

export function IconEllipse({ size = 18 }: IconProps = {}): ReactElement {
  return <ShapeIcon shape="ellipse" size={size} />
}
export function IconDoubleEllipse({ size = 18 }: IconProps = {}): ReactElement {
  return <ShapeIcon shape="doubleEllipse" size={size} />
}

export function IconDiamond({ size = 18 }: IconProps = {}): ReactElement {
  return <ShapeIcon shape="diamond" size={size} />
}

export function IconCylinder({ size = 18 }: IconProps = {}): ReactElement {
  return <ShapeIcon shape="cylinder" size={size} />
}

export function IconTriangle({ size = 18 }: IconProps = {}): ReactElement {
  return <ShapeIcon shape="triangle" size={size} />
}

export function IconHexagon({ size = 18 }: IconProps = {}): ReactElement {
  return <ShapeIcon shape="hexagon" size={size} />
}

export function IconCloud({ size = 18 }: IconProps = {}): ReactElement {
  return <ShapeIcon shape="cloud" size={size} />
}

export function AssistantMark({ size = 26 }: { size?: number }): ReactElement {
  return <AraSpark size={size} className="ara-assistant-badge" />
}

export function IconSparkles({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <path d="m12 3 1.9 5.8 5.8 1.9-5.8 1.9L12 20.5l-1.9-5.9-5.8-1.9 5.8-1.9zM19 3l.9 2.1L22 6l-2.1.9L19 9l-.9-2.1L16 6l2.1-.9zM5 16l.9 2.1L8 19l-2.1.9L5 22l-.9-2.1L2 19l2.1-.9z" />
    </SvgIcon>
  )
}

export function IconFlowchart({ size = 18 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <rect x="3" y="3" width="6" height="5" rx="1" />
      <rect x="15" y="16" width="6" height="5" rx="1" />
      <rect x="3" y="16" width="6" height="5" rx="1" />
      <path d="M6 8v4h12v4M6 12v4" />
    </SvgIcon>
  )
}

export function IconSequence({ size = 18 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <path d="M4 4v16M20 4v16M4 8h15m-3-3 3 3-3 3M20 16H5m3-3-3 3 3 3" />
    </SvgIcon>
  )
}

export function IconSchema({ size = 18 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <ellipse cx="12" cy="5" rx="8" ry="3" />
      <path d="M4 5v6c0 1.66 3.58 3 8 3s8-1.34 8-3V5M4 11v6c0 1.66 3.58 3 8 3s8-1.34 8-3v-6" />
    </SvgIcon>
  )
}
export function IconClear({ size = 16 }: IconProps = {}): ReactElement {
  return (
    <SvgIcon size={size}>
      <path d="M18 6 6 18M6 6l12 12" />
    </SvgIcon>
  )
}
