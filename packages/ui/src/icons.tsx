import type { ReactNode } from 'react'

export interface IconProps {
  size?: number
}

function Svg({ size = 16, children }: IconProps & { children: ReactNode }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.2"
      strokeLinecap="round"
      aria-hidden
    >
      {children}
    </svg>
  )
}

export function IconSend(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M2.2 8 13.8 2.6 11 13.4 7.6 9.6z" strokeLinejoin="round" />
      <path d="M7.6 9.6 13.8 2.6" />
    </Svg>
  )
}

/** compact upward arrow used by the AI composer send button */
export function IconArrowUp(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M8 13V3" strokeWidth="1.5" />
      <path d="m4 7 4-4 4 4" strokeWidth="1.5" />
    </Svg>
  )
}

/** compact paperclip used by the AI composer attachment button */
export function IconAttach(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="m6.3 9.7 4.6-4.6a2.1 2.1 0 0 1 3 3l-6.25 6.25a3.5 3.5 0 0 1-4.95-4.95l6.25-6.25a2.5 2.5 0 0 1 3.54 3.54l-4.6 4.6a1.25 1.25 0 0 1-1.77-1.77l4.6-4.6" />
    </Svg>
  )
}

export function IconStop(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4" y="4" width="8" height="8" rx="1.5" fill="currentColor" stroke="none" />
    </Svg>
  )
}

/** return/enter arrow (↵) for the icon-only send button */
export function IconEnter(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M13 3.5v4a2.5 2.5 0 0 1-2.5 2.5H3.5" />
      <path d="M6.5 7 3.5 10l3 3" />
    </Svg>
  )
}

export function IconNewChat(props: IconProps) {
  return (
    <Svg {...props}>
      <path
        d="M13.5 7.2v-3A1.7 1.7 0 0 0 11.8 2.5H4.2a1.7 1.7 0 0 0-1.7 1.7v6.1a1.7 1.7 0 0 0 1.7 1.7h1.1v2l2.6-2h1.3"
        strokeLinejoin="round"
      />
      <path d="M12.2 9.4v4M10.2 11.4h4" />
    </Svg>
  )
}

export function IconSidebarCollapse(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="1.5" y="2.5" width="13" height="11" rx="1" />
      <path d="M5.5 2.5v11" />
      <path d="M12.5 8H8.1M9.8 5.9 7.7 8l2.1 2.1" strokeWidth="1.3" strokeLinejoin="round" />
    </Svg>
  )
}

export function IconSidebarCollapseLeft(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="1.5" y="2.5" width="13" height="11" rx="1" />
      <path d="M10.5 2.5v11" />
      <path d="M3.5 8h4.4M6.2 5.9 8.3 8 6.2 10.1" strokeWidth="1.3" strokeLinejoin="round" />
    </Svg>
  )
}

export function IconClock(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="8" cy="8" r="5.5" />
      <polyline points="8 4.8 8 8 10.5 9.5" />
    </Svg>
  )
}

export function IconCopy(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="5.5" y="5.5" width="8" height="8" rx="1.5" />
      <path d="M10.5 3.5h-6A1.5 1.5 0 0 0 3 5v6" />
    </Svg>
  )
}

export function IconCheck(props: IconProps) {
  return (
    <Svg {...props}>
      <polyline points="3.5 8.5 6.5 11.5 12.5 4.5" strokeWidth="1.5" strokeLinejoin="round" />
    </Svg>
  )
}

export function IconRefresh(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M2.5 7.5A5.5 5.5 0 1 1 8 13.5c-2.4 0-4.4-1.5-5.1-3.6" />
      <path d="M2.5 3.5v4h4" strokeLinejoin="round" />
    </Svg>
  )
}

export function IconUndo(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3.5 6.5h7a3.5 3.5 0 1 1 0 7H7" />
      <path d="m6 4-2.5 2.5L6 9" strokeLinejoin="round" />
    </Svg>
  )
}

export function IconSparkles(props: IconProps) {
  return (
    <Svg {...props}>
      <path
        d="m8 1.5 1.2 5.3 5.3 1.2-5.3 1.2L8 14.5l-1.2-5.3-5.3-1.2 5.3-1.2z"
        strokeLinejoin="round"
      />
      <circle cx="12.5" cy="3.5" r="0.8" fill="currentColor" stroke="none" />
    </Svg>
  )
}

/** Animated ARA spark — same 4-point star as the omnibar (shell CommandCenterModal AraSpark). */
export function AraSpark({
  size = 16,
  live = false,
  className,
}: IconProps & { live?: boolean; className?: string }) {
  return (
    <svg
      className={`ara-spark${live ? ' ara-spark-live' : ''}${className ? ` ${className}` : ''}`}
      width={size}
      height={size}
      viewBox="0 0 16 16"
      fill="currentColor"
      aria-hidden="true"
    >
      <path d="M8 0c.7 4.5 3.5 7.3 8 8-4.5.7-7.3 3.5-8 8-.7-4.5-3.5-7.3-8-8 4.5-.7 7.3-3.5 8-8z" />
    </svg>
  )
}

export function IconJump(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="8" cy="8" r="5" />
      <circle cx="8" cy="8" r="1.5" fill="currentColor" stroke="none" />
      <path d="M8 1.5v2M8 12.5v2M1.5 8h2M12.5 8h2" />
    </Svg>
  )
}

export function IconChevronDown(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="m4 6 4 4 4-4" strokeWidth="1.4" strokeLinejoin="round" />
    </Svg>
  )
}

export function IconChevronRight(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="m6 4 4 4-4 4" strokeWidth="1.4" strokeLinejoin="round" />
    </Svg>
  )
}

export function IconExternalLink(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M12 8.5v4.5a1 1 0 0 1-1 1H3.5a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1H8" />
      <path d="M9.5 2.5H13.5v4" />
      <path d="M6.5 9.5 13.5 2.5" />
    </Svg>
  )
}
