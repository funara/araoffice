import { forwardRef, type HTMLAttributes, type ReactNode } from 'react'

export interface AiPanelShellProps extends HTMLAttributes<HTMLElement> {
  readonly children: ReactNode
}

export const AiPanelShell = forwardRef<HTMLElement, AiPanelShellProps>(function AiPanelShell(
  { children, className = '', ...props },
  ref,
) {
  return (
    <aside ref={ref} className={`ai-panel-shell${className ? ` ${className}` : ''}`} {...props}>
      {children}
    </aside>
  )
})
AiPanelShell.displayName = 'AiPanelShell'
