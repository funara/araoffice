import type { ReactNode } from 'react'

export interface QuickActionCard {
  readonly id: string
  readonly icon?: ReactNode | undefined
  readonly title: string
  readonly subtitle?: string | undefined
  readonly onClick: () => void
}

export interface ContextBadge {
  readonly icon?: ReactNode | undefined
  readonly label: string
  readonly title?: string | undefined
}

export interface AiWelcomeHubProps {
  readonly title?: string | undefined
  readonly description?: string | undefined
  readonly contextBadge?: ContextBadge | undefined
  readonly cards?: readonly QuickActionCard[] | undefined
  readonly children?: ReactNode | undefined
}

/**
 * Clean empty-state hub inside chatbox:
 * - Document/workspace context pill
 * - Rich 2-line quick action cards (icon + title + subtitle)
 */
export function AiWelcomeHub({
  title,
  description,
  contextBadge,
  cards,
  children,
}: AiWelcomeHubProps): React.JSX.Element {
  return (
    <div className="ai-welcome-hub">
      <div className="ai-hub-header">
        {title && <h4 className="ai-hub-title">{title}</h4>}
        {description && <p className="ai-hub-desc">{description}</p>}

        {contextBadge && (
          <div className="ai-hub-context-pill" title={contextBadge.title ?? contextBadge.label}>
            {contextBadge.icon}
            <span>{contextBadge.label}</span>
          </div>
        )}
      </div>

      {cards && cards.length > 0 && (
        <div className="ai-quick-grid">
          {cards.map((card) => (
            <button key={card.id} type="button" className="ai-quick-card" onClick={card.onClick}>
              {card.icon && <span className="quick-card-icon">{card.icon}</span>}
              <div className="quick-card-text">
                <strong className="quick-card-title">{card.title}</strong>
                {card.subtitle && <span className="quick-card-sub">{card.subtitle}</span>}
              </div>
            </button>
          ))}
        </div>
      )}

      {children}
    </div>
  )
}
