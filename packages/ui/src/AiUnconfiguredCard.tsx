/**
 * Shown instead of the AI composer while no provider is configured.
 *
 * ARAOffice has no accounts and no sign-in: the user configures a provider
 * (a Gemini API key, or any OpenAI-compatible endpoint) and the AI panel
 * becomes usable. The card points at that one action.
 */

export interface AiUnconfiguredCardProps {
  /** open the provider setup. Defaults to the 'ara:configure-ai' window event. */
  readonly onConfigure?: (() => void) | undefined
  /** override the body copy (localized) */
  readonly body?: string | undefined
  /** override the button label (localized) */
  readonly actionLabel?: string | undefined
}

/** Fired on `window` so a host (the shell) can open its provider settings. */
export const AI_CONFIGURE_EVENT = 'ara:configure-ai'

export function AiUnconfiguredCard({
  onConfigure,
  body = 'To start using ARA Assistant, add a free Gemini API key from Google AI Studio or connect your local model (Ollama / LM Studio).',
  actionLabel = 'Configure AI Provider',
}: AiUnconfiguredCardProps) {
  const configure = (): void => {
    if (onConfigure) {
      onConfigure()
      return
    }
    window.dispatchEvent(new CustomEvent(AI_CONFIGURE_EVENT))
  }

  return (
    <div className="ai-unconfigured-card">
      <div className="ai-unconfigured-mark" aria-hidden>
        ✦
      </div>
      <div className="ai-unconfigured-title">ARA Assistant</div>
      <div className="ai-unconfigured-body">{body}</div>
      <button className="ai-unconfigured-button" type="button" onClick={configure}>
        {actionLabel}
      </button>
    </div>
  )
}
