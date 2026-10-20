import React, { useEffect, useState } from 'react'

/** Some locales already end the label with an ellipsis — normalize to exactly one. */
function withEllipsis(label: string): string {
  return `${label.replace(/(?:…|\.{3})+$/u, '')}…`
}

/** Clean working/thinking indicator with elapsed time hint. */
export function AiTypingIndicator({ label }: { readonly label: string }): React.JSX.Element {
  const [elapsed, setElapsed] = useState(0)

  useEffect(() => {
    const timer = setInterval(() => setElapsed((s) => s + 1), 1000)
    return () => clearInterval(timer)
  }, [])

  return (
    <span className="ai-typing" role="status" aria-label={label}>
      <span className="ai-typing-label">
        {withEllipsis(label)}
        {elapsed >= 2 && <span className="ai-typing-elapsed" aria-hidden>{` · ${elapsed}s`}</span>}
      </span>
    </span>
  )
}
