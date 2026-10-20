import React, { Component, type ErrorInfo, type ReactNode } from 'react'

export interface ErrorBoundaryProps {
  children: ReactNode
  fallback?: (error: Error, reset: () => void) => ReactNode
  onCatch?: (error: Error, info: ErrorInfo) => void
}

interface ErrorBoundaryState {
  hasError: boolean
  error: Error | null
}

export class ErrorBoundary extends Component<ErrorBoundaryProps, ErrorBoundaryState> {
  constructor(props: ErrorBoundaryProps) {
    super(props)
    this.state = { hasError: false, error: null }
  }

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { hasError: true, error }
  }

  override componentDidCatch(error: Error, errorInfo: ErrorInfo): void {
    this.props.onCatch?.(error, errorInfo)
    console.error('[ErrorBoundary caught error]:', error, errorInfo)
  }

  private handleReset = (): void => {
    this.setState({ hasError: false, error: null })
  }

  override render(): ReactNode {
    if (this.state.hasError && this.state.error) {
      if (this.props.fallback) {
        return this.props.fallback(this.state.error, this.handleReset)
      }
      return (
        <div
          style={{
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            height: '100%',
            width: '100%',
            padding: 24,
            boxSizing: 'border-box',
            color: '#94a3b8',
            fontFamily: 'system-ui, -apple-system, sans-serif',
            background: 'var(--bg-canvas, #18191c)',
            textAlign: 'center',
          }}
        >
          <div
            style={{
              padding: '24px 32px',
              borderRadius: 16,
              background: 'rgba(255, 255, 255, 0.04)',
              border: '1px solid rgba(255, 255, 255, 0.08)',
              maxWidth: 480,
            }}
          >
            <h3 style={{ margin: '0 0 8px', fontSize: 16, fontWeight: 600, color: '#f1f5f9' }}>
              Something went wrong
            </h3>
            <p style={{ margin: '0 0 20px', fontSize: 13, lineHeight: 1.5, color: '#94a3b8' }}>
              An unexpected render error occurred. You can reload this view to restore the session.
            </p>
            <div style={{ display: 'flex', gap: 10, justifyContent: 'center' }}>
              <button
                type="button"
                onClick={this.handleReset}
                style={{
                  padding: '8px 18px',
                  fontSize: 13,
                  fontWeight: 500,
                  borderRadius: 8,
                  border: 'none',
                  background: '#2367ec',
                  color: '#ffffff',
                  cursor: 'pointer',
                }}
              >
                Try Again
              </button>
              <button
                type="button"
                onClick={() => window.location.reload()}
                style={{
                  padding: '8px 18px',
                  fontSize: 13,
                  fontWeight: 500,
                  borderRadius: 8,
                  border: '1px solid rgba(255, 255, 255, 0.12)',
                  background: 'transparent',
                  color: '#e2e8f0',
                  cursor: 'pointer',
                }}
              >
                Reload Window
              </button>
            </div>
          </div>
        </div>
      )
    }
    return this.props.children
  }
}
