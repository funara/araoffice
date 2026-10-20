import { useEffect, useState } from 'react'
import type { HomeApi } from '../../shared/home-api'
import { useI18n } from './locale'

declare global {
  interface Window {
    aiOffice: HomeApi
  }
}

interface DonationModalProps {
  onClose: () => void
}

export function DonationModal({ onClose }: DonationModalProps) {
  const { t } = useI18n()
  const [vendor, setVendor] = useState('ARAOffice')
  useEffect(() => {
    void window.aiOffice
      .getSupportLinks()
      .then((links) => {
        if (links.vendorName) setVendor(links.vendorName)
      })
      .catch(() => {})
  }, [])

  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [onClose])
  return (
    <div className="modal-overlay donation-overlay" onClick={onClose}>
      <div
        className="modal donation-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t('supportTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        {/* Subtle Ambient Header Light */}
        <div className="donation-ambient-glow" aria-hidden="true" />

        {/* Header */}
        <header className="donation-header">
          <div className="donation-header-left">
            <div className="donation-eyebrow-pill">
              <span className="donation-heart-icon" aria-hidden="true">
                ❤️
              </span>
              <span>{t('supportTitle')}</span>
            </div>
            <h2 className="donation-title">{t('supportSubtitle')}</h2>
          </div>
          <button
            type="button"
            className="donation-close-btn"
            onClick={onClose}
            aria-label={t('cancel')}
            title={t('cancel')}
          >
            <svg
              width="16"
              height="16"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.2"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <line x1="18" y1="6" x2="6" y2="18" />
              <line x1="6" y1="6" x2="18" y2="18" />
            </svg>
          </button>
        </header>

        {/* Body */}
        <div className="donation-body">
          {/* QRIS Card Frame */}
          <div className="donation-qris-card">
            <img src="./sponsor.png" alt="QRIS Donation" className="qris-img" />
          </div>

          <p className="donation-desc">{t('supportDesc', { vendor })}</p>
        </div>
      </div>
    </div>
  )
}
