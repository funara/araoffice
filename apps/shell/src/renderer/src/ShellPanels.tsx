import { useEffect, useState } from 'react'
import websiteIcon from './assets/website-icon.svg'
import instagramIcon from './assets/instagram-icon.svg'
import githubIcon from './assets/github-icon.svg'
import sponsorIcon from './assets/sponsor-icon.svg'
import { useI18n, type StringKey } from './locale'

type CapabilityIconName = 'analysis' | 'brain' | 'eye' | 'lightning' | 'feather'

type AraCopy = {
  label: StringKey
  description: StringKey
  bestFor: StringKey
  capabilities: ReadonlyArray<{ label: StringKey; icon: CapabilityIconName }>
}

/** One model, one card: ARA Assistant works with the provider the user configured. */
const ARA_COPY: AraCopy = {
  label: 'araLabel',
  description: 'araDescription',
  bestFor: 'araBestFor',
  capabilities: [
    { label: 'araAnalysis', icon: 'analysis' },
    { label: 'araReasoning', icon: 'brain' },
    { label: 'araVision', icon: 'eye' },
    { label: 'araLatency', icon: 'lightning' },
    { label: 'araTransform', icon: 'feather' },
  ],
}

function CapabilityIcon({ name }: { name: CapabilityIconName }) {
  return (
    <svg
      className="ai-model-capability-icon"
      width="13"
      height="13"
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.35"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      {name === 'analysis' ? (
        <>
          <circle cx="6.75" cy="6.75" r="4.25" />
          <path d="m9.75 9.75 4 4" />
          <path d="M5 6.75h3.5M6.75 5v3.5" />
        </>
      ) : name === 'brain' ? (
        <>
          <path d="M6.1 3.5a2.2 2.2 0 0 0-2.9 2.1 2.2 2.2 0 0 0 .2.9 2.2 2.2 0 0 0 .7 4.2 2 2 0 0 0 3.9.5" />
          <path d="M9.9 3.5a2.2 2.2 0 0 1 2.9 2.1 2.2 2.2 0 0 1-.2.9 2.2 2.2 0 0 1-.7 4.2 2 2 0 0 1-3.9.5" />
          <path d="M8 3.4v9.2M4.5 7.1H8M8 6h3.5M8 9.2h3" />
        </>
      ) : name === 'eye' ? (
        <>
          <path d="M1.8 8s2.2-3.4 6.2-3.4 6.2 3.4 6.2 3.4-2.2 3.4-6.2 3.4S1.8 8 1.8 8Z" />
          <circle cx="8" cy="8" r="1.5" />
        </>
      ) : name === 'lightning' ? (
        <path d="M9.4 1.8 3.7 8.3h3.4l-.6 5.9 5.8-7.3H9.1z" fill="currentColor" stroke="none" />
      ) : name === 'feather' ? (
        <>
          <path d="M13.3 2.7c-3.7-.1-6.7 1.1-8 3.6-1.2 2.3-.1 4.2 2.2 4.5 2.8.3 4.8-2.5 5.8-8.1Z" />
          <path d="M2.8 13.3c2.1-2.4 4.3-4.1 7.5-6.2" />
        </>
      ) : (
        <>
          <path d="M2.5 4h3v3h3v3h5" />
          <path d="M2.5 12h3" />
        </>
      )}
    </svg>
  )
}

function WebsiteIcon() {
  return (
    <img
      className="social-icon-website"
      src={websiteIcon}
      width="20"
      height="20"
      alt=""
      aria-hidden="true"
    />
  )
}
function InstagramIcon() {
  return <img src={instagramIcon} width="24" height="24" alt="" aria-hidden="true" />
}

function GitHubIcon() {
  return <img src={githubIcon} width="24" height="24" alt="" aria-hidden="true" />
}

function SponsorIcon() {
  return (
    <img
      className="social-icon-sponsor"
      src={sponsorIcon}
      width="24"
      height="24"
      alt=""
      aria-hidden="true"
    />
  )
}

export function SocialLinks({ onOpenDonation }: { onOpenDonation?: () => void }) {
  const { t } = useI18n()
  const [links, setLinks] = useState<{
    websiteUrl: string | null
    githubUrl: string | null
    instagramUrl: string | null
  }>({ websiteUrl: null, githubUrl: null, instagramUrl: null })
  useEffect(() => {
    void window.aiOffice
      .getSupportLinks()
      .then((next) =>
        setLinks({
          websiteUrl: next.websiteUrl,
          githubUrl: next.githubUrl,
          instagramUrl: next.instagramUrl,
        }),
      )
      .catch(() => {})
  }, [])
  return (
    <div className="about-links" aria-label="ARAOffice links">
      {links.websiteUrl && (
        <button
          type="button"
          title="Website"
          aria-label="Website"
          onClick={() => void window.aiOfficeWindow.openExternal(links.websiteUrl as string)}
        >
          <WebsiteIcon />
        </button>
      )}
      {links.githubUrl && (
        <button
          type="button"
          title="GitHub"
          aria-label="GitHub"
          onClick={() => void window.aiOfficeWindow.openExternal(links.githubUrl as string)}
        >
          <GitHubIcon />
        </button>
      )}
      {links.instagramUrl && (
        <button
          type="button"
          title="Instagram"
          aria-label="Instagram"
          onClick={() => void window.aiOfficeWindow.openExternal(links.instagramUrl as string)}
        >
          <InstagramIcon />
        </button>
      )}
      <button
        type="button"
        className="about-support-btn"
        title={t('supportTitle')}
        aria-label={t('supportTitle')}
        onClick={onOpenDonation}
      >
        <SponsorIcon />
      </button>
    </div>
  )
}

export function MeetAraPanel({
  onClose,
  className = '',
}: {
  onClose: () => void
  className?: string
}) {
  const { t } = useI18n()

  return (
    <div
      className={`topbar-popover ai-popover${className ? ` ${className}` : ''}`}
      role="dialog"
      aria-modal="false"
      aria-labelledby="meet-ara-title"
      aria-describedby="meet-ara-description"
    >
      <button className="topbar-popover-close" onClick={onClose} aria-label={t('araClose')}>
        ×
      </button>
      <div className="ai-popover-heading">
        <span className="ai-popover-eyebrow">
          <span className="ai-popover-spark" aria-hidden="true">
            ✦
          </span>
          {t('araEyebrow')}
        </span>
        <h2 id="meet-ara-title">{t('araTitle')}</h2>
        <p id="meet-ara-description">{t('araDescription')}</p>
      </div>
      <section className="ai-model-catalog" aria-labelledby="ai-model-info-title">
        <div className="ai-model-section-heading">
          <p className="ai-model-info-label" id="ai-model-info-title">
            {t('araFeatures')}
          </p>
        </div>
        <div className="ai-model-list">
          <article className="ai-model-card">
            <div className="ai-model-card-heading">
              <div>
                <strong>{t(ARA_COPY.label)}</strong>
                <span>{t(ARA_COPY.description)}</span>
              </div>
            </div>
            <p className="ai-model-card-best">
              <b>{t('araBestFor')}</b>
              {t(ARA_COPY.bestFor)}
            </p>
            <ul className="ai-model-capabilities">
              {ARA_COPY.capabilities.map((capability) => (
                <li key={capability.label}>
                  <CapabilityIcon name={capability.icon} />
                  <span>{t(capability.label)}</span>
                </li>
              ))}
            </ul>
          </article>
        </div>
      </section>
    </div>
  )
}
