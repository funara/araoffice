import { useEffect, useState } from 'react'
import type { UiTheme } from '../../shared/home-api'
import { CommandCenterModal } from './CommandCenterModal'
import { DonationModal } from './DonationModal'
import { FeedbackModal } from './FeedbackModal'
import { Home } from './Home'
import { Onboarding } from './Onboarding'
import { ShortcutsModal } from './ShortcutsModal'
import { TabBar } from './TabBar'

interface AppFrameProps {
  /** resolved before first paint (main.tsx) so home never flashes under the overlay */
  initialOnboardingSeen: boolean
  initialTheme: UiTheme
}

export function AppFrame({ initialTheme, initialOnboardingSeen }: AppFrameProps) {
  const [homeActive, setHomeActive] = useState(true)
  const [theme, setTheme] = useState<UiTheme>(initialTheme)
  const [systemDark, setSystemDark] = useState(
    typeof window === 'undefined'
      ? false
      : window.matchMedia('(prefers-color-scheme: dark)').matches,
  )
  const [commandCenterOpen, setCommandCenterOpen] = useState(false)
  const [feedbackOpen, setFeedbackOpen] = useState(false)
  const [shortcutsOpen, setShortcutsOpen] = useState(false)
  const [donationOpen, setDonationOpen] = useState(false)

  useEffect(() => {
    void window.aiOffice.getTheme().then(setTheme)
    const media = window.matchMedia('(prefers-color-scheme: dark)')
    const syncSystemTheme = () => setSystemDark(media.matches)
    syncSystemTheme()
    media.addEventListener('change', syncSystemTheme)
    return () => media.removeEventListener('change', syncSystemTheme)
  }, [])

  const isDarkTheme = theme === 'dark' || (theme === 'system' && systemDark)

  const changeTheme = (next: UiTheme) => {
    setTheme(next)
    if (next === 'system') {
      const isDark = window.matchMedia('(prefers-color-scheme: dark)').matches
      document.documentElement.setAttribute('data-theme', isDark ? 'dark' : 'light')
    } else {
      document.documentElement.setAttribute('data-theme', next)
    }
    void window.aiOffice.setTheme(next).catch(() => {})
  }

  const toggleTheme = () => changeTheme(isDarkTheme ? 'light' : 'dark')

  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault()
        setCommandCenterOpen((prev) => !prev)
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [])

  const [onboardingMode, setOnboardingMode] = useState<'first-run' | 'features' | null>(
    initialOnboardingSeen ? null : 'first-run',
  )
  useEffect(() => {
    const applyTabs = (tabs: Awaited<ReturnType<typeof window.aiOfficeTabs.list>>) => {
      const active = tabs.find((tab) => tab.active)
      setHomeActive(!active || active.kind === 'home')
    }
    void window.aiOfficeTabs.list().then(applyTabs)
    return window.aiOfficeTabs.onChanged(applyTabs)
  }, [])

  const finishOnboarding = () => {
    setOnboardingMode(null)
    void window.aiOffice.setOnboardingSeen().catch(() => {})
  }
  // The app is fully usable without AI: finishing the tour never opens a
  // sign-in or a consent dialog.
  const startWithAra = () => finishOnboarding()
  // ponytail: features carousel is a re-openable tour, not consent — closing it
  // must never overwrite any persisted state.
  const closeFeatures = () => setOnboardingMode(null)
  const openFeatureCarousel = () => {
    void window.aiOfficeTabs.activate('home')
    setOnboardingMode('features')
  }
  // An editor's "Configure AI Provider" card asks the shell to open the Home
  // AI dialog. The counter makes each request a distinct trigger.
  const [aiSettingsRequest, setAiSettingsRequest] = useState(0)
  useEffect(() => window.aiOffice.onConfigureAi(() => setAiSettingsRequest((n) => n + 1)), [])

  const isOverlayOpen =
    commandCenterOpen ||
    feedbackOpen ||
    shortcutsOpen ||
    donationOpen ||
    (onboardingMode !== null && homeActive)

  const [editorSnapshot, setEditorSnapshot] = useState<string | null>(null)

  useEffect(() => {
    let canceled = false
    if (isOverlayOpen) {
      void window.aiOfficeTabs
        ?.setOverlayVisible?.(true)
        .then((snapshot) => {
          if (!canceled && snapshot) {
            setEditorSnapshot(snapshot)
          }
        })
        .catch(() => {})
    } else {
      setEditorSnapshot(null)
      void window.aiOfficeTabs?.setOverlayVisible?.(false).catch(() => {})
    }
    return () => {
      canceled = true
    }
  }, [isOverlayOpen])

  return (
    <div className="app-frame">
      <TabBar
        isDarkTheme={isDarkTheme}
        onToggleTheme={toggleTheme}
        onOpenCommandCenter={() => setCommandCenterOpen(true)}
      />
      {/* docs/sheets tabs render as WebContentsView children of this window, positioned
       * by the main process to cover this area — only Home paints its own content here. */}
      <div className="app-frame-content" style={{ visibility: homeActive ? 'visible' : 'hidden' }}>
        <Home
          onOpenFeatures={openFeatureCarousel}
          onOpenCommandCenter={() => setCommandCenterOpen(true)}
          openAiSettings={aiSettingsRequest}
        />
      </div>
      {/* If an editor tab is active while a shell overlay is open, render the captured
       * snapshot behind the overlay so the backdrop blur blurs the document content seamlessly */}
      {!homeActive && editorSnapshot && (
        <div
          className="app-frame-snapshot"
          style={{
            backgroundImage: `url(${editorSnapshot})`,
          }}
        />
      )}
      {/* editor WebContentsViews paint above ALL shell DOM, so the overlay only
       * renders while the home tab is active — it comes back when home does */}
      {onboardingMode !== null && homeActive && (
        <Onboarding
          isDarkTheme={isDarkTheme}
          onToggleTheme={toggleTheme}
          initialTheme={initialTheme}
          onDone={onboardingMode === 'features' ? closeFeatures : () => finishOnboarding()}
          onStart={onboardingMode === 'features' ? closeFeatures : startWithAra}
        />
      )}
      {/* Always mounted so the exit fade plays before unmount */}
      <CommandCenterModal
        isOpen={commandCenterOpen}
        isDarkTheme={isDarkTheme}
        onClose={() => setCommandCenterOpen(false)}
        onToggleTheme={toggleTheme}
        onOpenFeedback={() => {
          setCommandCenterOpen(false)
          setFeedbackOpen(true)
        }}
        onOpenFeatures={openFeatureCarousel}
        onOpenShortcuts={() => {
          setCommandCenterOpen(false)
          setShortcutsOpen(true)
        }}
        onOpenDonation={() => {
          setCommandCenterOpen(false)
          setDonationOpen(true)
        }}
      />
      {feedbackOpen && <FeedbackModal onClose={() => setFeedbackOpen(false)} />}
      {shortcutsOpen && <ShortcutsModal onClose={() => setShortcutsOpen(false)} />}
      {donationOpen && <DonationModal onClose={() => setDonationOpen(false)} />}
    </div>
  )
}
