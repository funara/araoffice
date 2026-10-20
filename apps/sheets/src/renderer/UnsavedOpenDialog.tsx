import { useI18n } from './i18n/locale'

import type { UnsavedChoice } from './open-guard'

/// Excel-style Save/Discard/Cancel prompt shown before an open replaces a
/// workbook with unsaved edits.
export function UnsavedOpenDialog({
  name,
  onChoice,
}: {
  readonly name: string
  readonly onChoice: (choice: UnsavedChoice) => void
}): React.JSX.Element {
  const { t } = useI18n()
  return (
    <div className="dialog-backdrop" onClick={() => onChoice('cancel')}>
      <div
        className="format-cells-dialog"
        role="dialog"
        aria-label={t('appUnsavedOpenTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('appUnsavedOpenTitle')}</header>
        <section className="dialog-body">
          <p className="dialog-note">{t('appUnsavedOpenBody', { name })}</p>
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={() => onChoice('cancel')}>
            {t('dlgCancel')}
          </button>
          <button className="secondary" onClick={() => onChoice('discard')}>
            {t('appUnsavedOpenDiscard')}
          </button>
          <button className="primary-action" onClick={() => onChoice('save')}>
            {t('aiSave')}
          </button>
        </div>
      </div>
    </div>
  )
}
