import { createFilesSkill as createSharedFilesSkill } from '@araoffice/agent-core'
import type { AttachmentMeta } from '../../shared/desktop-api'
import { ATTACHMENT_IMAGE_EXTS } from '../../shared/desktop-api'
import { t } from '../i18n/locale'

/**
 * Chat-attachment capability as an AgentSkill (same structure as the files-skill
 * in apps/docs and apps/slides): each turn's context lists the attached local
 * files, and read_attachment reads their extracted text in pages (parsing
 * happens in the main process; files never leave the machine).
 */
export function createFilesSkill(getAttachments: () => readonly AttachmentMeta[]) {
  return createSharedFilesSkill({
    getAttachments,
    readAttachment: (path, offset, maxChars) =>
      window.desktopApi.readAttachment(path, offset, maxChars),
    isImageExt: (ext) => ATTACHMENT_IMAGE_EXTS.has(ext),
    texts: {
      unknownTool: (name) => `Unknown tool: ${name}`,
      invalidIndex: () => 'Invalid attachment index (see the attachment list)',
      invalidIndexSummary: () => t('aiToolReadAttachment'),
      imageOutput: (name) =>
        `${name} is an image attachment, already sent as an image with the user message — just look at the image in the message; no text to read.`,
      imageSummary: (name) => t('aiToolImageAttachment', { name }),
      readFail: (name) => t('aiToolReadFile', { name }),
      readSummary: (name) => t('aiToolReadAttachmentOf', { name }),
    },
  })
}
