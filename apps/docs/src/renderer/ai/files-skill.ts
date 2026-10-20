import { createFilesSkill as createSharedFilesSkill } from '@araoffice/agent-core'
import type { AttachmentMeta } from '../../shared/ipc'
import { ATTACHMENT_IMAGE_EXTS } from '../../shared/ipc'
import { t } from '../i18n/locale'

/**
 * Chat-attachment capability as an AgentSkill: the per-turn context lists the
 * attached local files, and read_attachment pages through their extracted
 * text (parsing happens in the main process; files never leave the machine).
 */
export function createFilesSkill(getAttachments: () => AttachmentMeta[]) {
  return createSharedFilesSkill({
    getAttachments,
    readAttachment: (path, offset, maxChars) =>
      window.desktop.readAttachment(path, offset, maxChars),
    isImageExt: (ext) => ATTACHMENT_IMAGE_EXTS.has(ext),
    texts: {
      unknownTool: (name) => `unknown tool: ${name}`,
      invalidIndex: () => 'invalid attachment index (see the attachment list)',
      invalidIndexSummary: () => t('aiSumReadAttachment'),
      imageOutput: (name) =>
        `${name} is an image attachment already sent as an image with the user message; just look at the image in the message, no text to read.`,
      imageSummary: (name) => t('aiSumImageAttachment', { name }),
      readFail: (name) => t('aiSumRead', { name }),
      readSummary: (name) => t('aiSumReadAttachmentName', { name }),
    },
  })
}
