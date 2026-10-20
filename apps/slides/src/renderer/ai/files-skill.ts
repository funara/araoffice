import { createFilesSkill as createSharedFilesSkill } from '@araoffice/agent-core'
import type { AttachmentMeta } from '../../shared/ipc'
import { ATTACHMENT_IMAGE_EXTS } from '../../shared/ipc'
import { t } from '../i18n/locale'

/**
 * Chat-attachment capability as an AgentSkill (isomorphic to apps/docs files-skill):
 * each turn's context lists the attached local files; read_attachment reads their
 * extracted text in pages (parsing happens in the main process; files never leave
 * this machine).
 */
export function createFilesSkill(
  getAttachments: () => AttachmentMeta[],
  /** Called on each successful text read — lets the deck generator gate on unread attachments */
  onTextRead?: (path: string) => void,
) {
  return createSharedFilesSkill({
    getAttachments,
    readAttachment: (path, offset, maxChars) =>
      window.desktop.readAttachment(path, offset, maxChars),
    isImageExt: (ext) => ATTACHMENT_IMAGE_EXTS.has(ext),
    onTextRead,
    texts: {
      unknownTool: (name) => `Unknown tool: ${name}`,
      invalidIndex: () => 'Invalid attachment index (see the attachment list)',
      invalidIndexSummary: () => t('aiSumReadAttachment'),
      imageOutput: (name) =>
        `${name} is an image attachment already sent as an image with the user message; just view the image in the message, no text reading needed.`,
      imageSummary: (name) => t('aiSumImageAttachment', { name }),
      readFail: () => 'Read failed',
      readSummary: (name) => t('aiSumReadAttachmentName', { name }),
    },
  })
}
