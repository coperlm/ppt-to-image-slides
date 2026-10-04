import type { MessageKey } from './i18n'

export type QosName = 'clear' | 'balanced' | 'small'

export interface Qos {
  name: QosName
  labelKey: MessageKey
  hintKey: MessageKey
  targetLongEdge: number
  quality: number
}

/** pptx-preview 的布局基准宽度(px)，出图倍率由 targetLongEdge 反推 */
export const RENDER_WIDTH = 1280

export const QOS: Record<QosName, Qos> = {
  clear: { name: 'clear', labelKey: 'qosClear', hintKey: 'qosClearHint', targetLongEdge: 3840, quality: 0.92 },
  balanced: { name: 'balanced', labelKey: 'qosBalanced', hintKey: 'qosBalancedHint', targetLongEdge: 2560, quality: 0.85 },
  small: { name: 'small', labelKey: 'qosSmall', hintKey: 'qosSmallHint', targetLongEdge: 1920, quality: 0.78 },
}

export const DEFAULT_QOS: QosName = 'balanced'

export const QOS_ORDER: QosName[] = ['clear', 'balanced', 'small']

export const QOS_STORAGE_KEY = 'ppt2img-qos'

export const THEME_STORAGE_KEY = 'ppt2img-theme'

export const LIMITS = {
  maxInputBytes: 100 * 1024 * 1024,
  maxUncompressedBytes: 600 * 1024 * 1024,
  maxSlides: 200,
  thumbLongEdge: 320,
}
