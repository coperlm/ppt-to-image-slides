export type QosName = 'clear' | 'balanced' | 'small'

export interface Qos {
  name: QosName
  label: string
  hint: string
  targetLongEdge: number
  quality: number
}

/** pptx-preview 的布局基准宽度(px)，出图倍率由 targetLongEdge 反推 */
export const RENDER_WIDTH = 1280

export const QOS: Record<QosName, Qos> = {
  clear: {
    name: 'clear',
    label: '清晰',
    hint: '长边 3840px · 投影与放大查看',
    targetLongEdge: 3840,
    quality: 0.92,
  },
  balanced: {
    name: 'balanced',
    label: '均衡',
    hint: '长边 2560px · 清晰度与体积折中',
    targetLongEdge: 2560,
    quality: 0.85,
  },
  small: {
    name: 'small',
    label: '小体积',
    hint: '长边 1920px · 便于网络传输',
    targetLongEdge: 1920,
    quality: 0.78,
  },
}

export const DEFAULT_QOS: QosName = 'balanced'

export const QOS_ORDER: QosName[] = ['clear', 'balanced', 'small']

export const LIMITS = {
  maxInputBytes: 100 * 1024 * 1024,
  maxUncompressedBytes: 600 * 1024 * 1024,
  maxSlides: 200,
  thumbLongEdge: 320,
}
