import { LIMITS, type Qos } from './config'
import { t } from './i18n'

export interface Raster {
  dataUrl: string
  thumbUrl: string
  width: number
  height: number
}

export async function rasterizeSlide(el: HTMLElement, qos: Qos): Promise<Raster> {
  const { snapdom } = await import('@zumer/snapdom')
  const longEdge = Math.max(el.offsetWidth, el.offsetHeight)
  if (!longEdge) throw new Error(t('errZeroSize'))

  // snapdom 的出图基准是 display box × scale × dpr；固定 dpr=1，否则输出分辨率会随用户设备的
  // devicePixelRatio（HiDPI / 浏览器缩放）漂移，同一档位在不同机器上体积和清晰度都不一样
  const canvas = await snapdom.toCanvas(el, { scale: qos.targetLongEdge / longEdge, dpr: 1 })
  try {
    return {
      dataUrl: canvas.toDataURL('image/jpeg', qos.quality),
      thumbUrl: toThumb(canvas),
      width: canvas.width,
      height: canvas.height,
    }
  } finally {
    canvas.width = 0
    canvas.height = 0
  }
}

/** 某页出图失败时的替代页：保持页数与背景判据成立，同时让用户一眼看到是哪页出了问题 */
export function placeholderSlide(cx: number, cy: number, label: string): Raster {
  const width = 1280
  const height = Math.max(1, Math.round((width * cy) / cx))
  const canvas = document.createElement('canvas')
  canvas.width = width
  canvas.height = height
  const ctx = canvas.getContext('2d')
  if (!ctx) throw new Error(t('errNoCanvas'))

  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, width, height)
  ctx.strokeStyle = '#e11d48'
  ctx.lineWidth = 6
  ctx.strokeRect(3, 3, width - 6, height - 6)
  ctx.fillStyle = '#e11d48'
  ctx.font = `600 ${Math.round(width / 24)}px system-ui, "Noto Sans CJK SC", sans-serif`
  ctx.textAlign = 'center'
  ctx.textBaseline = 'middle'
  ctx.fillText(label, width / 2, height / 2)

  const dataUrl = canvas.toDataURL('image/jpeg', 0.85)
  canvas.width = 0
  canvas.height = 0
  return { dataUrl, thumbUrl: dataUrl, width, height }
}

function toThumb(source: HTMLCanvasElement): string {
  const ratio = Math.min(1, LIMITS.thumbLongEdge / Math.max(source.width, source.height))
  const canvas = document.createElement('canvas')
  canvas.width = Math.max(1, Math.round(source.width * ratio))
  canvas.height = Math.max(1, Math.round(source.height * ratio))
  const ctx = canvas.getContext('2d')
  if (!ctx) return ''
  ctx.drawImage(source, 0, 0, canvas.width, canvas.height)
  const url = canvas.toDataURL('image/jpeg', 0.72)
  canvas.width = 0
  canvas.height = 0
  return url
}
