import type { Qos, QosName } from './config'

const CAL_KEY = 'ppt2img-cal-v1'
const HISTORY = 5

interface CalEntry {
  bpx: number[]
  msPage: number[]
}
type Cal = Partial<Record<QosName, CalEntry>>

// 首次（无本机历史）用的保守常数，来自 2026-10-04 对 11 页样本的实测外推
const FALLBACK_BPX: Record<QosName, number> = { clear: 0.1, balanced: 0.08, small: 0.06 }
const FALLBACK_MS_PAGE: Record<QosName, number> = { clear: 450, balanced: 260, small: 180 }
const FIXED_MS = 1800

export interface Estimate {
  pages: number
  bytesLow: number
  bytesHigh: number
  secondsLow: number
  secondsHigh: number
  calibrated: boolean
}

function readCal(): Cal {
  try {
    return JSON.parse(localStorage.getItem(CAL_KEY) ?? '{}') as Cal
  } catch {
    return {}
  }
}

function median(values: number[]): number | null {
  if (!values.length) return null
  const sorted = [...values].sort((a, b) => a - b)
  return sorted[Math.floor(sorted.length / 2)]
}

/** 转换完成后用真实数据自校准，之后预估会越来越准 */
export function recordActual(qos: QosName, pixels: number, bytes: number, rasterMs: number, pages: number): void {
  if (!pixels || !pages) return
  try {
    const cal = readCal()
    const entry = (cal[qos] ??= { bpx: [], msPage: [] })
    entry.bpx = [...entry.bpx, bytes / pixels].slice(-HISTORY)
    entry.msPage = [...entry.msPage, rasterMs / pages].slice(-HISTORY)
    localStorage.setItem(CAL_KEY, JSON.stringify(cal))
  } catch {
    // 隐私模式下 localStorage 可能不可用，忽略
  }
}

export function estimate(pages: number, aspect: number, qos: Qos): Estimate {
  const longEdge = qos.targetLongEdge
  const pixelsPerPage = (longEdge * longEdge) / Math.max(aspect, 1 / aspect)
  const entry = readCal()[qos.name]
  const bpx = median(entry?.bpx ?? []) ?? FALLBACK_BPX[qos.name]
  const msPage = median(entry?.msPage ?? []) ?? FALLBACK_MS_PAGE[qos.name]
  const calibrated = Boolean(entry?.bpx.length)

  const bytes = pages * pixelsPerPage * bpx
  const [lo, hi] = calibrated ? [0.85, 1.2] : [0.7, 1.4]
  const seconds = (FIXED_MS + pages * msPage) / 1000

  return {
    pages,
    bytesLow: bytes * lo,
    bytesHigh: bytes * hi,
    secondsLow: seconds * 0.6,
    secondsHigh: seconds * 1.7,
    calibrated,
  }
}
