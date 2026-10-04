import JSZip from 'jszip'
import { LIMITS } from './config'
import { fmtBytes } from './format'
import { t } from './i18n'

export const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
export const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
export const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const FALLBACK_SLD_SZ = { cx: 12192000, cy: 6858000 }

export interface SlideSize {
  cx: number
  cy: number
  found: boolean
}

export interface Guarded {
  buffer: ArrayBuffer
  zip: JSZip
  slideCount: number
  slidePaths: string[]
  slideCountSource: 'sldIdLst' | 'fileScan'
  sldSz: SlideSize
  uncompressedBytes: number
}

export async function guardAndRead(file: File, limits = LIMITS): Promise<Guarded> {
  if (!file.name.toLowerCase().endsWith('.pptx')) {
    throw new Error(t('errUnsupported', { name: file.name }))
  }
  if (file.size > limits.maxInputBytes) {
    throw new Error(t('errTooBig', { size: fmtBytes(file.size), limit: fmtBytes(limits.maxInputBytes) }))
  }
  await assertZipContainer(file)

  const buffer = await file.arrayBuffer()
  const zip = await JSZip.loadAsync(buffer)

  const uncompressedBytes = estimateUncompressedBytes(zip)
  if (uncompressedBytes > limits.maxUncompressedBytes) {
    throw new Error(
      t('errZipBomb', { size: fmtBytes(uncompressedBytes), limit: fmtBytes(limits.maxUncompressedBytes) }),
    )
  }

  const presentationFile = zip.file('ppt/presentation.xml')
  if (!presentationFile) throw new Error(t('errNoPresentation'))
  const presentationXml = await presentationFile.async('string')

  const sldSz = readSlideSize(presentationXml)
  const listed = await resolveSlidePaths(zip, presentationXml)
  const slidePaths = listed.length ? listed : scanSlidePaths(zip)
  if (!slidePaths.length) throw new Error(t('errNoSlides'))
  if (slidePaths.length > limits.maxSlides) {
    throw new Error(t('errTooManySlides', { count: slidePaths.length, limit: limits.maxSlides }))
  }

  return {
    buffer,
    zip,
    slideCount: slidePaths.length,
    slidePaths,
    slideCountSource: listed.length ? 'sldIdLst' : 'fileScan',
    sldSz,
    uncompressedBytes,
  }
}

async function assertZipContainer(file: File): Promise<void> {
  const head = new Uint8Array(await file.slice(0, 4).arrayBuffer())
  if (head[0] === 0xd0 && head[1] === 0xcf && head[2] === 0x11 && head[3] === 0xe0) {
    throw new Error(t('errCfb'))
  }
  if (head[0] !== 0x50 || head[1] !== 0x4b) {
    throw new Error(t('errNotZip'))
  }
}

function estimateUncompressedBytes(zip: JSZip): number {
  let total = 0
  for (const entry of Object.values(zip.files)) {
    // JSZip 未在公开类型中暴露解压后大小；字段缺失时按 0 计，不阻断转换
    const size = (entry as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize
    if (typeof size === 'number') total += size
  }
  return total
}

export function parseXml(text: string): Document {
  const doc = new DOMParser().parseFromString(text, 'application/xml')
  const failure = doc.querySelector('parsererror')
  if (failure) throw new Error(t('errXml', { detail: (failure.textContent ?? '').slice(0, 160) }))
  return doc
}

function readSlideSize(presentationXml: string): SlideSize {
  const el = parseXml(presentationXml).getElementsByTagNameNS(NS_P, 'sldSz')[0]
  const cx = Number(el?.getAttribute('cx'))
  const cy = Number(el?.getAttribute('cy'))
  if (!el || !Number.isFinite(cx) || !Number.isFinite(cy) || cx <= 0 || cy <= 0) {
    return { ...FALLBACK_SLD_SZ, found: false }
  }
  return { cx, cy, found: true }
}

/** 以 presentation.xml 的 sldIdLst + rels 为准解析页面顺序；解析不出时返回空数组由调用方回退 */
async function resolveSlidePaths(zip: JSZip, presentationXml: string): Promise<string[]> {
  const relIds = Array.from(parseXml(presentationXml).getElementsByTagNameNS(NS_P, 'sldId'))
    .map((el) => el.getAttributeNS(NS_R, 'id'))
    .filter((id): id is string => Boolean(id))
  const relsFile = zip.file('ppt/_rels/presentation.xml.rels')
  if (!relIds.length || !relsFile) return []

  const targetById = new Map<string, string>()
  for (const rel of Array.from(parseXml(await relsFile.async('string')).getElementsByTagName('Relationship'))) {
    const id = rel.getAttribute('Id')
    const target = rel.getAttribute('Target')
    if (id && target) targetById.set(id, target)
  }

  const paths = relIds
    .map((id) => targetById.get(id))
    .filter((target): target is string => Boolean(target))
    .map((target) => normalizePartPath('ppt/', target))
  return paths.filter((path) => zip.file(path) !== null)
}

function scanSlidePaths(zip: JSZip): string[] {
  return Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => slideNumberOf(a) - slideNumberOf(b))
}

export function slideNumberOf(path: string): number {
  return Number(path.match(/(\d+)\.xml$/)?.[1] ?? 0)
}

/** 把 OPC 关系里的 Target 归一化成 zip 内的绝对路径 */
export function normalizePartPath(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const parts = baseDir.replace(/\/+$/, '').split('/').filter(Boolean)
  for (const segment of target.split('/')) {
    if (segment === '' || segment === '.') continue
    if (segment === '..') parts.pop()
    else parts.push(segment)
  }
  return parts.join('/')
}
