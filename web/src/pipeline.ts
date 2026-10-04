import type { Qos } from './config'
import { errorMessage, fmtBytes, fmtSeconds, throwIfAborted, yieldToBrowser } from './format'
import { guardAndRead, type SlideSize } from './guards'
import { packBackgroundPptx } from './packer'
import { placeholderSlide, rasterizeSlide } from './rasterizer'
import { renderPptx } from './renderer'
import { verifyOutputPptx, type VerifyResult } from './verifier'

export interface PageResult {
  index: number
  ok: boolean
  ms: number
  width: number
  height: number
  thumbUrl: string
  error?: string
}

export interface Stats {
  totalMs: number
  guardMs: number
  renderMs: number
  rasterMs: number
  packMs: number
  verifyMs: number
  outputBytes: number
  okPages: number
  failedPages: number[]
}

export interface Hooks {
  onStage?(stage: string): void
  onLog?(message: string): void
  onProgress?(done: number, total: number, page: PageResult): void
}

export interface ConvertResult {
  blob: Blob
  fileName: string
  verify: VerifyResult
  stats: Stats
  pages: PageResult[]
  sldSz: SlideSize
  slideCount: number
  slideCountSource: 'sldIdLst' | 'fileScan'
  uncompressedBytes: number
  degraded: boolean
}

export async function convert(
  file: File,
  qos: Qos,
  hooks: Hooks = {},
  signal?: AbortSignal,
): Promise<ConvertResult> {
  const log = (message: string) => hooks.onLog?.(message)
  const stage = (name: string) => {
    hooks.onStage?.(name)
    log(name)
  }
  const started = performance.now()

  stage('校验输入…')
  const tGuard = performance.now()
  const guarded = await guardAndRead(file)
  const guardMs = performance.now() - tGuard
  log(
    `页面尺寸 ${guarded.sldSz.cx}×${guarded.sldSz.cy} EMU` +
      `${guarded.sldSz.found ? '' : '（未找到 sldSz，已按 16:9 兜底）'} · ` +
      `页数 ${guarded.slideCount}（来源 ${guarded.slideCountSource}） · ` +
      `解压后约 ${fmtBytes(guarded.uncompressedBytes)}`,
  )
  if (!guarded.sldSz.found) log('警告：缺少 <p:sldSz>，输出尺寸可能与原稿不一致')
  if (guarded.slideCountSource === 'fileScan') log('警告：无法解析 sldIdLst，页数已回退为按文件扫描统计')
  throwIfAborted(signal)

  await document.fonts.ready
  log('字体就绪（document.fonts.ready）')

  stage('渲染幻灯片…')
  const tRender = performance.now()
  const rendered = await renderPptx(guarded.buffer, guarded.sldSz)
  const renderMs = performance.now() - tRender
  log(`渲染出 ${rendered.slideEls.length} 个页面 DOM，用时 ${fmtSeconds(renderMs)}`)

  if (rendered.slideEls.length !== guarded.slideCount) {
    rendered.destroy()
    throw new Error(
      `渲染页数 ${rendered.slideEls.length} 与文件页数 ${guarded.slideCount} 不一致，已中止以免输出缺页`,
    )
  }
  throwIfAborted(signal)

  stage('逐页出图…')
  const tRaster = performance.now()
  const images: string[] = []
  const pages: PageResult[] = []
  for (const [offset, el] of rendered.slideEls.entries()) {
    throwIfAborted(signal)
    const index = offset + 1
    const tPage = performance.now()
    let raster
    let pageError: string | undefined
    try {
      raster = await rasterizeSlide(el, qos)
    } catch (error) {
      pageError = errorMessage(error)
      raster = placeholderSlide(guarded.sldSz.cx, guarded.sldSz.cy, `第 ${index} 页渲染失败`)
      log(`第 ${index} 页出图失败：${pageError} → 已用占位页替代`)
    }
    images.push(raster.dataUrl)
    const page: PageResult = {
      index,
      ok: !pageError,
      ms: performance.now() - tPage,
      width: raster.width,
      height: raster.height,
      thumbUrl: raster.thumbUrl,
      error: pageError,
    }
    pages.push(page)
    hooks.onProgress?.(pages.length, rendered.slideEls.length, page)
    await yieldToBrowser()
  }
  const rasterMs = performance.now() - tRaster
  rendered.destroy()

  const failedPages = pages.filter((page) => !page.ok).map((page) => page.index)
  if (failedPages.length) log(`警告：${failedPages.length} 页出图失败（第 ${failedPages.join('、')} 页）`)

  stage('重建 PPTX…')
  const tPack = performance.now()
  const blob = await packBackgroundPptx(images, guarded.sldSz)
  const packMs = performance.now() - tPack
  images.length = 0
  log(`输出 ${fmtBytes(blob.size)}，打包用时 ${fmtSeconds(packMs)}`)
  throwIfAborted(signal)

  stage('结构自检…')
  const tVerify = performance.now()
  const verify = await verifyOutputPptx(blob, guarded.slideCount)
  const verifyMs = performance.now() - tVerify

  return {
    blob,
    fileName: `${file.name.replace(/\.pptx$/i, '')}_image.pptx`,
    verify,
    stats: {
      totalMs: performance.now() - started,
      guardMs,
      renderMs,
      rasterMs,
      packMs,
      verifyMs,
      outputBytes: blob.size,
      okPages: pages.length - failedPages.length,
      failedPages,
    },
    pages,
    sldSz: guarded.sldSz,
    slideCount: guarded.slideCount,
    slideCountSource: guarded.slideCountSource,
    uncompressedBytes: guarded.uncompressedBytes,
    degraded: failedPages.length > 0,
  }
}
