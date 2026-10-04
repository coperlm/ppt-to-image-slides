import type { Qos } from './config'
import { errorMessage, fmtBytes, fmtSeconds, throwIfAborted, yieldToBrowser } from './format'
import { guardAndRead, type SlideSize } from './guards'
import { t } from './i18n'
import { extractNotesTexts, injectNotes } from './notes'
import { packBackgroundPptx } from './packer'
import { placeholderSlide, rasterizeSlide } from './rasterizer'
import { createRenderer } from './renderer'
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

  stage(t('stageGuard'))
  const tGuard = performance.now()
  const guarded = await guardAndRead(file)
  const guardMs = performance.now() - tGuard
  log(
    t('logMeta', {
      cx: guarded.sldSz.cx,
      cy: guarded.sldSz.cy,
      count: guarded.slideCount,
      source: guarded.slideCountSource,
      uncompressed: fmtBytes(guarded.uncompressedBytes),
    }) + (guarded.sldSz.found ? '' : t('logMetaSldSzMissing')),
  )
  if (!guarded.sldSz.found) log(t('logWarnSldSz'))
  if (guarded.slideCountSource === 'fileScan') log(t('logWarnFileScan'))
  const notesTexts = await extractNotesTexts(guarded.zip, guarded.slidePaths)
  const notedPages = notesTexts.filter((text) => text !== null).length
  if (notedPages) log(t('logNotes', { count: notedPages }))
  throwIfAborted(signal)

  await document.fonts.ready
  log(t('logFonts'))

  stage(t('stageRender'))
  const tRender = performance.now()
  const renderer = await createRenderer(guarded.buffer, guarded.sldSz)
  const renderMs = performance.now() - tRender
  log(t('logRenderer', { ms: fmtSeconds(renderMs) }))

  if (renderer.slideCount !== guarded.slideCount) {
    renderer.destroy()
    throw new Error(t('errRenderCount', { rendered: renderer.slideCount, expected: guarded.slideCount }))
  }
  throwIfAborted(signal)

  stage(t('stageRaster'))
  const tRaster = performance.now()
  const images: string[] = []
  const pages: PageResult[] = []
  for (let offset = 0; offset < renderer.slideCount; offset++) {
    throwIfAborted(signal)
    const index = offset + 1
    const tPage = performance.now()
    let raster
    let pageError: string | undefined
    try {
      raster = await rasterizeSlide(await renderer.renderSlide(offset), qos)
    } catch (error) {
      pageError = errorMessage(error)
      raster = placeholderSlide(guarded.sldSz.cx, guarded.sldSz.cy, t('placeholder', { index }))
      log(t('logPageFail', { index, message: pageError }))
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
    hooks.onProgress?.(pages.length, renderer.slideCount, page)
    await yieldToBrowser()
  }
  const rasterMs = performance.now() - tRaster
  renderer.destroy()

  const failedPages = pages.filter((page) => !page.ok).map((page) => page.index)
  if (failedPages.length) log(t('logWarnFailed', { count: failedPages.length, pages: failedPages.join('、') }))

  stage(t('stagePack'))
  const tPack = performance.now()
  let output = await packBackgroundPptx(images, guarded.sldSz)
  images.length = 0
  if (notedPages) {
    output = await injectNotes(output, notesTexts)
    log(t('logNotesInjected', { count: notedPages }))
  }
  const packMs = performance.now() - tPack
  log(t('logOutput', { size: fmtBytes(output.size), ms: fmtSeconds(packMs) }))
  throwIfAborted(signal)

  stage(t('stageVerify'))
  const tVerify = performance.now()
  const verify = await verifyOutputPptx(output, guarded.slideCount)
  const verifyMs = performance.now() - tVerify

  return {
    blob: output,
    fileName: `${file.name.replace(/\.pptx$/i, '')}_image.pptx`,
    verify,
    stats: {
      totalMs: performance.now() - started,
      guardMs,
      renderMs,
      rasterMs,
      packMs,
      verifyMs,
      outputBytes: output.size,
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
