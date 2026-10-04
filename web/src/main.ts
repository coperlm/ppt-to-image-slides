import './styles.css'
import {
  DEFAULT_QOS,
  LIMITS,
  QOS,
  QOS_ORDER,
  QOS_STORAGE_KEY,
  THEME_STORAGE_KEY,
  type QosName,
} from './config'
import { errorMessage, fmtBytes, fmtSeconds } from './format'
import { getLang, setLang, t, type Lang } from './i18n'
import { estimate } from './estimate'
import { guardAndRead, type Guarded } from './guards'
import { convert, type ConvertResult, type PageResult } from './pipeline'

function el<T extends HTMLElement>(id: string): T {
  const node = document.getElementById(id)
  if (!node) throw new Error(`page is missing #${id}`)
  return node as T
}

const dropzone = el<HTMLDivElement>('dropzone')
const fileInput = el<HTMLInputElement>('file')
const dzTitle = el<HTMLParagraphElement>('dz-title')
const fileMeta = el<HTMLParagraphElement>('filemeta')
const qosOptions = el<HTMLDivElement>('qos-options')
const convertBtn = el<HTMLButtonElement>('convert')
const cancelBtn = el<HTMLButtonElement>('cancel')
const downloadLink = el<HTMLAnchorElement>('download')
const statusEl = el<HTMLParagraphElement>('status')
const progressWrap = el<HTMLDivElement>('progress-wrap')
const progressBar = el<HTMLDivElement>('progress-bar')
const verifyEl = el<HTMLPreElement>('verify')
const thumbsEl = el<HTMLDivElement>('thumbs')
const logEl = el<HTMLPreElement>('log')
const estimateEl = el<HTMLParagraphElement>('estimate')
const langSelect = el<HTMLSelectElement>('lang')
const themeSelect = el<HTMLSelectElement>('theme')
const darkQuery = window.matchMedia('(prefers-color-scheme: dark)')

type Theme = 'light' | 'dark' | 'system'

let selectedFile: File | null = null
let controller: AbortController | null = null
let resultUrl: string | null = null
let lastResult: ConvertResult | null = null
let probed: Guarded | null = null

function log(message: string): void {
  logEl.textContent += `${message}\n`
  logEl.scrollTop = logEl.scrollHeight
}

type StatusKind = 'info' | 'error' | 'done'

function setStatus(text: string, kind: StatusKind = 'info'): void {
  statusEl.textContent = text
  statusEl.className = !text || kind === 'info' ? 'status' : `status ${kind}`
}

function setProgress(done: number, total: number): void {
  progressWrap.hidden = !total
  progressBar.style.width = total ? `${Math.round((done / total) * 100)}%` : '0%'
}

function addThumb(page: PageResult): void {
  const box = document.createElement('div')
  box.className = page.ok ? 'thumb' : 'thumb failed'
  const img = document.createElement('img')
  img.src = page.thumbUrl
  img.alt = `${page.index}`
  img.loading = 'lazy'
  const cap = document.createElement('div')
  cap.className = 'cap'
  cap.textContent = page.ok
    ? t('capPage', { index: page.index, w: page.width, h: page.height, ms: Math.round(page.ms) })
    : t('capPageFailed', { index: page.index })
  box.append(img, cap)
  thumbsEl.appendChild(box)
}

function selectedQos(): QosName {
  const checked = qosOptions.querySelector<HTMLInputElement>('input[name="qos"]:checked')
  return (checked?.value as QosName) ?? DEFAULT_QOS
}

function storedQos(): QosName {
  try {
    const stored = localStorage.getItem(QOS_STORAGE_KEY) as QosName | null
    return stored && QOS_ORDER.includes(stored) ? stored : DEFAULT_QOS
  } catch {
    return DEFAULT_QOS
  }
}

function buildQosOptions(): void {
  qosOptions.textContent = ''
  const preferred = storedQos()
  for (const name of QOS_ORDER) {
    const qos = QOS[name]
    const label = document.createElement('label')
    label.className = 'qos-option'

    const input = document.createElement('input')
    input.type = 'radio'
    input.name = 'qos'
    input.value = name
    input.checked = name === preferred

    const title = document.createElement('span')
    title.className = 'qos-name'
    title.textContent = t(qos.labelKey) + (name === DEFAULT_QOS ? t('qosRecommended') : '')

    const hint = document.createElement('span')
    hint.className = 'qos-hint'
    hint.textContent = t(qos.hintKey)

    label.append(input, title, hint)
    qosOptions.appendChild(label)
  }
  updateEstimate()
}

function updateButtons(): void {
  const running = controller !== null
  convertBtn.disabled = running || !selectedFile
  cancelBtn.hidden = !running
}

function updateEstimate(): void {
  if (!probed) {
    estimateEl.hidden = true
    return
  }
  const est = estimate(probed.slideCount, probed.sldSz.cx / probed.sldSz.cy, QOS[selectedQos()])
  estimateEl.hidden = false
  estimateEl.textContent =
    t('estLine', {
      pages: est.pages,
      bytes: `${fmtBytes(est.bytesLow)}–${fmtBytes(est.bytesHigh)}`,
      seconds: `${Math.max(1, Math.round(est.secondsLow))}–${Math.max(2, Math.round(est.secondsHigh))} s`,
    }) + (est.calibrated ? t('estCal') : t('estUncal'))
}

async function acceptFile(file: File | null | undefined): Promise<void> {
  if (!file || controller) return
  if (!file.name.toLowerCase().endsWith('.pptx')) {
    setStatus(t('statusUnsupported', { name: file.name }), 'error')
    return
  }
  if (file.size > LIMITS.maxInputBytes) {
    setStatus(t('statusTooBig', { size: fmtBytes(file.size), limit: fmtBytes(LIMITS.maxInputBytes) }), 'error')
    return
  }
  try {
    probed = await guardAndRead(file)
  } catch (error) {
    probed = null
    selectedFile = null
    setStatus(t('statusFailed', { message: errorMessage(error) }), 'error')
    updateButtons()
    updateEstimate()
    return
  }
  selectedFile = file
  dzTitle.textContent = file.name
  fileMeta.textContent = `${fmtBytes(file.size)} · ${t('dropHint', { limit: fmtBytes(LIMITS.maxInputBytes) })}`
  setStatus('')
  log(t('logSelected', { name: file.name, size: fmtBytes(file.size) }))
  updateButtons()
  updateEstimate()
  prefetchHeavyChunks()
}

function revokeResultUrl(): void {
  if (resultUrl) {
    URL.revokeObjectURL(resultUrl)
    resultUrl = null
  }
}

function resetPanels(): void {
  logEl.textContent = ''
  thumbsEl.textContent = ''
  verifyEl.textContent = t('verifyRunning')
  verifyEl.classList.remove('fail')
  downloadLink.hidden = true
  revokeResultUrl()
  setProgress(0, 0)
}

function showResult(result: ConvertResult): void {
  lastResult = result
  const { verify, stats } = result
  const lines = [
    `${t('verifyInputPages')}              ${result.slideCount}`,
    `${t('verifySlideXml')}     ${verify.slideCount}`,
    `${t('verifyWithBg')}    ${verify.withBackground}`,
    `${t('verifyMedia')}     ${verify.mediaCount}`,
    `${t('verifySize')}          ${fmtBytes(stats.outputBytes)}`,
    t('verifyTime', {
      total: fmtSeconds(stats.totalMs),
      render: fmtSeconds(stats.renderMs),
      raster: fmtSeconds(stats.rasterMs),
      pack: fmtSeconds(stats.packMs),
    }),
    '',
  ]
  if (stats.failedPages.length) {
    lines.push(t('verifyFailedPages', { pages: stats.failedPages.join('、') }), '')
  }
  if (verify.ok) {
    lines.push(t('verifyPass'))
    verifyEl.classList.remove('fail')
  } else {
    lines.push(t('verifyFail'), ...verify.problems.map((problem) => `  · ${problem}`))
    verifyEl.classList.add('fail')
  }
  verifyEl.textContent = lines.join('\n')

  setProgress(result.pages.length, result.pages.length)
  if (!verify.ok) {
    downloadLink.hidden = true
    setStatus(t('statusVerifyFail'), 'error')
    return
  }

  revokeResultUrl()
  resultUrl = URL.createObjectURL(result.blob)
  downloadLink.href = resultUrl
  downloadLink.download = result.fileName
  downloadLink.textContent = t('download', { name: result.fileName, size: fmtBytes(result.blob.size) })
  downloadLink.hidden = false
  setStatus(
    stats.failedPages.length
      ? t('statusDoneDegraded', { pages: stats.failedPages.join('、') })
      : t('statusDone'),
    stats.failedPages.length ? 'error' : 'done',
  )
}

function rerenderResult(): void {
  thumbsEl.textContent = ''
  if (!lastResult) return
  for (const page of lastResult.pages) addThumb(page)
  showResult(lastResult)
}

async function run(): Promise<void> {
  const file = selectedFile
  if (!file || controller) return

  const qos = QOS[selectedQos()]
  try {
    localStorage.setItem(QOS_STORAGE_KEY, qos.name)
  } catch {
    // 隐私模式下 localStorage 可能不可用，忽略
  }
  controller = new AbortController()
  updateButtons()
  resetPanels()
  log(t('logStart', { qos: t(qos.labelKey), edge: qos.targetLongEdge, quality: qos.quality }))

  try {
    const result = await convert(
      file,
      qos,
      {
        onStage: (stage) => setStatus(t('statusStage', { stage, name: file.name })),
        onLog: log,
        onProgress: (done, total, page) => {
          setProgress(done, total)
          addThumb(page)
          setStatus(t('statusRaster', { done, total, name: file.name }))
        },
      },
      controller.signal,
      probed ?? undefined,
    )
    showResult(result)
  } catch (error) {
    if (error instanceof DOMException && error.name === 'AbortError') {
      setStatus(t('statusCancelled'), 'done')
      log(t('logCancelled'))
    } else {
      const message = errorMessage(error)
      setStatus(t('statusFailed', { message }), 'error')
      verifyEl.textContent = t('verifyFailLine', { message })
      verifyEl.classList.add('fail')
      log(t('logFailed', { message }))
    }
  } finally {
    controller = null
    updateButtons()
  }
}

function storedTheme(): Theme {
  try {
    const stored = localStorage.getItem(THEME_STORAGE_KEY)
    return stored === 'light' || stored === 'dark' || stored === 'system' ? stored : 'system'
  } catch {
    return 'system'
  }
}

function effectiveTheme(): 'light' | 'dark' {
  const attr = document.documentElement.getAttribute('data-theme')
  if (attr === 'light' || attr === 'dark') return attr
  return darkQuery.matches ? 'dark' : 'light'
}

/** 让手机浏览器地址栏颜色跟着主题走 */
function syncThemeColor(): void {
  const meta = document.querySelector<HTMLMetaElement>('meta[name="theme-color"]')
  if (meta) meta.content = effectiveTheme() === 'dark' ? '#0b1220' : '#f6f7f9'
}

function applyTheme(theme: Theme): void {
  if (theme === 'system') document.documentElement.removeAttribute('data-theme')
  else document.documentElement.setAttribute('data-theme', theme)
  try {
    localStorage.setItem(THEME_STORAGE_KEY, theme)
  } catch {
    // 隐私模式下 localStorage 可能不可用，忽略
  }
  syncThemeColor()
}

/** 静态文案统一走 data-i18n；字符串全部来自本仓库 catalog，故 innerHTML 安全 */
function applyLang(): void {
  document.documentElement.lang = getLang() === 'zh' ? 'zh-CN' : 'en'
  document.title = t('appTitle')
  for (const node of Array.from(document.querySelectorAll<HTMLElement>('[data-i18n]'))) {
    node.innerHTML = t(node.dataset.i18n as never)
  }
  langSelect.value = getLang()
  themeSelect.value = storedTheme()
}

function switchLang(lang: Lang): void {
  setLang(lang)
  applyLang()
  buildQosOptions()
  dzTitle.textContent = selectedFile ? selectedFile.name : t('dropTitle')
  fileMeta.textContent = selectedFile
    ? `${fmtBytes(selectedFile.size)} · ${t('dropHint', { limit: fmtBytes(LIMITS.maxInputBytes) })}`
    : t('dropHint', { limit: fmtBytes(LIMITS.maxInputBytes) })
  if (lastResult) rerenderResult()
  else {
    setStatus('')
    verifyEl.textContent = t('verifyNotRun')
  }
}

dropzone.addEventListener('click', (event) => {
  if (event.target === fileInput) return
  fileInput.click()
})
dropzone.addEventListener('keydown', (event) => {
  if (event.key !== 'Enter' && event.key !== ' ') return
  event.preventDefault()
  fileInput.click()
})
fileInput.addEventListener('change', () => acceptFile(fileInput.files?.[0]))

// 整窗拖放：先阻止浏览器默认的「打开文件」行为
for (const type of ['dragover', 'drop'] as const) {
  window.addEventListener(type, (event) => event.preventDefault())
}
for (const type of ['dragenter', 'dragover'] as const) {
  dropzone.addEventListener(type, (event) => {
    event.preventDefault()
    dropzone.classList.add('over')
  })
}
for (const type of ['dragleave', 'drop'] as const) {
  dropzone.addEventListener(type, (event) => {
    event.preventDefault()
    dropzone.classList.remove('over')
  })
}
dropzone.addEventListener('drop', (event) => acceptFile(event.dataTransfer?.files?.[0]))
window.addEventListener('drop', (event) => acceptFile(event.dataTransfer?.files?.[0]))

qosOptions.addEventListener('change', updateEstimate)

convertBtn.addEventListener('click', () => void run())
cancelBtn.addEventListener('click', () => controller?.abort())
langSelect.addEventListener('change', () => switchLang(langSelect.value as Lang))
themeSelect.addEventListener('change', () => applyTheme(themeSelect.value as Theme))
darkQuery.addEventListener('change', syncThemeColor)

// 离线可用：只注册生产构建，避免 dev 的资源被缓存
if (import.meta.env.PROD && 'serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js').catch(() => {})
  })
}

/** 选完文件后空闲预取重 chunk，让「开始转换」少等一次网络；不碰首屏 */
function prefetchHeavyChunks(): void {
  const whenIdle = (callback: () => void) => {
    if (typeof window.requestIdleCallback === 'function') window.requestIdleCallback(callback, { timeout: 4000 })
    else window.setTimeout(callback, 800)
  }
  whenIdle(() => {
    void import('pptx-preview')
    void import('@zumer/snapdom')
    void import('pptxgenjs')
  })
}

// 出图依赖浏览器的渲染帧，标签页切到后台时会被挂起；切回来会自动继续
document.addEventListener('visibilitychange', () => {
  if (controller && document.visibilityState === 'hidden') {
    setStatus(t('statusPaused'))
  }
})

applyLang()
syncThemeColor()
buildQosOptions()
updateButtons()
verifyEl.textContent = t('verifyNotRun')
fileMeta.textContent = t('dropHint', { limit: fmtBytes(LIMITS.maxInputBytes) })
log(t('logReady', { limit: fmtBytes(LIMITS.maxInputBytes), slides: LIMITS.maxSlides }))
