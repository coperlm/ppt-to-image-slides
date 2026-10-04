import './styles.css'
import { DEFAULT_QOS, LIMITS, QOS, QOS_ORDER, type QosName } from './config'
import { errorMessage, fmtBytes, fmtSeconds } from './format'
import { convert, type ConvertResult, type PageResult } from './pipeline'

function el<T extends HTMLElement>(id: string): T {
  const node = document.getElementById(id)
  if (!node) throw new Error(`页面缺少元素 #${id}`)
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

let selectedFile: File | null = null
let controller: AbortController | null = null
let resultUrl: string | null = null

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
  img.alt = `第 ${page.index} 页`
  img.loading = 'lazy'
  const cap = document.createElement('div')
  cap.className = 'cap'
  cap.textContent = page.ok
    ? `第 ${page.index} 页 · ${page.width}×${page.height} · ${Math.round(page.ms)} ms`
    : `第 ${page.index} 页 · 渲染失败（已用占位页）`
  box.append(img, cap)
  thumbsEl.appendChild(box)
}

function selectedQos(): QosName {
  const checked = qosOptions.querySelector<HTMLInputElement>('input[name="qos"]:checked')
  return (checked?.value as QosName) ?? DEFAULT_QOS
}

function buildQosOptions(): void {
  for (const name of QOS_ORDER) {
    const qos = QOS[name]
    const label = document.createElement('label')
    label.className = 'qos-option'

    const input = document.createElement('input')
    input.type = 'radio'
    input.name = 'qos'
    input.value = name
    input.checked = name === DEFAULT_QOS

    const title = document.createElement('span')
    title.className = 'qos-name'
    title.textContent = name === DEFAULT_QOS ? `${qos.label}（推荐）` : qos.label

    const hint = document.createElement('span')
    hint.className = 'qos-hint'
    hint.textContent = qos.hint

    label.append(input, title, hint)
    qosOptions.appendChild(label)
  }
}

function updateButtons(): void {
  const running = controller !== null
  convertBtn.disabled = running || !selectedFile
  cancelBtn.hidden = !running
}

function acceptFile(file: File | null | undefined): void {
  if (!file || controller) return
  if (!file.name.toLowerCase().endsWith('.pptx')) {
    setStatus(`不支持的文件类型：${file.name}（仅支持 .pptx）`, 'error')
    return
  }
  if (file.size > LIMITS.maxInputBytes) {
    setStatus(`文件 ${fmtBytes(file.size)} 超过上限 ${fmtBytes(LIMITS.maxInputBytes)}`, 'error')
    return
  }
  selectedFile = file
  dzTitle.textContent = file.name
  fileMeta.textContent = `${fmtBytes(file.size)} · 上限 ${fmtBytes(LIMITS.maxInputBytes)}`
  setStatus('')
  log(`已选择：${file.name}（${fmtBytes(file.size)}）`)
  updateButtons()
}

function resetPanels(): void {
  logEl.textContent = ''
  thumbsEl.textContent = ''
  verifyEl.textContent = '（运行中…）'
  verifyEl.classList.remove('fail')
  downloadLink.hidden = true
  revokeResultUrl()
  setProgress(0, 0)
}

function revokeResultUrl(): void {
  if (resultUrl) {
    URL.revokeObjectURL(resultUrl)
    resultUrl = null
  }
}

function showResult(result: ConvertResult): void {
  const { verify, stats } = result
  const lines = [
    `输入页数:              ${result.slideCount}`,
    `输出 slide XML 数:     ${verify.slideCount}`,
    `背景图解析成功的页:    ${verify.withBackground}`,
    `包内 media 图片数:     ${verify.mediaCount}`,
    `输出文件大小:          ${fmtBytes(stats.outputBytes)}`,
    `总耗时:                ${fmtSeconds(stats.totalMs)}（渲染 ${fmtSeconds(stats.renderMs)} / 出图 ${fmtSeconds(
      stats.rasterMs,
    )} / 打包 ${fmtSeconds(stats.packMs)}）`,
    '',
  ]
  if (stats.failedPages.length) {
    lines.push(`出图失败并替换为占位页: 第 ${stats.failedPages.join('、')} 页`, '')
  }
  if (verify.ok) {
    lines.push('判定: ✅ PASS —— 每页背景的 r:embed 均解析到包内图片，且字节与扩展名一致')
    verifyEl.classList.remove('fail')
  } else {
    lines.push('判定: ❌ FAIL', ...verify.problems.map((problem) => `  · ${problem}`))
    verifyEl.classList.add('fail')
  }
  verifyEl.textContent = lines.join('\n')

  setProgress(result.pages.length, result.pages.length)
  if (!verify.ok) {
    downloadLink.hidden = true
    setStatus('结构自检未通过，已阻止下载（问题见上）', 'error')
    return
  }

  revokeResultUrl()
  resultUrl = URL.createObjectURL(result.blob)
  downloadLink.href = resultUrl
  downloadLink.download = result.fileName
  downloadLink.textContent = `下载 ${result.fileName}（${fmtBytes(result.blob.size)}）`
  downloadLink.hidden = false
  setStatus(
    stats.failedPages.length
      ? `完成，但第 ${stats.failedPages.join('、')} 页为占位页，请核对预览后再下载`
      : '完成，请核对预览后下载',
    stats.failedPages.length ? 'error' : 'done',
  )
}

async function run(): Promise<void> {
  const file = selectedFile
  if (!file || controller) return

  const qos = QOS[selectedQos()]
  controller = new AbortController()
  updateButtons()
  resetPanels()
  log(`开始转换 · 档位「${qos.label}」长边 ${qos.targetLongEdge}px · JPEG 质量 ${qos.quality}`)

  try {
    const result = await convert(
      file,
      qos,
      {
        onStage: (stage) => setStatus(`${stage}（${file.name}）`),
        onLog: log,
        onProgress: (done, total, page) => {
          setProgress(done, total)
          addThumb(page)
          setStatus(`逐页出图 ${done}/${total}（${file.name}）`)
        },
      },
      controller.signal,
    )
    showResult(result)
  } catch (error) {
    if (error instanceof DOMException && error.name === 'AbortError') {
      setStatus('已取消', 'done')
      log('已取消')
    } else {
      const message = errorMessage(error)
      setStatus(`转换失败：${message}`, 'error')
      verifyEl.textContent = `失败：${message}`
      verifyEl.classList.add('fail')
      log(`❌ ${message}`)
    }
  } finally {
    controller = null
    updateButtons()
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

convertBtn.addEventListener('click', () => void run())
cancelBtn.addEventListener('click', () => controller?.abort())

// 出图依赖浏览器的渲染帧，标签页切到后台时会被挂起；切回来会自动继续
document.addEventListener('visibilitychange', () => {
  if (controller && document.visibilityState === 'hidden') {
    setStatus('已暂停：浏览器会挂起后台标签页，切回本页即自动继续')
  }
})

buildQosOptions()
updateButtons()
log(`就绪。选择 .pptx 后点击「开始转换」（上限 ${fmtBytes(LIMITS.maxInputBytes)} / ${LIMITS.maxSlides} 页）。`)
