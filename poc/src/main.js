import JSZip from 'jszip'
import { init as pptxPreviewInit } from 'pptx-preview'
import { snapdom } from '@zumer/snapdom'
import PptxGenJS from 'pptxgenjs'

// ---- 配置 ----
const EMU_PER_INCH = 914400
const MAX_INPUT_BYTES = 100 * 1024 * 1024        // 输入上限 100MB（可调）
const MAX_UNCOMPRESSED_BYTES = 600 * 1024 * 1024 // zip 炸弹防御：解压后总量上限
const RENDER_WIDTH = 1280                        // 渲染基准宽度(px)，再乘 scale 出图
const DEFAULT_SLD = { cx: 12192000, cy: 6858000 } // 找不到 sldSz 时的兜底(16:9)

// ---- DOM ----
const $ = (id) => document.getElementById(id)
const fileInput = $('file')
const convertBtn = $('convert')
const downloadLink = $('download')
const logEl = $('log')
const verifyEl = $('verify')
const leftCol = $('left')
const rightCol = $('right')
const stage = $('stage')

let selectedFile = null
let sourceArrayBuffer = null
let lastOutputBlob = null

// ---- 小工具 ----
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))
const tick = () => new Promise((r) => requestAnimationFrame(() => r()))
function log(msg) {
  logEl.textContent += `${msg}\n`
  logEl.scrollTop = logEl.scrollHeight
  console.log('[poc]', msg)
}
function fmtBytes(n) {
  if (n < 1024) return n + ' B'
  if (n < 1024 * 1024) return (n / 1024).toFixed(1) + ' KB'
  return (n / 1024 / 1024).toFixed(2) + ' MB'
}

// ---- 1. 读取页面尺寸 & zip 炸弹防御 ----
async function readSlideSize(zip) {
  const f = zip.file('ppt/presentation.xml')
  if (!f) return { ...DEFAULT_SLD, found: false }
  const xml = await f.async('string')
  const m = xml.match(/<p:sldSz[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/)
  if (!m) return { ...DEFAULT_SLD, found: false }
  return { cx: parseInt(m[1], 10), cy: parseInt(m[2], 10), found: true }
}

async function guardZipBomb(zip) {
  let total = 0
  for (const name of Object.keys(zip.files)) {
    const d = zip.files[name]._data
    if (d && typeof d.uncompressedSize === 'number') total += d.uncompressedSize
  }
  log(`解压后总量(估算): ${fmtBytes(total)}`)
  if (total > MAX_UNCOMPRESSED_BYTES) {
    throw new Error(`解压后体积 ${fmtBytes(total)} 超过上限 ${fmtBytes(MAX_UNCOMPRESSED_BYTES)}，已拒绝（防 zip 炸弹）`)
  }
  return total
}

async function countSlideFiles(zip) {
  return Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length
}

// ---- 2. 渲染 + 出图 ----
async function waitImages(root, timeout = 8000) {
  const imgs = Array.from(root.querySelectorAll('img'))
  await Promise.all(
    imgs.map(
      (img) =>
        new Promise((res) => {
          if (img.complete && img.naturalWidth) return res()
          const done = () => res()
          img.addEventListener('load', done, { once: true })
          img.addEventListener('error', done, { once: true })
          setTimeout(done, timeout)
        })
    )
  )
}

async function renderAndRasterize(arrayBuffer, sld, scale, quality) {
  const height = Math.round((RENDER_WIDTH * sld.cy) / sld.cx)
  const container = document.createElement('div')
  stage.appendChild(container)

  const previewer = pptxPreviewInit(container, { width: RENDER_WIDTH, height, mode: 'list' })
  log(`开始渲染 (基准 ${RENDER_WIDTH}x${height})…`)
  await previewer.preview(arrayBuffer)
  // 等待幻灯片元素与内部图片就绪
  await sleep(400)
  await waitImages(container)

  const wrappers = container.querySelectorAll('.pptx-preview-slide-wrapper')
  const els = wrappers.length ? Array.from(wrappers) : Array.from(container.children)
  log(`检测到 ${els.length} 个幻灯片元素，逐页出图 (scale=${scale})…`)

  const images = []
  for (let i = 0; i < els.length; i++) {
    const t0 = performance.now()
    let dataUrl = null
    try {
      // 注意：snapdom 的 reconcile:true 可让文本像素级精确，但实测本样本 11 页从 4.1s 涨到 93.6s(约23x)，
      // 故默认关闭；若某些 deck 文本换行明显，再按需开启或改用"确保字体可用"的其他手段。
      const canvas = await snapdom.toCanvas(els[i], { scale })
      dataUrl = canvas.toDataURL('image/jpeg', quality)
      const w = canvas.width, h = canvas.height
      canvas.width = 0
      canvas.height = 0
      log(`  第 ${i + 1}/${els.length} 页 ✓ ${w}x${h} ${((performance.now() - t0) | 0)}ms`)
    } catch (e) {
      log(`  第 ${i + 1}/${els.length} 页 ✗ 渲染失败: ${e.message}（该页将留空白）`)
    }
    images.push(dataUrl)
    addThumb(leftCol, dataUrl, i + 1)
    await tick()
  }

  try { previewer.destroy && previewer.destroy() } catch (_) {}
  container.remove()
  return images
}

function addThumb(col, dataUrl, idx) {
  const d = document.createElement('div')
  d.className = 'thumb'
  if (dataUrl) {
    const img = document.createElement('img')
    img.src = dataUrl
    d.appendChild(img)
  } else {
    const bx = document.createElement('div')
    bx.style.cssText = 'aspect-ratio:16/9;background:#eee;'
    d.appendChild(bx)
  }
  const cap = document.createElement('div')
  cap.className = 'cap'
  cap.textContent = `第 ${idx} 页${dataUrl ? '' : '（空白）'}`
  d.appendChild(cap)
  col.appendChild(d)
}

// ---- 3. 以图片作背景重建 PPTX ----
async function buildPptx(images, sld) {
  const pptx = new PptxGenJS()
  const wIn = sld.cx / EMU_PER_INCH
  const hIn = sld.cy / EMU_PER_INCH
  log(`设置页面尺寸: ${wIn.toFixed(2)} x ${hIn.toFixed(2)} 英寸`)
  pptx.defineLayout({ name: 'ORIG', width: wIn, height: hIn })
  pptx.layout = 'ORIG'

  for (let i = 0; i < images.length; i++) {
    const slide = pptx.addSlide()
    if (images[i]) {
      slide.background = { data: images[i] } // ← 真背景 p:bg/blipFill
    } else {
      slide.background = { color: 'FFFFFF' }
    }
  }
  const blob = await pptx.write({ outputType: 'blob' })
  return blob
}

// ---- 4. 输出结构自检（客观判据）----
async function verifyOutput(blob, expectedSlides) {
  const zip = await JSZip.loadAsync(blob)
  const slideNames = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
  const media = Object.keys(zip.files).filter((n) => /^ppt\/media\/.+\.(jpe?g|png)$/i.test(n))
  let withBg = 0
  for (const n of slideNames) {
    const xml = await zip.files[n].async('string')
    if (xml.includes('<p:bg>') && xml.includes('<a:blipFill')) withBg++
  }
  const ok = slideNames.length === expectedSlides && withBg === expectedSlides && media.length >= withBg
  const lines = [
    `输入页面数(渲染成功): ${expectedSlides}`,
    `输出 slideXML 数:      ${slideNames.length}`,
    `含 <p:bg>+<a:blipFill> 的页: ${withBg}`,
    `嵌入 media 图片数:     ${media.length}`,
    `输出文件大小:          ${fmtBytes(blob.size)}`,
    '',
    ok ? '判定: ✅ PASS —— 每页均为真背景图片填充' : '判定: ❌ FAIL（见上）',
  ]
  verifyEl.textContent = lines.join('\n')
  return ok
}

// ---- 5. 回读输出 pptx 渲染（验证背景能否被预览器呈现）----
async function renderOutput(blob) {
  rightCol.innerHTML = ''
  const buf = await blob.arrayBuffer()
  const sld = await (async () => {
    const zip = await JSZip.loadAsync(buf)
    return readSlideSize(zip)
  })()
  const height = Math.round((RENDER_WIDTH * sld.cy) / sld.cx)
  const container = document.createElement('div')
  stage.appendChild(container)
  try {
    const previewer = pptxPreviewInit(container, { width: RENDER_WIDTH, height, mode: 'list' })
    await previewer.preview(buf)
    await sleep(400)
    await waitImages(container)
    const wrappers = container.querySelectorAll('.pptx-preview-slide-wrapper')
    log(`输出回读: 检测到 ${wrappers.length} 个幻灯片元素`)
    if (!wrappers.length) {
      log('注意: 预览器可能未实现背景渲染，回读为空白属正常——以左侧“结构自检”为准')
    }
    for (let i = 0; i < wrappers.length; i++) {
      try {
        const canvas = await snapdom.toCanvas(wrappers[i], { scale: 1 })
        addThumb(rightCol, canvas.toDataURL('image/jpeg', 0.8), i + 1)
        canvas.width = canvas.height = 0
      } catch (_) {
        addThumb(rightCol, null, i + 1)
      }
      await tick()
    }
    try { previewer.destroy && previewer.destroy() } catch (_) {}
  } finally {
    container.remove()
  }
}

// ---- 事件 ----
fileInput.addEventListener('change', async (e) => {
  const f = e.target.files[0]
  if (!f) return
  if (!f.name.toLowerCase().endsWith('.pptx')) {
    alert('PoC 仅支持 .pptx')
    return
  }
  if (f.size > MAX_INPUT_BYTES) {
    alert(`文件 ${fmtBytes(f.size)} 超过上限 ${fmtBytes(MAX_INPUT_BYTES)}`)
    return
  }
  selectedFile = f
  sourceArrayBuffer = await f.arrayBuffer()
  $('filemeta').textContent = `${f.name} · ${fmtBytes(f.size)}`
  convertBtn.disabled = false
  downloadLink.style.display = 'none'
  log(`已选择: ${f.name} (${fmtBytes(f.size)})`)
})

convertBtn.addEventListener('click', async () => {
  if (!sourceArrayBuffer) return
  convertBtn.disabled = true
  downloadLink.style.display = 'none'
  leftCol.innerHTML = ''
  rightCol.innerHTML = ''
  verifyEl.textContent = '（运行中…）'
  const scale = parseFloat($('scale').value) || 2
  const quality = parseFloat($('quality').value) || 0.92
  const t0 = performance.now()
  try {
    const zip = await JSZip.loadAsync(sourceArrayBuffer)
    await guardZipBomb(zip)
    const sld = await readSlideSize(zip)
    const fileSlides = await countSlideFiles(zip)
    if (!sld.found) log('警告: 未找到 <p:sldSz>，使用兜底 16:9 尺寸')
    log(`页面尺寸 EMU: ${sld.cx} x ${sld.cy}；slide XML 文件数: ${fileSlides}`)

    if (document.fonts && document.fonts.ready) {
      await document.fonts.ready
      log('字体已就绪 (document.fonts.ready)')
    }

    const images = await renderAndRasterize(sourceArrayBuffer, sld, scale, quality)
    const okCount = images.filter(Boolean).length
    if (!okCount) throw new Error('没有任何页面成功渲染')

    log('重建 PPTX（图片作背景）…')
    lastOutputBlob = await buildPptx(images, sld)
    const base = selectedFile.name.replace(/\.pptx$/i, '')
    downloadLink.href = URL.createObjectURL(lastOutputBlob)
    downloadLink.download = `${base}_image.pptx`
    downloadLink.textContent = `下载结果 PPTX (${fmtBytes(lastOutputBlob.size)})`
    downloadLink.style.display = 'inline-block'

    log('输出结构自检…')
    await verifyOutput(lastOutputBlob, images.length)

    log('回读输出 PPTX…')
    await renderOutput(lastOutputBlob)

    log(`✅ 完成，用时 ${(((performance.now() - t0) / 1000)).toFixed(1)}s（成功 ${okCount}/${images.length} 页）`)
  } catch (e) {
    log(`❌ 失败: ${e.message}`)
    console.error(e)
    verifyEl.textContent = `失败: ${e.message}`
  } finally {
    convertBtn.disabled = false
  }
})

log('PoC 就绪。请选择 .pptx 文件。')
