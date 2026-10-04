export type Lang = 'zh' | 'en'

const zh = {
  appTitle: 'PPT → 图片背景版 PPTX',
  appSub: '在浏览器内把每页渲染成图片，再重建为「图片作幻灯片真背景」的 <code>.pptx</code>，避免换机器后字体缺失与排版错乱。',
  privacy: '文件不会上传：解析、渲染、出图、打包全部在本机浏览器内完成。',
  dropTitle: '选择或拖入 .pptx 文件',
  dropHint: '上限 {limit} · 仅支持未加密的 .pptx（不支持 .ppt）',
  qosLegend: '输出质量',
  qosRecommended: '（推荐）',
  qosClear: '清晰',
  qosClearHint: '长边 3840px · 投影与放大查看',
  qosBalanced: '均衡',
  qosBalancedHint: '长边 2560px · 清晰度与体积折中',
  qosSmall: '小体积',
  qosSmallHint: '长边 1920px · 便于网络传输',
  btnConvert: '开始转换',
  btnCancel: '取消',
  hVerify: '输出结构自检',
  hPreview: '页面预览',
  previewHint: '下面就是将被嵌入结果文件的图片（缩略图）。请确认无误后再下载 —— 应用内无法回读验证 pptx 的背景渲染效果。',
  hLog: '运行日志',
  hLimits: '已知限制',
  limitPptx: '仅支持未加密的 <code>.pptx</code>；不支持 97-2003 的 <code>.ppt</code>。',
  limitStatic: '输出为静态图片页：<strong>动画、切换效果、可编辑文本均不保留</strong>；演讲者备注以<strong>纯文本</strong>保留（格式不保留）。',
  limitFonts: '渲染使用<strong>你本机的字体</strong>。若系统缺少原稿字体（如微软雅黑、等线），文字可能换行或溢出。',
  limitMemory: '页数多、分辨率高的文件受浏览器内存限制；手机与低配设备请选「小体积」档。',
  limitForeground: '转换期间请<strong>保持本标签页在前台</strong>：浏览器会挂起后台标签页的渲染，转换将暂停（切回后自动继续）。',
  footerLine: '开源在 GitHub · 由 GitHub Actions 构建并部署到 GitHub Pages ·',
  footerSource: '查看源码',
  langLabel: '语言',

  statusUnsupported: '不支持的文件类型：{name}（仅支持 .pptx）',
  statusTooBig: '文件 {size} 超过上限 {limit}',
  statusPaused: '已暂停：浏览器会挂起后台标签页，切回本页即自动继续',
  statusStage: '{stage}（{name}）',
  statusRaster: '逐页出图 {done}/{total}（{name}）',
  statusDone: '完成，请核对预览后下载',
  statusDoneDegraded: '完成，但第 {pages} 页为占位页，请核对预览后再下载',
  statusCancelled: '已取消',
  statusFailed: '转换失败：{message}',
  statusVerifyFail: '结构自检未通过，已阻止下载（问题见上）',

  logReady: '就绪。选择 .pptx 后点击「开始转换」（上限 {limit} / {slides} 页）。',
  logSelected: '已选择：{name}（{size}）',
  logStart: '开始转换 · 档位「{qos}」长边 {edge}px · JPEG 质量 {quality}',
  logMeta: '页面尺寸 {cx}×{cy} EMU · 页数 {count}（来源 {source}） · 解压后约 {uncompressed}',
  logMetaSldSzMissing: '（未找到 sldSz，已按 16:9 兜底）',
  logWarnSldSz: '警告：缺少 <p:sldSz>，输出尺寸可能与原稿不一致',
  logWarnFileScan: '警告：无法解析 sldIdLst，页数已回退为按文件扫描统计',
  logNotes: '检测到 {count} 页演讲者备注，将以纯文本保留',
  logNotesInjected: '已把 {count} 页备注（纯文本）写回输出包',
  logFonts: '字体就绪（document.fonts.ready）',
  logRenderer: '渲染器就绪（逐页模式，DOM 中同时只保留一页），用时 {ms}',
  logPageFail: '第 {index} 页出图失败：{message} → 已用占位页替代',
  logWarnFailed: '警告：{count} 页出图失败（第 {pages} 页）',
  logOutput: '输出 {size}，打包用时 {ms}',
  logCancelled: '已取消',
  logFailed: '❌ {message}',

  stageGuard: '校验输入…',
  stageRender: '渲染幻灯片…',
  stageRaster: '逐页出图…',
  stagePack: '重建 PPTX…',
  stageVerify: '结构自检…',

  verifyNotRun: '（尚未运行）',
  verifyRunning: '（运行中…）',
  verifyInputPages: '输入页数:',
  verifySlideXml: '输出 slide XML 数:',
  verifyWithBg: '背景图解析成功的页:',
  verifyMedia: '包内 media 图片数:',
  verifySize: '输出文件大小:',
  verifyTime: '总耗时:                {total}（渲染 {render} / 出图 {raster} / 打包 {pack}）',
  verifyFailedPages: '出图失败并替换为占位页: 第 {pages} 页',
  verifyPass: '判定: ✅ PASS —— 每页背景的 r:embed 均解析到包内图片，且字节与扩展名一致',
  verifyFail: '判定: ❌ FAIL',
  verifyFailLine: '失败：{message}',

  capPage: '第 {index} 页 · {w}×{h} · {ms} ms',
  capPageFailed: '第 {index} 页 · 渲染失败（已用占位页）',
  download: '下载 {name}（{size}）',
  placeholder: '第 {index} 页渲染失败',

  errUnsupported: '不支持的文件类型：{name}（仅支持未加密的 .pptx）',
  errTooBig: '文件 {size} 超过上限 {limit}',
  errCfb: '文件已加密，或为 97-2003 的 .ppt 二进制格式；本工具仅支持未加密的 .pptx。请在 PowerPoint/WPS 中「另存为 .pptx」后重试',
  errNotZip: '不是有效的 .pptx（缺少 ZIP 文件头 PK）',
  errNoPresentation: '缺少 ppt/presentation.xml，不是有效的 .pptx',
  errNoSlides: '演示文稿中没有幻灯片',
  errTooManySlides: '幻灯片 {count} 页，超过上限 {limit} 页',
  errZipBomb: '解压后约 {size}，超过上限 {limit}（防 zip 炸弹，已拒绝）',
  errXml: 'XML 解析失败：{detail}',
  errRenderCount: '渲染页数 {rendered} 与文件页数 {expected} 不一致，已中止以免输出缺页',
  errZeroSize: '幻灯片元素尺寸为 0，无法出图',
  errNoCanvas: '浏览器拒绝创建 2D 画布',
  errNoElement: '第 {index} 页渲染后未找到页面元素',
  errAborted: '已取消转换',

  verUnreadable: '第 {index} 页：{path} 无法读取',
  verNoBg: '第 {index} 页：未找到 <p:bg> + <a:blipFill> 背景',
  verUnresolved: '第 {index} 页：背景关系 {id} 未解析到包内图片',
  verUnknownImage: '第 {index} 页：背景图 {path} 不是可识别的 JPEG/PNG',
  verMismatch: '第 {index} 页：背景图 {path} 声明为 {declared}，实际字节是 {actual}',
  verPageCount: '页数不符：期望 {expected} 页，实际输出 {actual} 页',
  verMediaCount: 'media 图片数 {media} 少于带背景的页数 {withBg}',
} as const

export type MessageKey = keyof typeof zh

const en: Record<MessageKey, string> = {
  appTitle: 'PPT → image-background PPTX',
  appSub: 'Renders every slide to an image in your browser and rebuilds a <code>.pptx</code> with those images as true slide backgrounds, so fonts and layout survive any machine.',
  privacy: 'Nothing is uploaded: parsing, rendering, rasterizing and packing all happen locally in your browser.',
  dropTitle: 'Choose or drop a .pptx file',
  dropHint: 'Up to {limit} · unencrypted .pptx only (.ppt is not supported)',
  qosLegend: 'Output quality',
  qosRecommended: ' (recommended)',
  qosClear: 'Clear',
  qosClearHint: '3840px long edge · for projection and zooming',
  qosBalanced: 'Balanced',
  qosBalancedHint: '2560px long edge · clarity vs. size',
  qosSmall: 'Small',
  qosSmallHint: '1920px long edge · easy to share',
  btnConvert: 'Convert',
  btnCancel: 'Cancel',
  hVerify: 'Structural self-check',
  hPreview: 'Page preview',
  previewHint: 'These thumbnails are exactly the images that will be embedded. Review them before downloading — the app cannot read back how PowerPoint renders the backgrounds.',
  hLog: 'Run log',
  hLimits: 'Known limitations',
  limitPptx: 'Unencrypted <code>.pptx</code> only; 97-2003 <code>.ppt</code> is not supported.',
  limitStatic: 'Output is static image pages: <strong>animations, transitions and editable text are not kept</strong>; speaker notes are kept as <strong>plain text</strong> (formatting dropped).',
  limitFonts: 'Rendering uses <strong>the fonts installed on your machine</strong>. Missing source fonts (e.g. Microsoft YaHei) may re-wrap or overflow text.',
  limitMemory: 'Large decks and high resolutions are bounded by browser memory; on phones or weak machines pick the "small" preset.',
  limitForeground: '<strong>Keep this tab in the foreground</strong> while converting: browsers suspend rendering in background tabs, so the conversion pauses until you return.',
  footerLine: 'Open source on GitHub · built and deployed to GitHub Pages by GitHub Actions ·',
  footerSource: 'View source',
  langLabel: 'Language',

  statusUnsupported: 'Unsupported file type: {name} (.pptx only)',
  statusTooBig: 'File {size} exceeds the {limit} limit',
  statusPaused: 'Paused: browsers suspend background tabs; switch back to this tab to resume',
  statusStage: '{stage} ({name})',
  statusRaster: 'Rasterizing {done}/{total} ({name})',
  statusDone: 'Done — review the preview, then download',
  statusDoneDegraded: 'Done, but page {pages} is a placeholder; review the preview before downloading',
  statusCancelled: 'Cancelled',
  statusFailed: 'Conversion failed: {message}',
  statusVerifyFail: 'Self-check failed; download blocked (see problems above)',

  logReady: 'Ready. Pick a .pptx and press Convert (limit {limit} / {slides} pages).',
  logSelected: 'Selected: {name} ({size})',
  logStart: 'Converting · preset "{qos}" long edge {edge}px · JPEG quality {quality}',
  logMeta: 'Slide size {cx}×{cy} EMU · pages {count} (source {source}) · uncompressed ≈ {uncompressed}',
  logMetaSldSzMissing: ' (<p:sldSz> not found, falling back to 16:9)',
  logWarnSldSz: 'Warning: <p:sldSz> missing; output size may differ from the source',
  logWarnFileScan: 'Warning: sldIdLst could not be parsed; page count fell back to file scan',
  logNotes: 'Found speaker notes on {count} page(s); they will be kept as plain text',
  logNotesInjected: 'Wrote {count} page(s) of plain-text notes back into the output package',
  logFonts: 'Fonts ready (document.fonts.ready)',
  logRenderer: 'Renderer ready (per-slide mode, one page in the DOM at a time) in {ms}',
  logPageFail: 'Page {index} failed to rasterize: {message} → replaced with a placeholder page',
  logWarnFailed: 'Warning: {count} page(s) failed (page {pages})',
  logOutput: 'Output {size}, packing took {ms}',
  logCancelled: 'Cancelled',
  logFailed: '❌ {message}',

  stageGuard: 'Validating input…',
  stageRender: 'Rendering slides…',
  stageRaster: 'Rasterizing pages…',
  stagePack: 'Rebuilding PPTX…',
  stageVerify: 'Structural self-check…',

  verifyNotRun: '(not run yet)',
  verifyRunning: '(running…)',
  verifyInputPages: 'Input pages:',
  verifySlideXml: 'Output slide XML parts:',
  verifyWithBg: 'Pages whose background resolved:',
  verifyMedia: 'Media images in package:',
  verifySize: 'Output file size:',
  verifyTime: 'Total:                {total} (render {render} / raster {raster} / pack {pack})',
  verifyFailedPages: 'Pages replaced by placeholders: {pages}',
  verifyPass: 'Verdict: ✅ PASS — every background r:embed resolves to an in-package image whose bytes match its extension',
  verifyFail: 'Verdict: ❌ FAIL',
  verifyFailLine: 'Failed: {message}',

  capPage: 'Page {index} · {w}×{h} · {ms} ms',
  capPageFailed: 'Page {index} · rasterize failed (placeholder used)',
  download: 'Download {name} ({size})',
  placeholder: 'Page {index} failed to render',

  errUnsupported: 'Unsupported file type: {name} (unencrypted .pptx only)',
  errTooBig: 'File {size} exceeds the {limit} limit',
  errCfb: 'The file is encrypted or a 97-2003 binary .ppt; only unencrypted .pptx is supported. Open it in PowerPoint/WPS and "Save As .pptx", then retry',
  errNotZip: 'Not a valid .pptx (missing the PK zip header)',
  errNoPresentation: 'Missing ppt/presentation.xml — not a valid .pptx',
  errNoSlides: 'The presentation contains no slides',
  errTooManySlides: '{count} slides exceeds the {limit}-slide limit',
  errZipBomb: 'Uncompressed size ≈ {size} exceeds the {limit} cap (zip-bomb defence, rejected)',
  errXml: 'XML parse failure: {detail}',
  errRenderCount: 'Rendered pages {rendered} != file pages {expected}; aborted to avoid missing pages',
  errZeroSize: 'Slide element has zero size; cannot rasterize',
  errNoCanvas: 'The browser refused to create a 2D canvas',
  errNoElement: 'No page element found after rendering page {index}',
  errAborted: 'Conversion cancelled',

  verUnreadable: 'Page {index}: cannot read {path}',
  verNoBg: 'Page {index}: no <p:bg> + <a:blipFill> background found',
  verUnresolved: 'Page {index}: background relationship {id} does not resolve to an in-package image',
  verUnknownImage: 'Page {index}: background {path} is not a recognizable JPEG/PNG',
  verMismatch: 'Page {index}: background {path} is declared {declared} but the bytes are {actual}',
  verPageCount: 'Page count mismatch: expected {expected}, output has {actual}',
  verMediaCount: 'media images {media} fewer than pages with backgrounds {withBg}',
}

const CATALOG: Record<Lang, Record<MessageKey, string>> = { zh, en }
const STORAGE_KEY = 'ppt2img-lang'

let current: Lang = readStoredLang()

function readStoredLang(): Lang {
  try {
    return localStorage.getItem(STORAGE_KEY) === 'en' ? 'en' : 'zh'
  } catch {
    return 'zh'
  }
}

export function getLang(): Lang {
  return current
}

export function setLang(lang: Lang): void {
  current = lang
  try {
    localStorage.setItem(STORAGE_KEY, lang)
  } catch {
    // 隐私模式下 localStorage 可能不可用，忽略
  }
}

export function t(key: MessageKey, params?: Record<string, string | number>): string {
  let text = CATALOG[current][key]
  if (params) {
    for (const [name, value] of Object.entries(params)) {
      text = text.replaceAll(`{${name}}`, String(value))
    }
  }
  return text
}
