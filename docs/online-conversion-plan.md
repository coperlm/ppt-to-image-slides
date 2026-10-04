# PPT → 图片背景版 PPTX · 纯前端在线化实施计划（可执行版）

> 状态：`web/` 已实现并在浏览器跑通全流水线（2026-10-04）；部署工作流已就位，待合并到 `main` 触发
> 目标：部署在 GitHub Pages 的纯前端应用。用户本地选 `.pptx` → 浏览器内渲染出图 → 以图片作**幻灯片真背景**重建单个 `.pptx` → 下载。文件全程不出浏览器。
> 本文档中与实际实现不一致的地方已按实现更正，并在 §2 记录了 PoC 阶段的一处测量误判。

---

## 0. 一句话结论

纯前端可行，已验证。唯一真正的变量是 `pptx-preview` 的渲染保真度（需语料评测）。其余（解析、出图、回包、部署、内存管控）均为成熟工程。

---

## 1. 范围

**做**：单个 `.pptx` 输入（OOXML，未加密）；单个 `.pptx` 输出（每页一张全屏图作背景）；纯前端静态托管；输入上限默认 100 MB、页数上限 200（均为**浏览器内存自限**，不是 GitHub 平台限制——Pages 的 100 MiB 限制针对仓库文件，与客户端能处理多大无关）。
**不做**：`.ppt`（97-2003 二进制，OLE2/CFB，浏览器内无解析+排版引擎，见 §14）；PDF / 图片 ZIP 输出；像素级一致（目标是「肉眼难辨」）；母版主题 / 动画与切换效果的保留。
**备注**：以**纯文本**保留——从输入的 `notesSlides` 抽取正文占位符文本，回填到 PptxGenJS 输出包自带的空 `notesSlide` 部件（rels 与 Content_Types 不动），格式不保留。

> 与桌面版的差异：`main.py` 把原文件重新打开当模板，所以备注/母版/切换效果**连格式一起**保留；web 版只保备注文本。若将来要连格式保留，路径是绕过 PptxGenJS 自行用 JSZip 组包（见 §9），把输入的 `ppt/notesSlides/*` 及其 rels 原样注入。

---

## 2. PoC 实测基线（真实浏览器）

样本 `PPT_test.pptx`（456.3 KB / 11 页 / 16:9 `12192000×6858000` EMU），Chromium：

| 配置 | 每页分辨率 | 每页耗时 | 总耗时 | 输出体积 | 结构自检 |
|---|---|---|---|---|---|
| scale=2, q=0.92 | 3840×2160 | 166–361 ms | 4.1 s | 6.58 MB | ✅ 11/11 `<p:bg>`+`<a:blipFill>` |
| scale=1, q=0.80 | 1920×1080 | 88–219 ms | 2.8 s | 1.89 MB | ✅ PASS |

**已确认事实（约束后续设计）**：
1. 无需 WASM / COOP-COEP → GitHub Pages 可承载。
2. `pptx-preview` **不渲染背景** → 回读预览必然空白，**视觉验证只能靠结构自检 + PowerPoint/WPS 人工确认**。
3. snapdom `reconcile:true` 实测 **23× 变慢**（4.1s→93.6s）→ 默认关闭。
4. snapdom 1.x 有保真 bug，须用 `^2.22`；仍会警告「文本可能因字体回退换行」。
5. 打包产物 1.81 MB（gzip 591 KB），大头是 `pptx-preview` 带的 `echarts`（单文件 1010 KB，占 56%）→ 需代码分割。
6. **上表的「每页分辨率」不是由 scale 单独决定的**：snapdom 的出图基准是 `display box × scale × dpr`。PoC 在 `devicePixelRatio=1.5` 的机器上跑，`RENDER_WIDTH=1280 × scale 2 × dpr 1.5 = 3840`，于是被记成「scale=2 → 3840×2160」。这意味着**同一档位在 HiDPI 与普通屏上会输出不同分辨率和体积**，不可接受 → 实现中显式传 `dpr: 1`，档位改由「输出长边像素」定义（见 §6）。实测确认：100×60 元素在 dpr=1.5 环境下 `{scale:2, dpr:1}` 产出 200×120。
7. **转换需要标签页处于前台**：出图依赖浏览器渲染帧。实测同一个 11 页样本：前台总耗时 3.8 s；后台（`visibilityState: hidden`）时**第 1 页耗了 231 s**，之后每页恢复正常的 130–310 ms，总计 234 s。另一次后台运行超过 5 分钟后被 Chrome **冻结**，卡在首页 35 分钟无进展（隐藏页约 5 分钟后进入 freezing，需用户切回才解冻）。UI 已在 `visibilitychange` 时提示「已暂停」，页脚也写明了这条限制。
8. 页间让出主线程必须用 `MessageChannel`，不能用 `requestAnimationFrame`：rAF 在后台标签页完全不触发，会让循环永久停在两页之间（实测踩到）。

---

## 3. 技术栈与依赖（锁定版本）

```jsonc
// web/package.json —— 实际锁定靠 package-lock.json + CI 的 npm ci（`^` 只是允许范围）
// dependencies（括号内为 lockfile 解析到的版本）
"jszip":          "^3.10.1",   // 3.10.2   读解 pptx(zip)
"pptx-preview":   "^1.0.7",    // 1.0.7    渲染每页到 DOM；license ISC、个人维护，间接带入 echarts/lodash/uuid
"@zumer/snapdom": "^2.24.18",  // 2.24.18  DOM→图片（勿用 1.x）；出图基准 = display box × scale × dpr
"pptxgenjs":      "^3.12.0"    // 3.12.0   生成 pptx，支持 slide.background={data}
// devDependencies
"vite":         "^5.4.21",     // 5.4.21
"typescript":   "^5.6.3"       // 5.9.3；npm run build = tsc --noEmit && vite build
```

> 测试已补齐：`vitest`（jsdom 环境）单测 `guards` / `verifier` / `notes` 共 16 例（含用 JSZip 现造假包验证 magic/扩展名判据与备注注入）；`@playwright/test` 冒烟 5 例（整条流水线、非 pptx 拒绝、CFB 拒绝、英文切换、备注保留），headless Chromium 约 17 s。两者都进了 CI（见 §11）。注意本机 shell 若带 `NODE_ENV=production`，npm 会剪掉 devDependencies，本地跑测试要 `npm install --include=dev`。

---

## 4. 目录结构

```
web/
  index.html                # 单页 UI + 已知限制说明
  vite.config.ts            # base:'./'，dev 端口 5174
  src/
    main.ts                 # 装配 UI 与 pipeline（拖放、档位、进度、取消、缩略图、下载）
    config.ts               # 档位（按输出长边像素）、上限、RENDER_WIDTH
    format.ts               # fmtBytes / fmtSeconds / yieldToBrowser / throwIfAborted
    guards.ts               # 输入校验 + 容器魔数 + zip 炸弹防御 + sldSz/sldIdLst 解析
    renderer.ts             # pptx-preview 渲染 + 就绪判定（图片加载完 + DOM 静默期）
    rasterizer.ts           # snapdom 出图（dpr:1）+ 缩略图 + 失败页占位图
    packer.ts               # PptxGenJS 背景回包
    verifier.ts             # 输出结构自检（r:embed → rels → 包内 JPEG/PNG）
    pipeline.ts             # 编排（串行/进度/取消/失败页/内存释放顺序）
    styles.css
.github/workflows/deploy-pages.yml
docs/online-conversion-plan.md
poc/                        # 已验证原型，保留（poc/dist 不再纳入版本控制）
```

---

## 5. 模块接口（TypeScript 签名）

```ts
// config.ts —— 档位以「输出长边像素」定义，不再用 scale × 隐式基准
export type QosName = 'clear' | 'balanced' | 'small'
export interface Qos { name: QosName; label: string; hint: string; targetLongEdge: number; quality: number }
export const QOS: Record<QosName, Qos>
export const DEFAULT_QOS: QosName                  // 'balanced'
export const RENDER_WIDTH: number                  // 1280，pptx-preview 的布局基准宽度
export const LIMITS = {
  maxInputBytes: 100 * 1024 * 1024,
  maxUncompressedBytes: 600 * 1024 * 1024,
  maxSlides: 200,
  thumbLongEdge: 320,
}

// guards.ts —— 页数以 presentation.xml 的 sldIdLst + rels 为准（数 slideN.xml 不可靠）
export interface SlideSize { cx: number; cy: number; found: boolean }
export interface Guarded {
  buffer: ArrayBuffer
  zip: JSZip
  slideCount: number
  slidePaths: string[]
  slideCountSource: 'sldIdLst' | 'fileScan'        // fileScan 为回退路径，UI 会告警
  sldSz: SlideSize
  uncompressedBytes: number
}
export function guardAndRead(file: File, limits = LIMITS): Promise<Guarded>
export function normalizePartPath(baseDir: string, target: string): string   // OPC 关系 Target 归一化

// renderer.ts
export interface Rendered { slideEls: HTMLElement[]; destroy(): void }
export function renderPptx(buffer: ArrayBuffer, sldSz: SlideSize): Promise<Rendered>

// rasterizer.ts —— 固定 dpr:1；失败页有占位图，保证页数与背景判据仍成立
export interface Raster { dataUrl: string; thumbUrl: string; width: number; height: number }
export function rasterizeSlide(el: HTMLElement, qos: Qos): Promise<Raster>
export function placeholderSlide(cx: number, cy: number, label: string): Raster

// packer.ts
export function packBackgroundPptx(images: string[], sldSz: SlideSize): Promise<Blob>

// verifier.ts —— 判据是「r:embed 经 rels 解析到包内真实存在的 JPEG/PNG」，不是字符串匹配
export interface SlideCheck { index: number; path: string; mediaPath: string | null; ok: boolean }
export interface VerifyResult {
  expected: number; slideCount: number; withBackground: number; mediaCount: number
  slides: SlideCheck[]; problems: string[]; ok: boolean
}
export function verifyOutputPptx(blob: Blob, expectedSlides: number): Promise<VerifyResult>

// pipeline.ts
export interface PageResult { index: number; ok: boolean; ms: number; width: number; height: number; thumbUrl: string; error?: string }
export interface Stats {
  totalMs: number; guardMs: number; renderMs: number; rasterMs: number
  packMs: number; verifyMs: number; outputBytes: number; okPages: number; failedPages: number[]
}
export interface Hooks {
  onStage?(stage: string): void
  onLog?(message: string): void
  onProgress?(done: number, total: number, page: PageResult): void
}
export function convert(file: File, qos: Qos, hooks?: Hooks, signal?: AbortSignal): Promise<ConvertResult>
```

**与原草案的关键差异**：
- `Guarded.slideCount` 来自 `sldIdLst`，并在 pipeline 里与**渲染出的 DOM 页数**强校验；不一致直接中止，避免「少渲染一页却仍然 PASS」。
- 出图失败的页不再留 `null`（纯色背景没有 `blipFill`，会让硬判据必然 FAIL），而是替换为**带红框提示的占位图**：页数与背景判据仍成立，同时用户在预览里一眼能看到是哪页失败。
- `Rendered.destroy()` 在**打包之前**调用，先释放整副 deck 的渲染 DOM 再进 PptxGenJS。
- 全部图片的 base64 只在 pipeline 局部数组里存活，打包完成后立即清空；UI 侧只保留 320px 缩略图。

---

## 6. 质量档位（QoS 预设）

| 档位 | 输出长边 | JPEG 质量 | 16:9 单页像素 | 11 页样本 |
|---|---|---|---|---|
| 清晰 clear | 3840 px | 0.92 | 3840×2160 | 待实测 |
| 均衡 balanced（默认） | 2560 px | 0.85 | 2560×1440 | 待实测 |
| 小体积 small | 1920 px | 0.78 | 1920×1080 | 待实测 |

倍率按 `targetLongEdge / max(el.offsetWidth, el.offsetHeight)` 反推，并显式传 `dpr: 1`，因此输出像素在任何显示器缩放比例下都一致（见 §2-6）。

> **不设 PNG / 文字优先档**（2026-10-04 实测，见 `web/tools/fidelity/corpus.md`）：2560×1440 文字+表格页 JPEG q0.85 = 177 KB、PNG = 244 KB，且 3× 放大下两者肉眼无差别。PNG 在该场景既不小也不更清楚，照片页只会更差。

> 目前唯一实测数据点来自修正前的一次运行：11 页样本、档位标称「均衡 2560px」，但因 dpr=1.5 实际输出 3840×2160、JPEG 质量 0.85 → **5.32 MB / 3.8 s**（渲染 0.7 s、出图 2.7 s、打包 0.3 s）。dpr 修正后各档的体积与耗时需在 M2 重新实测。

---

## 7. 数据流

```
File(.pptx)
 → guards: 扩展名 + 容器魔数(PK / 排除 CFB 加密与 .ppt) + size≤100MB + 解压总量≤600MB + 页数≤200
 → 解析 ppt/presentation.xml：<p:sldSz cx cy>；页数与页序取 <p:sldIdLst> 的 r:id，
   经 ppt/_rels/presentation.xml.rels 解析到实际 slide 部件（解析不出才回退为扫描 slideN.xml 并告警）
 → await document.fonts.ready
 → renderer: pptx-preview.preview(buffer) → 等 <img> 全部加载 + MutationObserver 静默期（不是固定 sleep）
            → 取 .pptx-preview-slide-wrapper[]
 → 强校验：渲染出的页数 == guards 的页数，否则中止（不产出缺页文件）
 → 逐页(串行):
       scale = targetLongEdge / 页面长边，snapdom.toCanvas(el, { scale, dpr: 1 }) → JPEG dataURL
       顺带产出 320px 缩略图给 UI；释放大 canvas；onProgress
       (单页异常 → 生成带红框提示的占位图，页数与背景判据仍成立，并告警)
       (每页之间用 MessageChannel 让出主线程；不用 rAF —— 后台标签页 rAF 不触发)
 → 释放渲染 DOM（destroy）后再打包，压低内存峰值
 → packer: pptx.defineLayout({w:cx/914400,h:cy/914400}); 每页 slide.background={data}
 → blob → 清空 base64 数组 → verifier
 → verifier 通过才放出下载链接：<name>_image.pptx
```

---

## 8. 输出规格与硬性判据

- 单文件 `<name>_image.pptx`，页数 = 输入页数，每页仅一张全屏背景图。
- **硬性判据（自动，不通过就不放下载按钮）**：
  1. `输出 slide 部件数 == 输入页数`（输入页数取自 `sldIdLst`，不是数文件）；
  2. 每页 `<p:bg>` 下的 `<a:blip>` 的 `r:embed` 经该页 rels **解析到包内真实存在的部件**；
  3. 该部件扩展名是 JPEG/PNG（不用 WebP，PowerPoint 支持不稳）；
  4. `ppt/media` 图片数 ≥ 带背景的页数。
- 判据用 `DOMParser` + 命名空间查询实现，不用 `xml.includes('<p:bg>')` 这类字符串匹配（前缀或属性一变就失效）。
- **背景图的真实字节必须与声明的扩展名一致**（magic number 对 `.jpeg`/`.png`）。这条是实测踩出来的：PptxGenJS 在只给 `background.data` 时把扩展名硬编码成 `png`（`pptxgen.es.js:2727`），于是 JPEG 字节被写成 `Slide-1-image-1.png`、`[Content_Types].xml` 声明 `image/png` —— 正是它自己注释里写的会触发 PowerPoint 启动时内容警告的情形。修法：同时传 `path: 'slide-N.jpeg'`（有 `data` 时 PptxGenJS 不会去 fetch `path`，只用它推导扩展名，见 `pptxgen.es.js:4886`）。
- 统计 media 数时要排除 JSZip 的目录条目 `ppt/media/`，否则会多算一个。
- 出图失败的页用占位图补齐，因此「有失败页」不会让判据 FAIL，而是以 `degraded` + 失败页码显式呈现给用户，由用户决定是否下载。

---

## 9. 性能与内存策略

**已实现**：
- **串行逐页**，每页出图后立即 `canvas.width = canvas.height = 0`；不并发持有多个大 canvas。
- **打包前先 `destroy()` 释放整副 deck 的渲染 DOM**，再进 PptxGenJS；base64 数组是 pipeline 局部变量，打包完成后立刻清空，UI 只保留 320px 缩略图。
- 页间让出主线程用 **MessageChannel**，不用 `requestAnimationFrame`：rAF 在后台标签页不触发，会让长转换永久挂起（实测踩到）。
- 支持 `AbortSignal` 取消（在页边界生效）。
- 选中文件后**空闲预取**重 chunk（`requestIdleCallback` 里动态 import），让「开始转换」少等一次网络；不碰首屏体积。
- 每页出图失败最多尝试 3 次（间隔 250 ms）再落占位页：瞬时故障（字体/解码竞态）有机会自愈，确定性故障只多花两次尝试。
- 选文件后即给出体积/耗时**区间预估**（`estimate.ts`）：系数取本机历史实测中位数，无历史时用保守常数并标注「首次为保守估计」。
- 代码分割：`pptx-preview` / `snapdom` / `pptxgenjs` 全部动态 `import()`，首屏只含应用代码 + JSZip。实测首屏 **114.08 KB（gzip 38.50 KB）**，懒加载块 gzip 分别为 pptx-preview 404.55 KB、pptxgenjs 97.87 KB、snapdom 53.73 KB。

**已排除的方案**：
- `URL.createObjectURL(blob)` + `slide.background={path:url}` 替代 base64：PptxGenJS 在浏览器端对 `path` 同样是 fetch → 转 dataURL 再入 zip，base64 照样驻留，拿不到「省 ~33% 字符串」的收益。真要省，得绕过 PptxGenJS 自己用 JSZip 直接写二进制 `ppt/media/*` + 手写 slide XML（顺带还能保留备注，见 §14-6）。

**待办 / 未验证**：
- 内存峰值只在 11 页 / 456 KB 样本上跑过；**100 MB 或上百页输入的峰值仍未实测**，`maxSlides` 暂设 200 是保守值，需要实测后再定。
- 编码搬到 Web Worker（`createImageBitmap` + `OffscreenCanvas.convertToBlob`）未做：snapdom 必须在 DOM 侧，能搬走的只有 JPEG 编码那一段，收益待评估。
- ~~逐页渲染 spike~~ **已落地**：renderer 改用 `mode:'slide'` + `renderSingleSlide(i)` + `removeCurrentSlide()`，DOM 中同时只驻留一页，内存峰值与页数解耦。已确认翻页按钮/分页器挂在 slide wrapper 的**同级**，只截 wrapper 不会烤进背景图；列表模式则完全不加分页器。headless Chromium 端到端 7.2 s 通过。

---

## 10. 安全与隐私

- 文件不上传，纯本地；适合学术/涉密内容（写入文案）。
- 处理不可信 ZIP+XML：已设解压总量与规模上限；解析失败优雅降级并提示。

---

## 11. 部署（GitHub Actions → Pages）

```yaml
# .github/workflows/deploy-pages.yml（已落地，与实际文件一致）
name: Deploy web to GitHub Pages
on:
  push:
    branches: [main]
    paths: ['web/**', '.github/workflows/deploy-pages.yml']   # 改 Python 端不触发重新部署
  workflow_dispatch:
permissions: { contents: read, pages: write, id-token: write }
concurrency: { group: pages, cancel-in-progress: true }
jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/configure-pages@v5
        with: { enablement: true }        # 首次运行自动把 Pages 源设为 GitHub Actions
      - uses: actions/setup-node@v4
        with: { node-version: 20, cache: npm, cache-dependency-path: web/package-lock.json }
      - run: npm ci --include=dev          # vite/tsc 在 devDependencies；防 runner 带 NODE_ENV=production
        working-directory: web
      - run: npm run build                 # = tsc --noEmit && vite build，类型检查进 CI
        working-directory: web
      - uses: actions/upload-pages-artifact@v3
        with: { path: web/dist }
  deploy:
    needs: build
    runs-on: ubuntu-latest
    environment: { name: github-pages, url: "${{ steps.deployment.outputs.page_url }}" }
    steps:
      - id: deployment
        uses: actions/deploy-pages@v4
```

- 无需自定义响应头（无 WASM/SAB）。
- 单页工具没有路由，不需要 `404.html` 兜底。
- `vite.config.ts` 里 `base: './'`，项目页子路径下资源可正常加载。
- 上线前置：仓库 Settings → Pages → Source 选 **GitHub Actions**（`configure-pages` 的 `enablement: true` 会尝试自动开启）；工作流只在 `main` 上触发，功能分支需先合并。
- `web/dist`、`node_modules` 已进 `.gitignore`；`poc/dist` 曾在提交 `8cbe0a6` 里被误纳入版本控制，已从索引移除（磁盘文件保留）。

---

## 12. 里程碑（可勾选）

- [x] **M1 工程化**：`web/`（Vite+TS）已按 §5 建好，档位、串行流水线、进度/日志/缩略图/自检面板、取消均已实现；`PPT_test.pptx` 11 页跑通，结构自检 PASS（11/11 背景关系解析到包内 JPEG）。
     测试已补齐并进 CI（见 §3 注）。
- [ ] **M2 保真度评测与调优**（2–3 天）：按 §13 语料集逐类比对，记录差异；调档位与字体策略。
     交付：评测报告（逐类 pass/fail + 截图）+ 修正后各档的实测体积/耗时（§6 目前是待实测）。
     验收：文字/图文/表格/图表四类「肉眼可辨为一致」。
     **前置条件（原文档漏了）**：需要能打开 PowerPoint 或 WPS 的环境做 ground truth；当前开发机是 Arch Linux，没有 PowerPoint。
- [~] **M3 健壮性**：容器魔数（含加密/`.ppt` 的 CFB 识别）、页数与体积上限、失败页占位、取消、错误提示、渲染页数与文件页数强校验 —— 均已实现。
     未做：zip 炸弹防御仍依赖 JSZip 私有字段 `_data.uncompressedSize`（升级即碎），且信的是压缩包内声明值；真正稳妥的做法是解压时按累计字节流式截断。超大/损坏文件的边界用例也未实测。
- [x] **M4 部署与优化**：工作流已就位；代码分割已生效，首屏 gzip 38.19 KB。
     注：原验收标准「产物 < 1 MB gzip」在改造前就已满足（旧单包 gzip 570 KB），不是有效判据；现以「首屏 ≤ 250 KB gzip 且重依赖必须懒加载」为准，已达成。
     遗留：尚未真正跑过一次 Actions 部署（需合并到 `main`）。
- [~] **M5 文档**：README 已加 Web 版说明、与桌面版差异对照、已知限制；应用内已写明隐私与限制。计划书已同步实现。

**当前实际状态**：代码与部署配置完成，剩下 (1) 一次真实 Actions 部署，(2) M2 保真度评测，(3) 大文件内存实测与测试补齐。

---

## 13. 保真度验收流程（先定判据）

**语料集（≥10 个 deck，覆盖）**：纯文字（中英混排）/ 图文混排 / 表格 / 图表（柱饼折线）/ SmartArt / 艺术字 / 公式 / 大字报 / 复杂母版 / WPS 导出的 pptx。

**步骤**：
1. 应用内转换 → 记录耗时、体积、结构自检结果（必须 PASS）。
2. 输出 pptx 在 PowerPoint/WPS 打开，逐页截图。
3. 与**原始 pptx 在 PowerPoint 的原貌**逐页比对。
4. 每类记录：一致 / 轻微偏差(可接受) / 明显差异(不可接受)。

**判定**：核心四类（文字/图文/表格/图表）必须「一致或轻微可接受」；SmartArt/艺术字/公式允许标注为已知限制。

---

## 14. 开放风险

1. **客户端字体不可控（公网部署后的头号保真度风险）**：渲染用的是**访问者本机**的字体。Windows 有微软雅黑/等线，macOS 是苹方，Linux 常缺 → 同一份 pptx 在不同机器上换行与溢出程度不同，本机 Chromium 上测出的保真度不代表用户所见。CJK 全量字体无法内嵌，不可根治。缓解：UI/README 明示；文字型 deck 可考虑提供 `reconcile` 开关（带耗时预警）；M2 语料至少在两台字体环境不同的机器上各跑一遍。
2. **只在 Chromium 上验证过**：snapdom 走 SVG foreignObject，Safari 在这条路上历史问题最多，Firefox 性能也不同，iOS Safari 还有 canvas 面积上限。M2 需补浏览器矩阵（Chrome / Firefox / Safari 各一）。
3. `pptx-preview` 保真度上限（M2 前置信度：中）；SmartArt/艺术字/公式为主要风险。图表由 echarts 异步绘制，就绪判定已改为「图片加载完 + DOM 静默期」而非固定 sleep，但仍需 M2 在图表语料上确认不会截到半成品。
4. 内存峰值只在 11 页小样本上验证过；**100 MB / 上百页未实测**。逐页渲染已落地（§9），峰值与页数解耦，但大输入的绝对峰值仍待实测后再定 `maxSlides`。
5. **转换要求标签页在前台**：后台标签页会被节流到极慢（实测首页 231 s，见 §2-7），但不会丢进度。已在 UI 与页脚提示；未验证的是「后台几十分钟后再切回」的恢复情况。
6. 母版/动画/切换效果丢失（见 §1）。备注已改为**纯文本保留**（格式不保），讲课场景的主要回退已消除。
7. 应用内无法视觉自验背景渲染（`pptx-preview` 不渲染背景）→ 已用「出图缩略图即嵌入内容 + 结构自检 + 下载前人工核对」的 UX 替代，并在页面上写明。
8. 单一渲染依赖 `pptx-preview`：实装 1.0.7，license **ISC**，作者为个人（package.json 的 author 字段是微信号），无组织维护 → 锁 lockfile + 保留备选（`ChristopherVR/pptx-viewer`）；公开站点还应考虑依赖固定与构建可复现（CI 用 `npm ci`）。
9. 文本重排：`reconcile` 代价过高（23×）；备选方案（确保字体可用 / 换渲染库）待 M2 评测。
10. zip 炸弹防御依赖 JSZip 私有字段且信的是声明值（见 §12-M3 遗留）。

---

## 附：PoC 运行

```bash
cd poc && npm install && npm run dev   # http://127.0.0.1:5173
```
