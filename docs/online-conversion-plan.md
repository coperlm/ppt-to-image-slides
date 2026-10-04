# PPT → 图片背景版 PPTX · 纯前端在线化实施计划（可执行版）

> 状态：PoC 已通过（2026-10-04）
> 目标：部署在 GitHub Pages 的纯前端应用。用户本地选 `.pptx` → 浏览器内渲染出图 → 以图片作**幻灯片真背景**重建单个 `.pptx` → 下载。文件全程不出浏览器。

---

## 0. 一句话结论

纯前端可行，已验证。唯一真正的变量是 `pptx-preview` 的渲染保真度（需语料评测）。其余（解析、出图、回包、部署、内存管控）均为成熟工程。

---

## 1. 范围

**做**：单个 `.pptx` 输入（OOXML）；单个 `.pptx` 输出（每页一张全屏图作背景）；纯前端静态托管；输入上限默认 100 MB（可调）。
**不做**：`.ppt`（97-2003 二进制）；PDF / 图片 ZIP 输出；像素级一致（目标是「肉眼难辨」）。

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
5. 打包产物 1.81 MB（gzip 591 KB），大头是 `pptx-preview` 带的 `echarts` → 需代码分割。

---

## 3. 技术栈与依赖（锁定版本）

```jsonc
// web/package.json（dependencies）
"jszip":         "^3.10.1",   // 读解 pptx(zip)
"pptx-preview":  "^1.0.7",    // 渲染每页到 DOM（纯前端）
"@zumer/snapdom":"^2.24.0",   // DOM→图片（勿用 1.x）
"pptxgenjs":     "^3.12.0"    // 生成 pptx，支持 slide.background={data}
// devDependencies
"vite": "^5.4.0", "typescript": "^5.6.0"
```

---

## 4. 目录结构

```
web/
  index.html
  vite.config.ts            # base:'./'
  src/
    main.ts                 # 装配 UI 与 pipeline
    config.ts               # 上限、档位、常量
    guards.ts               # 输入校验 + zip 炸弹防御 + 解析元信息
    renderer.ts             # pptx-preview 渲染
    rasterizer.ts           # snapdom 出图
    packer.ts               # PptxGenJS 背景回包
    verifier.ts             # 输出结构自检
    pipeline.ts             # 编排（串行/进度/取消/失败页）
    ui/                     # 组件：文件区、档位、进度、日志、自检面板
.github/workflows/deploy-pages.yml
docs/online-conversion-plan.md
poc/                        # 已验证原型，保留
```

---

## 5. 模块接口（TypeScript 签名）

```ts
// config.ts
export type QosName = 'clear' | 'balanced' | 'small'
export interface Qos { name: QosName; scale: number; quality: number }
export const QOS: Record<QosName, Qos>          // 见 §6
export const LIMITS = {
  maxInputBytes: 100 * 1024 * 1024,
  maxUncompressedBytes: 600 * 1024 * 1024,
  maxSlides: 500,
  renderWidth: 1280,
}

// guards.ts
export interface SlideSize { cx: number; cy: number; found: boolean }
export interface Guarded {
  zip: JSZip; slideCount: number; sldSz: SlideSize
  uncompressedBytes: number
}
export function guardAndRead(file: File, limits = LIMITS): Promise<Guarded>

// renderer.ts
export interface Rendered { slideEls: HTMLElement[]; destroy(): void }
export function renderPptx(buf: ArrayBuffer, sldSz: SlideSize, renderWidth: number): Promise<Rendered>

// rasterizer.ts
export function rasterizeToJpegDataUrl(el: HTMLElement, qos: Qos): Promise<string>

// packer.ts
export function packBackgroundPptx(images: (string | null)[], sldSz: SlideSize): Promise<Blob>

// verifier.ts
export interface VerifyResult { slides: number; withBg: number; media: number; ok: boolean }
export function verifyOutputPptx(blob: Blob, expectedSlides: number): Promise<VerifyResult>

// pipeline.ts
export interface Stats { totalMs: number; perSlideMs: number[]; outputBytes: number; okPages: number }
export interface Hooks { onProgress?(done: number, total: number): void; onLog?(msg: string): void }
export function convert(file: File, qos: Qos, hooks: Hooks, signal?: AbortSignal)
  : Promise<{ blob: Blob; verify: VerifyResult; stats: Stats }>
```

---

## 6. 质量档位（QoS 预设）

| 档位 | scale | JPEG 质量 | 16:9 单页像素 | 预期体积(11页样本) |
|---|---|---|---|---|
| 清晰 clear | 2.0 | 0.92 | 3840×2160 | ≈6.6 MB |
| 均衡 balanced（默认） | 1.5 | 0.85 | 2880×1620 | ≈3–4 MB |
| 小体积 small | 1.0 | 0.75 | 1920×1080 | ≈1.5–1.9 MB |

> 体积随内容变化，上表为 `PPT_test.pptx` 外推值，正式版以实测为准。

---

## 7. 数据流

```
File(.pptx)
 → guards: 扩展名 + ZIP magic(PK\x03\x04) + size≤100MB + 解压总量≤600MB + 页数≤500
 → 读 ppt/presentation.xml 的 <p:sldSz cx cy>；统计 ppt/slides/slideN.xml 数
 → await document.fonts.ready
 → renderer: pptx-preview.preview(buf) → 取 .pptx-preview-slide-wrapper[]
 → 逐页(串行):
       dataUrl = rasterizeToJpegDataUrl(el, qos)
       释放 canvas；追加；onProgress
       (单页异常→该页记 null 并告警，不中断)
 → packer: pptx.defineLayout({w:cx/914400,h:cy/914400}); 每页 slide.background={data}
 → blob → 下载 <name>_image.pptx
 → verifier: 解压输出，校验 页数一致 & 每页含 <p:bg>+<a:blipFill> & media≥页数
```

---

## 8. 输出规格与硬性判据

- 单文件 `<name>_image.pptx`，页数 = 输入页数，每页仅一张全屏背景图。
- **硬性判据（自动，必须 100% 通过）**：`输出页数 == 输入页数` 且 `每页含 <p:bg>+<a:blipFill` 且 `media 图片数 ≥ 页数`。
- 背景图用 JPEG/PNG，**不用 WebP**（PowerPoint 支持不稳）。

---

## 9. 性能与内存策略

- **串行逐页**，每页出图后 `canvas.width=canvas.height=0` 释放；禁止并发持有多个大 canvas。
- 缓存优化（M1 评估）：用 `URL.createObjectURL(blob)` + `slide.background={path:url}` 替代 base64，减少 ~33% 字符串开销（需确认 PptxGenJS 对 blob URL 的 fetch 支持）。
- 大文件内存峰值 = 全部图片 + 最终 zip（PptxGenJS 非流式）。**100 MB 输入必须实测峰值**；超限时降档或提示。
- 编码放 Web Worker（`createImageBitmap`+`OffscreenCanvas`）；snapdom 属 DOM 侧留主线程 + 分块 `await` 让出。
- 支持 `AbortSignal` 取消。
- 代码分割懒加载 `pptx-preview`/`echarts`/`pptxgenjs`。

---

## 10. 安全与隐私

- 文件不上传，纯本地；适合学术/涉密内容（写入文案）。
- 处理不可信 ZIP+XML：已设解压总量与规模上限；解析失败优雅降级并提示。

---

## 11. 部署（GitHub Actions → Pages）

```yaml
# .github/workflows/deploy-pages.yml
name: Deploy web to GitHub Pages
on:
  push: { branches: [main] }
  workflow_dispatch:
permissions: { contents: read, pages: write, id-token: write }
concurrency: { group: pages, cancel-in-progress: true }
jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with: { node-version: 20, cache: npm, cache-dependency-path: web/package-lock.json }
      - run: npm ci
        working-directory: web
      - run: npm run build
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
- 路由用 hash 或 `404.html` 兜底。

---

## 12. 里程碑（可勾选）

- [ ] **M1 工程化**（2–3 天）：搭 `web/`（Vite+TS），按 §5 拆模块，接入 §6 档位，跑通单文件转换。
     交付：`web` 可 `dev` 跑通；模块接口与 §5 一致。
     验收：`PPT_test.pptx` 输出结构自检 PASS，UI 显示进度/日志/下载。
- [ ] **M2 保真度评测与调优**（2–3 天）：按 §13 语料集逐类比对，记录差异；调 scale/quality/字体策略。
     交付：评测报告（逐类 pass/fail + 截图）。
     验收：文字/图文/表格/图表四类「肉眼可辨为一致」。
- [ ] **M3 健壮性**（1–2 天）：zip 炸弹、页数/体积上限、失败页留白、取消、错误提示。
     交付：边界用例通过（超大/损坏/空页/单页）。
- [ ] **M4 部署与优化**（1 天）：Actions→Pages；代码分割降首屏；产物 < 1 MB gzip。
     交付：线上可访问 URL；首屏懒加载生效。
- [ ] **M5 文档**（0.5 天）：README、隐私说明、使用示例。

合计约 **1.5 周**（不含 M2 反复调优）。

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

1. `pptx-preview` 保真度上限（M2 前置信度：中）；SmartArt/艺术字/公式为主要风险。
2. 文本重排：`reconcile` 代价过高；备选方案（确保字体可用 / 换渲染库）待评测。
3. 100 MB 内存峰值需实测；可能需分批或降档。
4. 应用内无法视觉自验（背景不被预览器渲染）→ UX 需明确说明。
5. 单一渲染依赖 `pptx-preview` 的维护风险 → 锁版本 + 保留备选（`ChristopherVR/pptx-viewer`）。

---

## 附：PoC 运行

```bash
cd poc && npm install && npm run dev   # http://127.0.0.1:5173
```
