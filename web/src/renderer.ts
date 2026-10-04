import { RENDER_WIDTH } from './config'
import type { SlideSize } from './guards'

export interface SlideRenderer {
  slideCount: number
  /** 渲染指定页并返回其元素；保证 DOM 中同时只驻留一页，避免整副 deck 常驻内存 */
  renderSlide(index: number): Promise<HTMLElement>
  destroy(): void
}

/**
 * 用 mode:'slide' 逐页渲染。列表模式会把全部页面的 DOM（含所有内嵌图片的 base64）
 * 一次性常驻，百页级输入会直接压垮标签页；slide 模式下 renderSingleSlide 会先移除
 * 上一页，内存峰值与页数无关。翻页按钮和分页器挂在 slide wrapper 的同级，
 * 只截 wrapper 就不会把它们烤进背景图。
 */
export async function createRenderer(buffer: ArrayBuffer, sldSz: SlideSize): Promise<SlideRenderer> {
  const { init } = await import('pptx-preview')
  const height = Math.round((RENDER_WIDTH * sldSz.cy) / sldSz.cx)

  const host = document.createElement('div')
  host.className = 'render-host'
  document.body.appendChild(host)

  const previewer = init(host, { width: RENDER_WIDTH, height, mode: 'slide' })
  try {
    await previewer.preview(buffer)
    await waitForRender(host)
  } catch (error) {
    host.remove()
    throw error
  }

  return {
    slideCount: previewer.slideCount,
    async renderSlide(index) {
      if (index > 0) {
        previewer.renderSingleSlide(index)
        await waitForRender(host)
      }
      const el = host.querySelector<HTMLElement>('.pptx-preview-slide-wrapper')
      if (!el) throw new Error(`第 ${index + 1} 页渲染后未找到页面元素`)
      return el
    },
    destroy() {
      try {
        previewer.destroy()
      } catch {
        // destroy 只是尽力回收，失败不影响已经拿到的位图
      }
      host.remove()
    },
  }
}

/** 渲染就绪 = 图片全部加载完 + DOM 停止变化（echarts 图表是异步绘制的，固定 sleep 会截到半成品） */
async function waitForRender(root: HTMLElement): Promise<void> {
  await Promise.all([waitForImages(root), waitForQuiet(root)])
}

async function waitForImages(root: HTMLElement): Promise<void> {
  const pending = Array.from(root.querySelectorAll('img')).map(
    (img) =>
      new Promise<void>((resolve) => {
        if (img.complete) return resolve()
        img.addEventListener('load', () => resolve(), { once: true })
        img.addEventListener('error', () => resolve(), { once: true })
      }),
  )
  await withTimeout(Promise.all(pending), 30000)
}

async function waitForQuiet(root: HTMLElement, quietMs = 400): Promise<void> {
  await new Promise<void>((resolve) => {
    let timer = 0
    let observer: MutationObserver | undefined
    const finish = () => {
      window.clearTimeout(timer)
      observer?.disconnect()
      resolve()
    }
    observer = new MutationObserver(() => {
      window.clearTimeout(timer)
      timer = window.setTimeout(finish, quietMs)
    })
    observer.observe(root, { childList: true, subtree: true, attributes: true, characterData: true })
    timer = window.setTimeout(finish, quietMs)
    window.setTimeout(finish, 30000)
  })
}

async function withTimeout(promise: Promise<unknown>, ms: number): Promise<void> {
  let timer = 0
  const timeout = new Promise<void>((resolve) => {
    timer = window.setTimeout(resolve, ms)
  })
  try {
    await Promise.race([promise, timeout])
  } finally {
    window.clearTimeout(timer)
  }
}
