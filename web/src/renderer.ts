import { RENDER_WIDTH } from './config'
import type { SlideSize } from './guards'

export interface Rendered {
  slideEls: HTMLElement[]
  destroy(): void
}

export async function renderPptx(buffer: ArrayBuffer, sldSz: SlideSize): Promise<Rendered> {
  const { init } = await import('pptx-preview')
  const height = Math.round((RENDER_WIDTH * sldSz.cy) / sldSz.cx)

  const host = document.createElement('div')
  host.className = 'render-host'
  document.body.appendChild(host)

  const previewer = init(host, { width: RENDER_WIDTH, height, mode: 'list' })
  try {
    await previewer.preview(buffer)
    await waitForRender(host)
  } catch (error) {
    host.remove()
    throw error
  }

  const wrappers = Array.from(host.querySelectorAll<HTMLElement>('.pptx-preview-slide-wrapper'))
  const slideEls = wrappers.length ? wrappers : (Array.from(host.children) as HTMLElement[])

  return {
    slideEls,
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

async function withTimeout<T>(promise: Promise<T>, ms: number): Promise<T | void> {
  let timer = 0
  const timeout = new Promise<void>((resolve) => {
    timer = window.setTimeout(resolve, ms)
  })
  try {
    return await Promise.race([promise, timeout])
  } finally {
    window.clearTimeout(timer)
  }
}
