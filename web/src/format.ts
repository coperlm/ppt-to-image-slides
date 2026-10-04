export function fmtBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / 1024 / 1024).toFixed(2)} MB`
}

export function fmtSeconds(ms: number): string {
  return `${(ms / 1000).toFixed(1)} s`
}

/** 让出主线程；不用 requestAnimationFrame —— 标签页隐藏时 rAF 不触发，长转换会被永久挂起 */
export function yieldToBrowser(): Promise<void> {
  return new Promise((resolve) => {
    const channel = new MessageChannel()
    channel.port1.onmessage = () => resolve()
    channel.port2.postMessage(null)
  })
}

export function throwIfAborted(signal?: AbortSignal): void {
  if (signal?.aborted) throw new DOMException('已取消转换', 'AbortError')
}

export function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
