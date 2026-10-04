import type { SlideSize } from './guards'

const EMU_PER_INCH = 914400

export async function packBackgroundPptx(images: string[], sldSz: SlideSize): Promise<Blob> {
  const { default: PptxGenJS } = await import('pptxgenjs')
  const pptx = new PptxGenJS()
  pptx.defineLayout({
    name: 'SOURCE',
    width: sldSz.cx / EMU_PER_INCH,
    height: sldSz.cy / EMU_PER_INCH,
  })
  pptx.layout = 'SOURCE'

  for (const [offset, data] of images.entries()) {
    const slide = pptx.addSlide()
    // 只给 data 时 PptxGenJS 把扩展名默认成 png，而字节其实是 JPEG —— 声明类型与内容不符会触发
    // PowerPoint 启动时的内容警告。path 仅用于推导扩展名，有 data 时不会去 fetch 它。
    slide.background = { data, path: `slide-${offset + 1}.jpeg` }
  }

  return (await pptx.write({ outputType: 'blob' })) as Blob
}
