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

  for (const data of images) {
    const slide = pptx.addSlide()
    slide.background = { data }
  }

  return (await pptx.write({ outputType: 'blob' })) as Blob
}
