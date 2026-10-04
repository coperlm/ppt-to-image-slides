import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { LIMITS } from './config'
import { guardAndRead, normalizePartPath, slideNumberOf } from './guards'

const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'

const presentation =
  `<?xml version="1.0"?><p:presentation xmlns:p="${NS_P}" xmlns:r="${NS_R}">` +
  '<p:sldSz cx="12192000" cy="6858000"/>' +
  '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>'
const presentationRels =
  `<?xml version="1.0"?><Relationships xmlns="${PKG_REL}">` +
  `<Relationship Id="rId2" Type="${NS_R}/slide" Target="slides/slide1.xml"/></Relationships>`
const slide = `<?xml version="1.0"?><p:sld xmlns:p="${NS_P}"/>`

async function makePptx(name = 'deck.pptx'): Promise<File> {
  const zip = new JSZip()
  zip.file('ppt/presentation.xml', presentation)
  zip.file('ppt/_rels/presentation.xml.rels', presentationRels)
  zip.file('ppt/slides/slide1.xml', slide)
  return new File([await zip.generateAsync({ type: 'blob' })], name)
}

describe('normalizePartPath', () => {
  it('解析相对 Target', () => {
    expect(normalizePartPath('ppt/', 'slides/slide1.xml')).toBe('ppt/slides/slide1.xml')
  })
  it('解析 ../ 上跳', () => {
    expect(normalizePartPath('ppt/slides/', '../media/image1.png')).toBe('ppt/media/image1.png')
  })
  it('解析包内绝对 Target', () => {
    expect(normalizePartPath('ppt/', '/ppt/slides/slide2.xml')).toBe('ppt/slides/slide2.xml')
  })
})

describe('slideNumberOf', () => {
  it('取文件名里的序号', () => {
    expect(slideNumberOf('ppt/slides/slide12.xml')).toBe(12)
  })
})

describe('guardAndRead', () => {
  it('以 sldIdLst 为准解析页数与页面尺寸', async () => {
    const guarded = await guardAndRead(await makePptx())
    expect(guarded.slideCount).toBe(1)
    expect(guarded.slideCountSource).toBe('sldIdLst')
    expect(guarded.sldSz).toEqual({ cx: 12192000, cy: 6858000, found: true })
    expect(guarded.slidePaths).toEqual(['ppt/slides/slide1.xml'])
  })

  it('拒绝非 .pptx 扩展名', async () => {
    const file = new File([new Uint8Array([0x50, 0x4b, 3, 4])], 'deck.txt')
    await expect(guardAndRead(file)).rejects.toThrow('不支持的文件类型')
  })

  it('拒绝加密 / 97-2003 的 CFB 容器并提示另存为', async () => {
    const file = new File([new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0, 0, 0, 0])], 'old.pptx')
    await expect(guardAndRead(file)).rejects.toThrow('另存为 .pptx')
  })

  it('拒绝超过输入上限的文件', async () => {
    const file = await makePptx()
    await expect(guardAndRead(file, { ...LIMITS, maxInputBytes: 8 })).rejects.toThrow('超过上限')
  })

  it('拒绝没有幻灯片的包', async () => {
    const zip = new JSZip()
    zip.file('ppt/presentation.xml', presentation.replace('<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>', ''))
    zip.file('ppt/_rels/presentation.xml.rels', presentationRels)
    const file = new File([await zip.generateAsync({ type: 'blob' })], 'empty.pptx')
    await expect(guardAndRead(file)).rejects.toThrow('没有幻灯片')
  })
})
