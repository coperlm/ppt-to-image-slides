import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { NS_A, NS_P, NS_R } from './guards'
import { verifyOutputPptx } from './verifier'

const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const JPEG = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00])

const slideXml = (rid: string) =>
  `<?xml version="1.0"?><p:sld xmlns:p="${NS_P}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">` +
  `<p:bg><p:bgPr><a:blipFill><a:blip r:embed="${rid}"/></a:blipFill></p:bgPr></p:bg></p:sld>`

const slideRels = (target: string) =>
  `<?xml version="1.0"?><Relationships xmlns="${PKG_REL}">` +
  `<Relationship Id="rId1" Type="${NS_R}/image" Target="${target}"/></Relationships>`

async function makeOutput(mediaName: string, mediaBytes: Uint8Array, withRels = true): Promise<Blob> {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', slideXml('rId1'))
  if (withRels) zip.file('ppt/slides/_rels/slide1.xml.rels', slideRels(`../media/${mediaName}`))
  zip.file(`ppt/media/${mediaName}`, mediaBytes)
  return zip.generateAsync({ type: 'blob' })
}

describe('verifyOutputPptx', () => {
  it('扩展名与字节一致时通过', async () => {
    const result = await verifyOutputPptx(await makeOutput('image1.jpeg', JPEG), 1)
    expect(result.ok).toBe(true)
    expect(result.withBackground).toBe(1)
    expect(result.mediaCount).toBe(1)
  })

  it('JPEG 字节挂在 .png 名下时判 FAIL', async () => {
    const result = await verifyOutputPptx(await makeOutput('image1.png', JPEG), 1)
    expect(result.ok).toBe(false)
    expect(result.problems.join('\n')).toContain('声明为 png')
  })

  it('关系解析不到包内图片时判 FAIL', async () => {
    const result = await verifyOutputPptx(await makeOutput('image1.jpeg', JPEG, false), 1)
    expect(result.ok).toBe(false)
    expect(result.problems.join('\n')).toContain('未解析到包内图片')
  })

  it('页数不符时判 FAIL', async () => {
    const result = await verifyOutputPptx(await makeOutput('image1.jpeg', JPEG), 2)
    expect(result.ok).toBe(false)
    expect(result.problems.join('\n')).toContain('页数不符')
  })
})
