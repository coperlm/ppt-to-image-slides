import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { NS_A, NS_P, NS_R } from './guards'
import { extractNotesTexts, injectNotes } from './notes'

const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const NOTES_REL = `${NS_R}/notesSlide`

const notesSlideXml = (text: string) =>
  `<?xml version="1.0"?><p:notes xmlns:p="${NS_P}" xmlns:a="${NS_A}"><p:cSld><p:spTree>` +
  '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes"/><p:cNvSpPr/>' +
  `<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:lstStyle/>` +
  `<a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`

async function makeInputZip(): Promise<JSZip> {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0"?><p:sld xmlns:p="${NS_P}"/>`)
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="${PKG_REL}">` +
      `<Relationship Id="rId9" Type="${NOTES_REL}" Target="../notesSlides/notesSlide1.xml"/></Relationships>`,
  )
  zip.file('ppt/notesSlides/notesSlide1.xml', notesSlideXml('hello notes'))
  return zip
}

async function makeOutputZip(): Promise<Blob> {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0"?><p:sld xmlns:p="${NS_P}"/>`)
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="${PKG_REL}">` +
      `<Relationship Id="rId9" Type="${NOTES_REL}" Target="../notesSlides/notesSlide1.xml"/></Relationships>`,
  )
  zip.file('ppt/notesSlides/notesSlide1.xml', notesSlideXml(''))
  return zip.generateAsync({ type: 'blob' })
}

describe('notes', () => {
  it('从输入包抽出备注正文', async () => {
    const texts = await extractNotesTexts(await makeInputZip(), ['ppt/slides/slide1.xml'])
    expect(texts).toEqual(['hello notes'])
  })

  it('没有备注的页返回 null', async () => {
    const zip = new JSZip()
    zip.file('ppt/slides/slide1.xml', `<?xml version="1.0"?><p:sld xmlns:p="${NS_P}"/>`)
    const texts = await extractNotesTexts(zip, ['ppt/slides/slide1.xml'])
    expect(texts).toEqual([null])
  })

  it('把备注文本写回输出包的 notesSlide', async () => {
    const injected = await injectNotes(await makeOutputZip(), ['hello notes'])
    const zip = await JSZip.loadAsync(injected)
    const out = await zip.file('ppt/notesSlides/notesSlide1.xml')!.async('string')
    expect(out).toContain('<a:t>hello notes</a:t>')
  })
})
