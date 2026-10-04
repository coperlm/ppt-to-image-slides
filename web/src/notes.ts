import JSZip from 'jszip'
import { NS_A, NS_P, normalizePartPath, parseXml } from './guards'

const NOTES_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'

/**
 * 桌面版靠「原文件当模板」天然保留备注；这里 PptxGenJS 是从零建包，
 * 所以把输入备注的纯文本抽出来，回填到输出包自带的空 notesSlide 部件里。
 * 只保文本，不保格式。
 */
export async function extractNotesTexts(zip: JSZip, slidePaths: string[]): Promise<(string | null)[]> {
  const texts: (string | null)[] = []
  for (const slidePath of slidePaths) {
    const notesPath = await findNotesSlidePath(zip, slidePath)
    if (!notesPath) {
      texts.push(null)
      continue
    }
    const file = zip.file(notesPath)
    texts.push(file ? readBodyText(await file.async('string')) : null)
  }
  return texts
}

export async function injectNotes(blob: Blob, notesTexts: (string | null)[]): Promise<Blob> {
  const zip = await JSZip.loadAsync(blob)
  const slidePaths = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => Number(a.match(/(\d+)\.xml$/)?.[1] ?? 0) - Number(b.match(/(\d+)\.xml$/)?.[1] ?? 0))

  for (const [offset, slidePath] of slidePaths.entries()) {
    const text = notesTexts[offset]
    if (!text) continue
    const notesPath = await findNotesSlidePath(zip, slidePath)
    if (!notesPath || !zip.file(notesPath)) continue
    zip.file(notesPath, notesSlideXml(text))
  }

  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' })
}

async function findNotesSlidePath(zip: JSZip, slidePath: string): Promise<string | null> {
  const dir = slidePath.slice(0, slidePath.lastIndexOf('/') + 1)
  const relsFile = zip.file(`${dir}_rels/${slidePath.slice(dir.length)}.rels`)
  if (!relsFile) return null
  for (const rel of Array.from(parseXml(await relsFile.async('string')).getElementsByTagName('Relationship'))) {
    if (rel.getAttribute('Type') !== NOTES_REL_TYPE) continue
    const target = rel.getAttribute('Target')
    return target ? normalizePartPath(dir, target) : null
  }
  return null
}

function readBodyText(notesXml: string): string | null {
  const doc = parseXml(notesXml)
  for (const shape of Array.from(doc.getElementsByTagNameNS(NS_P, 'sp'))) {
    const placeholder = shape.getElementsByTagNameNS(NS_P, 'ph')[0]
    if (placeholder?.getAttribute('type') !== 'body') continue
    const paragraphs = Array.from(shape.getElementsByTagNameNS(NS_A, 'p')).map((paragraph) =>
      Array.from(paragraph.getElementsByTagNameNS(NS_A, 't'))
        .map((run) => run.textContent ?? '')
        .join(''),
    )
    const text = paragraphs.join('\n').trim()
    return text || null
  }
  return null
}

function notesSlideXml(text: string): string {
  const paragraphs = text
    .split('\n')
    .map(
      (line) =>
        `<a:p><a:r><a:rPr lang="zh-CN" dirty="0"/><a:t>${escapeXml(line)}</a:t></a:r></a:p>`,
    )
    .join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree>' +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
    '<p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
    '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>' +
    `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>` +
    '</p:spTree></p:cSld></p:notes>'
  )
}

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}
