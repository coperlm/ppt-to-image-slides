import JSZip from 'jszip'
import { NS_A, NS_P, NS_R, normalizePartPath, parseXml, slideNumberOf } from './guards'

export interface SlideCheck {
  index: number
  path: string
  mediaPath: string | null
  ok: boolean
}

export interface VerifyResult {
  expected: number
  slideCount: number
  withBackground: number
  mediaCount: number
  slides: SlideCheck[]
  problems: string[]
  ok: boolean
}

/**
 * 判据不是「XML 里出现了某个字符串」，而是：每页的 p:bg/blipFill 的 r:embed
 * 必须经 rels 解析到包内真实存在的 JPEG/PNG 部件，且页数与输入一致。
 */
export async function verifyOutputPptx(blob: Blob, expectedSlides: number): Promise<VerifyResult> {
  const zip = await JSZip.loadAsync(blob)
  const slidePaths = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => slideNumberOf(a) - slideNumberOf(b))
  const mediaCount = Object.keys(zip.files).filter((name) => /^ppt\/media\//.test(name)).length

  const slides: SlideCheck[] = []
  const problems: string[] = []

  for (const [offset, path] of slidePaths.entries()) {
    const index = offset + 1
    const file = zip.file(path)
    if (!file) {
      problems.push(`第 ${index} 页：${path} 无法读取`)
      slides.push({ index, path, mediaPath: null, ok: false })
      continue
    }

    const embedId = findBackgroundEmbedId(await file.async('string'))
    if (!embedId) {
      problems.push(`第 ${index} 页：未找到 <p:bg> + <a:blipFill> 背景`)
      slides.push({ index, path, mediaPath: null, ok: false })
      continue
    }

    const mediaPath = await resolveRelationship(zip, path, embedId)
    const exists = mediaPath !== null && zip.file(mediaPath) !== null
    if (!exists) {
      problems.push(`第 ${index} 页：背景关系 ${embedId} 未解析到包内图片`)
    } else if (!/\.(jpe?g|png)$/i.test(mediaPath!)) {
      problems.push(`第 ${index} 页：背景图 ${mediaPath} 不是 JPEG/PNG（PowerPoint 兼容性差）`)
    }
    slides.push({ index, path, mediaPath, ok: exists })
  }

  if (slidePaths.length !== expectedSlides) {
    problems.push(`页数不符：期望 ${expectedSlides} 页，实际输出 ${slidePaths.length} 页`)
  }

  const withBackground = slides.filter((slide) => slide.ok).length
  if (mediaCount < withBackground) {
    problems.push(`media 图片数 ${mediaCount} 少于带背景的页数 ${withBackground}`)
  }

  return {
    expected: expectedSlides,
    slideCount: slidePaths.length,
    withBackground,
    mediaCount,
    slides,
    problems,
    ok: problems.length === 0,
  }
}

function findBackgroundEmbedId(slideXml: string): string | null {
  const background = parseXml(slideXml).getElementsByTagNameNS(NS_P, 'bg')[0]
  if (!background) return null
  const blip = background.getElementsByTagNameNS(NS_A, 'blip')[0]
  return blip?.getAttributeNS(NS_R, 'embed') ?? null
}

async function resolveRelationship(zip: JSZip, slidePath: string, relationshipId: string): Promise<string | null> {
  const dir = slidePath.slice(0, slidePath.lastIndexOf('/') + 1)
  const relsFile = zip.file(`${dir}_rels/${slidePath.slice(dir.length)}.rels`)
  if (!relsFile) return null

  for (const rel of Array.from(parseXml(await relsFile.async('string')).getElementsByTagName('Relationship'))) {
    if (rel.getAttribute('Id') !== relationshipId) continue
    const target = rel.getAttribute('Target')
    return target ? normalizePartPath(dir, target) : null
  }
  return null
}
