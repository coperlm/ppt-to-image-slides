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
 * 判据不是「XML 里出现了某个字符串」，而是：每页 p:bg 里 blip 的 r:embed 必须经该页 rels
 * 解析到包内真实存在的部件，且部件的 magic number 与其扩展名一致，页数与输入一致。
 */
export async function verifyOutputPptx(blob: Blob, expectedSlides: number): Promise<VerifyResult> {
  const zip = await JSZip.loadAsync(blob)
  const slidePaths = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => slideNumberOf(a) - slideNumberOf(b))
  const mediaCount = Object.keys(zip.files).filter((name) => /^ppt\/media\//.test(name) && !zip.files[name].dir).length

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
    const mediaFile = mediaPath ? zip.file(mediaPath) : null
    if (!mediaPath || !mediaFile) {
      problems.push(`第 ${index} 页：背景关系 ${embedId} 未解析到包内图片`)
      slides.push({ index, path, mediaPath, ok: false })
      continue
    }

    // 扩展名与真实字节必须一致：PptxGenJS 只给 data 时会把 JPEG 字节写成 .png，
    // PowerPoint 打开时会弹「发现内容有问题」的修复提示
    const actual = sniffImageType(await mediaFile.async('uint8array'))
    const declared = declaredImageType(mediaPath)
    const mismatch = actual === 'unknown' || actual !== declared
    if (mismatch) {
      problems.push(
        actual === 'unknown'
          ? `第 ${index} 页：背景图 ${mediaPath} 不是可识别的 JPEG/PNG`
          : `第 ${index} 页：背景图 ${mediaPath} 声明为 ${declared}，实际字节是 ${actual}`,
      )
    }
    slides.push({ index, path, mediaPath, ok: !mismatch })
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

type ImageType = 'jpeg' | 'png' | 'unknown'

function sniffImageType(bytes: Uint8Array): ImageType {
  if (bytes[0] === 0xff && bytes[1] === 0xd8) return 'jpeg'
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return 'png'
  return 'unknown'
}

function declaredImageType(mediaPath: string): ImageType {
  const ext = mediaPath.split('.').pop()?.toLowerCase()
  if (ext === 'png') return 'png'
  if (ext === 'jpg' || ext === 'jpeg') return 'jpeg'
  return 'unknown'
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
