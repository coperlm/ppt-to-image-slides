// 用法（在 web/ 目录下）: node tools/fidelity/extract-output.mjs <converted.pptx> <out-dir>
// 把转换结果包里的背景图按页序解出来，供与参考渲染逐页比对。
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'

const [pptxPath, outDir] = process.argv.slice(2)
if (!pptxPath || !outDir) {
  console.error('usage: node tools/fidelity/extract-output.mjs <converted.pptx> <out-dir>')
  process.exit(1)
}

const zip = await JSZip.loadAsync(readFileSync(pptxPath))
const media = Object.keys(zip.files)
  .filter((name) => /^ppt\/media\/.+\.(jpe?g|png)$/i.test(name))
  .sort((a, b) => Number(a.match(/(\d+)/)?.[1] ?? 0) - Number(b.match(/(\d+)/)?.[1] ?? 0))

mkdirSync(outDir, { recursive: true })
for (const [offset, name] of media.entries()) {
  const bytes = await zip.file(name).async('uint8array')
  const ext = name.split('.').pop()?.toLowerCase() === 'png' ? 'png' : 'jpg'
  writeFileSync(join(outDir, `out-${String(offset + 1).padStart(2, '0')}.${ext}`), Buffer.from(bytes))
}
console.log(`已解出 ${media.length} 张背景图 → ${outDir}/out-*.jpg|png`)
