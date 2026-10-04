import { expect, test } from '@playwright/test'
import fs from 'node:fs'
import JSZip from 'jszip'

const FIXTURE = new URL('../../PPT_test.pptx', import.meta.url).pathname

test('整条流水线：转换 PPT_test.pptx 并通过结构自检', async ({ page }) => {
  await page.goto('/')
  await page.setInputFiles('#file', FIXTURE)
  await page.getByRole('button', { name: '开始转换' }).click()

  await expect(page.locator('#verify')).toContainText('PASS', { timeout: 120_000 })
  await expect(page.locator('.thumb')).toHaveCount(11)
  await expect(page.locator('.thumb .cap').first()).toContainText('2560×1440')
  await expect(page.locator('#download')).toBeVisible()
  await expect(page.locator('#verify')).toContainText('背景图解析成功的页:    11')
})

test('拒绝非 pptx 文件并给出提示', async ({ page }) => {
  await page.goto('/')
  await page.setInputFiles('#file', { name: 'notes.txt', mimeType: 'text/plain', buffer: Buffer.from('hi') })
  await expect(page.locator('#status')).toContainText('不支持的文件类型')
  await expect(page.getByRole('button', { name: '开始转换' })).toBeDisabled()
})

test('拒绝 97-2003 的 .ppt / 加密文件并提示另存为', async ({ page }) => {
  await page.goto('/')
  const cfb = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
  await page.setInputFiles('#file', { name: 'old.pptx', mimeType: 'application/octet-stream', buffer: cfb })
  await page.getByRole('button', { name: '开始转换' }).click()
  await expect(page.locator('#status')).toContainText('另存为 .pptx')
})

test('主题可切换、面板配色随主题变化且选择被记住', async ({ page }) => {
  await page.goto('/')
  await page.selectOption('#theme', 'dark')
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark')
  const darkPanel = await page.locator('#verify').evaluate((el) => getComputedStyle(el).backgroundColor)

  await page.selectOption('#theme', 'light')
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'light')
  const lightPanel = await page.locator('#verify').evaluate((el) => getComputedStyle(el).backgroundColor)
  expect(darkPanel).not.toBe(lightPanel)

  await page.reload()
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'light')

  await page.selectOption('#theme', 'system')
  await expect(page.locator('html')).not.toHaveAttribute('data-theme')
})

test('切换语言后界面文案变为英文', async ({ page }) => {
  await page.goto('/')
  await page.selectOption('#lang', 'en')
  await expect(page.locator('#convert')).toHaveText('Convert')
  await expect(page.locator('h1')).toContainText('image-background')
  await expect(page.locator('#verify')).toHaveText('(not run yet)')
})

test('演讲者备注以纯文本保留到输出包', async ({ page }) => {
  const zip = await JSZip.loadAsync(fs.readFileSync(FIXTURE))
  const notesPath = 'ppt/notesSlides/notesSlide1.xml'
  const anchor = '<p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>'
  const xml = await zip.file(notesPath)!.async('string')
  expect(xml).toContain(anchor)
  zip.file(notesPath, xml.replace(anchor, `${anchor}<a:p><a:r><a:t>NOTE-MARKER-42</a:t></a:r></a:p>`))
  const buffer = Buffer.from(await zip.generateAsync({ type: 'uint8array' }))

  await page.goto('/')
  await page.setInputFiles('#file', {
    name: 'with-notes.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer,
  })
  await page.getByRole('button', { name: '开始转换' }).click()
  await expect(page.locator('#download')).toBeVisible({ timeout: 120_000 })

  const [download] = await Promise.all([page.waitForEvent('download'), page.locator('#download').click()])
  const outZip = await JSZip.loadAsync(fs.readFileSync((await download.path())!))
  const outNotesNames = Object.keys(outZip.files).filter((n) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(n))
  const outNotes = await Promise.all(outNotesNames.map((n) => outZip.file(n)!.async('string')))
  expect(outNotes.some((xml) => xml.includes('NOTE-MARKER-42'))).toBe(true)
})
