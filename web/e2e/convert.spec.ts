import { expect, test } from '@playwright/test'
import fs from 'node:fs'
import JSZip from 'jszip'

const FIXTURE = new URL('../../PPT_test.pptx', import.meta.url).pathname

// 界面默认语言跟随浏览器；测试固定为中文，语言行为单独用例覆盖
test.use({ locale: 'zh-CN' })

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
  await expect(page.locator('#status')).toContainText('另存为 .pptx')
  await expect(page.getByRole('button', { name: '开始转换' })).toBeDisabled()
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
  // 转换前切换语言时，日志里的就绪行也要跟着翻译（日志是纯文本追加，需专门重刷）
  await expect(page.locator('#log')).toContainText('Ready. Pick a .pptx')
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

test('键盘焦点环可见', async ({ page }) => {
  await page.goto('/')
  await page.keyboard.press('Tab')
  const outline = await page.evaluate(() => getComputedStyle(document.activeElement as Element).outlineWidth)
  expect(outline).toBe('2px')
})

test('选择文件后显示转换预估', async ({ page }) => {
  await page.goto('/')
  await page.setInputFiles('#file', FIXTURE)
  await expect(page.locator('#estimate')).toBeVisible()
  await expect(page.locator('#estimate')).toContainText('11 页')
})

test('离线可用：service worker 接管后断网仍能打开', async ({ page, context }) => {
  await page.goto('/')
  await page.evaluate(() => navigator.serviceWorker.ready)
  await page.reload()
  await context.setOffline(true)
  await page.reload()
  await expect(page.locator('#convert')).toBeVisible()
  await context.setOffline(false)
})

test('首次访问按浏览器语言选择德语', async ({ browser }) => {
  const context = await browser.newContext({ locale: 'de-DE' })
  const page = await context.newPage()
  await page.goto('/')
  await expect(page.locator('#convert')).toHaveText('Konvertieren')
  await expect(page.locator('#lang')).toHaveValue('de')
  await expect(page.locator('#log')).toContainText('Bereit.')
  await context.close()
})
