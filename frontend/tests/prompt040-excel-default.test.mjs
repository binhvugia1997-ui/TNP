/**
 * PROMPT-040 — the remembered Excel selections are shown, reused and replaced explicitly.
 *
 * jsdom cannot prove what the operator actually sees, so these tests boot the REAL application
 * (`App` + `StoreProvider`, real `SourcePanel`) through Vite in a real Chromium tab with only the
 * pywebview bridge mocked (`tests/layout/excel-default-harness.tsx`).
 *
 * Covered here:
 *   • a restored default is displayed (file name + folder + state chip) and NO picker runs on startup,
 *   • a stored file that disappeared is reported (“Không tìm thấy” + message) and never auto-replaced,
 *   • “Đổi file…” opens exactly one native picker and the newly persisted path is displayed again,
 *   • “Xóa lựa chọn” calls the backend once and the row returns to “Chưa chọn”,
 *   • the two Excel rows are independent (clearing one keeps the other).
 *
 * Chromium comes from the checked-in dev dependency `@sparticuz/chromium` (offline, no download). When no
 * usable browser is available the tests SKIP with a printed reason instead of failing — they must never be
 * reported as passing.
 */
import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { createServer } from 'vite'
import { fileURLToPath } from 'node:url'
import { execFileSync } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { brotliDecompressSync } from 'node:zlib'

const frontendRoot = fileURLToPath(new URL('..', import.meta.url))
const HARNESS = '/tests/layout/excel-default-harness.html'
const TEMPLATE_LABEL = 'File kiểm chứng Excel'
const OUTPUT_LABEL = 'File kết quả Excel'
const PICK_SELECTOR = (label) => `button[aria-label="Chọn ${label}"]`
const CLEAR_SELECTOR = (label) => `button[aria-label="Xóa lựa chọn ${label}"]`

let server
let browser
let baseUrl = ''
let unavailableReason = ''
let libDir = process.env.TNP_CHROMIUM_LIB_DIR ?? ''

const delay = (ms) => new Promise((resolve) => globalThis.setTimeout(resolve, ms))

/** Linux/macOS Chromium needs NSS/NSPR; the AL2023 bundle shipped with @sparticuz/chromium carries them. */
function extractChromiumLibraries() {
  const archive = `${frontendRoot}node_modules/@sparticuz/chromium/bin/al2023.tar.br`
  if (!existsSync(archive)) return ''
  const target = join(tmpdir(), 'tnp-chromium-al2023')
  try {
    mkdirSync(target, { recursive: true })
    if (!existsSync(join(target, 'lib', 'libnspr4.so'))) {
      const tarPath = join(tmpdir(), 'tnp-chromium-al2023.tar')
      writeFileSync(tarPath, brotliDecompressSync(readFileSync(archive)))
      execFileSync('tar', ['-xf', tarPath, '-C', target])
    }
    return existsSync(join(target, 'lib', 'libnspr4.so')) ? join(target, 'lib') : ''
  } catch {
    return ''
  }
}

async function launchBrowser(puppeteer, executablePath) {
  const args = ['--no-sandbox', '--disable-dev-shm-usage', '--disable-gpu', '--font-render-hinting=none']
  const attempt = async (libraries) => {
    if (libraries) {
      process.env.LD_LIBRARY_PATH = [libraries, process.env.LD_LIBRARY_PATH || ''].filter(Boolean).join(':')
    }
    return puppeteer.launch({ args, executablePath, headless: true, defaultViewport: { width: 1440, height: 940 } })
  }
  try {
    return await attempt(libDir)
  } catch (error) {
    const message = String(error?.message ?? error)
    if (!/libnspr4|libnss3|shared librar/i.test(message)) throw error
    libDir = libDir || extractChromiumLibraries()
    if (!libDir) throw error
    return attempt(libDir)
  }
}

before(async () => {
  try {
    const puppeteer = (await import('puppeteer')).default
    const chromium = (await import('@sparticuz/chromium')).default
    const executablePath = await chromium.executablePath()
    browser = await launchBrowser(puppeteer, executablePath)
  } catch (error) {
    unavailableReason = String(error?.message ?? error).split('\n')[0]
    browser = undefined
  }

  server = await createServer({
    configFile: `${frontendRoot}vite.config.ts`,
    root: frontendRoot,
    appType: 'spa',
    server: { middlewareMode: false, host: '127.0.0.1', port: 5198, strictPort: false, hmr: false },
    logLevel: 'error',
  })
  await server.listen()
  baseUrl = server.resolvedUrls.local[0].replace(/\/$/, '')
})

after(async () => {
  await browser?.close().catch(() => undefined)
  await server?.close()
})

function browserUnavailable(t) {
  if (browser) return false
  t.skip(`real-browser UI measurement unavailable: ${unavailableReason || 'no chromium'}`)
  return true
}

/** Read the rendered row of one Excel kind through the real DOM (runs INSIDE the page). */
function readExcelRow(label) {
  const pick = document.querySelector(`button[aria-label="Chọn ${label}"]`)
  const clear = document.querySelector(`button[aria-label="Xóa lựa chọn ${label}"]`)
  const field = pick && pick.parentElement ? pick.parentElement.parentElement : null
  if (!field || !pick) return null
  const input = field.querySelector('input')
  return {
    text: field.textContent ?? '',
    inputValue: input ? input.value : null,
    pickButtonText: pick.textContent ?? '',
    clearDisabled: clear ? clear.disabled : null,
  }
}

async function openPage() {
  const page = await browser.newPage()
  await page.goto(`${baseUrl}${HARNESS}`, { waitUntil: 'load' })
  const row = (label) => page.evaluate(readExcelRow, label)
  const calls = () => page.evaluate(() => window.__tnpExcelHarness.bridgeCalls())
  const cleared = () => page.evaluate(() => window.__tnpExcelHarness.clearedKinds())
  /** The store polls every 750 ms; wait until the rendered row contains the expected text. */
  const waitForText = async (label, needle, timeoutMs = 6000) => {
    const deadline = Date.now() + timeoutMs
    let last = null
    while (Date.now() < deadline) {
      last = await row(label)
      if (last && last.text.includes(needle)) return last
      await delay(80)
    }
    assert.fail(`row "${label}" never showed ${JSON.stringify(needle)}; last=${JSON.stringify(last)}`)
  }
  const waitForCalls = async (method, minimum, timeoutMs = 4000) => {
    const deadline = Date.now() + timeoutMs
    let last = 0
    while (Date.now() < deadline) {
      last = (await calls())[method] ?? 0
      if (last >= minimum) return last
      await delay(60)
    }
    assert.fail(`${method} was called ${last} times, expected >= ${minimum}`)
  }
  return { page, row, calls, cleared, waitForText, waitForCalls }
}

test('a restored Excel default is shown and no picker runs on startup', async (t) => {
  if (browserUnavailable(t)) return
  const { page, row, calls, waitForText } = await openPage()
  const template = await waitForText(TEMPLATE_LABEL, 'Biểu mẫu kiểm chứng.xlsx')
  const output = await waitForText(OUTPUT_LABEL, 'Kiem chung 09-2026.xlsx')

  assert.match(template.text, /Excel mặc định:/)
  assert.match(template.text, /D:\\Kiểm chứng 2026/)
  assert.match(template.text, /Dùng được/)
  assert.match(output.text, /Dùng được/)
  assert.equal(template.inputValue, 'D:\\Kiểm chứng 2026\\Biểu mẫu kiểm chứng.xlsx')
  assert.equal(template.pickButtonText, 'Đổi file…')

  await delay(1200)                                       // several dashboard polls
  const native = await calls()
  assert.equal(native.choose_template_file ?? 0, 0, 'startup must not open the Excel picker')
  assert.equal(native.choose_output_file ?? 0, 0, 'startup must not open the Excel picker')
  await page.close()
})

test('a stored file that disappeared is reported and not replaced', async (t) => {
  if (browserUnavailable(t)) return
  const { page, row, calls, waitForText } = await openPage()
  await waitForText(TEMPLATE_LABEL, 'Biểu mẫu kiểm chứng.xlsx')

  await page.evaluate(() => window.__tnpExcelHarness.setExcel({
    template: { state: 'missing', message: 'File kiểm chứng đã lưu không còn tồn tại. Hãy chọn lại file.' },
  }))
  const rowState = await waitForText(TEMPLATE_LABEL, 'Không tìm thấy')
  assert.match(rowState.text, /không còn tồn tại/)
  const native = await calls()
  assert.equal(native.choose_template_file ?? 0, 0)
  await page.close()
})

test('“Đổi file…” persists a new pick and displays it again', async (t) => {
  if (browserUnavailable(t)) return
  const { page, row, calls, waitForCalls, waitForText } = await openPage()
  await waitForText(TEMPLATE_LABEL, 'Biểu mẫu kiểm chứng.xlsx')

  await page.evaluate(() => window.__tnpExcelHarness.queuePick('template', 'D:\\Kiểm chứng 2026\\Biểu mẫu (bản mới).xlsx'))
  await page.click(PICK_SELECTOR(TEMPLATE_LABEL))
  await waitForCalls('choose_template_file', 1)
  const updated = await waitForText(TEMPLATE_LABEL, 'Biểu mẫu (bản mới).xlsx')
  assert.match(updated.text, /Dùng được/)
  assert.equal(updated.inputValue, 'D:\\Kiểm chứng 2026\\Biểu mẫu (bản mới).xlsx')
  const native = await calls()
  assert.equal(native.choose_template_file, 1, 'exactly one native dialog per click')
  assert.ok((native.get_excel_selection ?? 0) >= 1, 'the persisted default is re-read from Python')
  await page.close()
})

test('“Xóa lựa chọn” clears only that Excel default', async (t) => {
  if (browserUnavailable(t)) return
  const { page, row, cleared, waitForText } = await openPage()
  await waitForText(TEMPLATE_LABEL, 'Biểu mẫu kiểm chứng.xlsx')

  await page.click(CLEAR_SELECTOR(TEMPLATE_LABEL))
  const clearedRow = await waitForText(TEMPLATE_LABEL, 'chưa chọn')
  assert.equal(clearedRow.inputValue, '')
  assert.deepEqual(await cleared(), ['template'])
  const output = await row(OUTPUT_LABEL)
  assert.match(output.text, /Kiem chung 09-2026.xlsx/, 'the other Excel default is untouched')
  await page.close()
})
