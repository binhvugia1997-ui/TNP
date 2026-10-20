/**
 * PROMPT-034R — the “Danh sách báo cáo” list is a bounded viewport, not a page-growing table.
 *
 * jsdom performs no layout, so every claim in this task (“~6 visible rows”, “the column header stays
 * sticky”, “1000 reports do not stretch the page”, “no horizontal scrollbar at desktop width”) has to be
 * measured in a real engine. These tests boot the REAL application (`App` + `StoreProvider`, real
 * `ReportTable`) through Vite in a real Chromium tab, with only the pywebview bridge mocked
 * (`tests/layout/report-list-harness.tsx`), and measure the rendered geometry.
 *
 * Covered here:
 *   • 0 rows            → no oversized empty report area (the old fixed 430px column placeholder must be gone)
 *   • 1…6 rows          → natural compact height, no internal scrollbar, no reserved blank space
 *   • 7 / 30 / 1000 / 1200 rows → bounded viewport (≈6 rows + header), scrolls INSIDE the list
 *   • sticky column header stays pinned and fully visible at the top, mid and bottom of the scroll
 *   • the whole page does not grow with the number of reports (page height is dataset-independent)
 *   • every report stays in the dataset: no pagination, no truncation, last row reachable by scrolling
 *   • search + status filter still work, and filtered results keep the bounded/natural-height behaviour
 *   • row selection and select-all still operate on the whole dataset
 *   • incremental scan updates still render into the bounded list
 *   • no horizontal scrollbar at 1440/1366/1280/1024/900 px, 8 columns preserved
 *   • a large fixture does not add per-row native bridge traffic (no fan-out to Python)
 *
 * Chromium comes from the checked-in dev dependency `@sparticuz/chromium` (offline, no download).
 * When no usable browser is available the geometry tests SKIP with a printed reason instead of failing —
 * they must never be reported as passing.
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
const DESKTOP = { width: 1440, height: 940 }
/** Window sizes the application is actually used at (the default pywebview window is 1440×940). */
const DESKTOP_WIDTHS = [1440, 1366, 1280, 1024, 900]
/** One measured row at 13px/19px body text: 34px result cell + 2×6px cell padding + 1px separator. */
const ROW_HEIGHT = 47
const ROW_TOLERANCE = 6

/**
 * The list viewport, located structurally so the SAME test measures the pre-fix markup as well:
 * `.report-list-viewport` is the PROMPT-034R cap, `div.overflow-auto` was the unbounded box before it.
 * A regression therefore fails on geometry, not on a missing class name.
 */
const VIEWPORT = '.report-list-viewport, div.overflow-auto'
const SECTION_TITLE = 'Danh sách báo cáo'
const SEARCH_INPUT = 'input[placeholder^="Tìm theo Management Number"]'
/** Same box, class-only form, for selector concatenation (a comma would split the selector list). */
const VIEWPORT_LIST = '.report-list-viewport'

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
    return puppeteer.launch({ args, executablePath, headless: true, defaultViewport: DESKTOP })
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

  // `appType: 'spa'` is required: the harness is a real HTML entry the Vite HTML middleware must transform.
  // A dedicated port keeps a developer's `npm run dev` on 5173 untouched (strictPort: false auto-increments).
  server = await createServer({
    configFile: `${frontendRoot}vite.config.ts`,
    root: frontendRoot,
    appType: 'spa',
    server: { middlewareMode: false, host: '127.0.0.1', port: 5199, strictPort: false, hmr: false },
    logLevel: 'error',
  })
  await server.listen()
  baseUrl = server.resolvedUrls.local[0].replace(/\/$/, '')
})

after(async () => {
  await browser?.close().catch(() => undefined)
  await server?.close()
})

/** True when the browser is unusable; the geometry tests then skip with the recorded reason. */
function browserUnavailable(t) {
  if (browser) return false
  t.skip(`real-browser layout measurement unavailable: ${unavailableReason || 'no chromium'}`)
  return true
}

/** Real measurement performed inside the page: list viewport, rows, sticky header and page boxes. */
function measureInPage(selector, title) {
  const section = Array.from(document.querySelectorAll('section.card'))
    .find((element) => element.querySelector('h2')?.textContent === title)
  if (!section) return null
  const viewport = section.querySelector(selector)
  if (!viewport) return null
  const table = section.querySelector('table')
  const head = table?.querySelector('thead') ?? null
  const bodyRows = Array.from(table?.querySelectorAll('tbody tr') ?? [])
  const box = viewport.getBoundingClientRect()
  const headBox = head?.getBoundingClientRect()
  const bodyRowBoxes = bodyRows.map((row) => row.getBoundingClientRect())
  const main = document.querySelector('main')
  const scrolling = document.scrollingElement
  const computed = getComputedStyle(viewport)
  const firstRow = bodyRows[0] ? bodyRows[0].getBoundingClientRect() : null
  return {
    viewport: {
      clientHeight: viewport.clientHeight,
      scrollHeight: viewport.scrollHeight,
      clientWidth: viewport.clientWidth,
      scrollWidth: viewport.scrollWidth,
      scrollTop: viewport.scrollTop,
      overflowY: computed.overflowY,
      maxHeight: computed.maxHeight,
      /** The documented cap: `--report-rows` rows + the sticky header + the box frame. */
      capPx: Number.parseFloat(computed.maxHeight),
      rowVar: Number.parseFloat(computed.getPropertyValue('--report-row-h')),
      headVar: Number.parseFloat(computed.getPropertyValue('--report-head-h')),
      rowsVar: Number.parseFloat(computed.getPropertyValue('--report-rows')),
      frameVar: Number.parseFloat(computed.getPropertyValue('--report-frame')),
      top: box.top,
      bottom: box.bottom,
    },
    head: head ? {
      top: headBox.top,
      height: headBox.height,
      bottomRow: headBox.bottom,
      position: getComputedStyle(head).position,
      background: getComputedStyle(head.querySelector('th')).backgroundColor,
    } : null,
    rowCount: bodyRows.length,
    rowHeight: firstRow ? firstRow.height : null,
    visibleRows: bodyRowBoxes.filter((row) => row.top >= box.top - 1 && row.bottom <= box.bottom + 1).length,
    firstRowStt: bodyRows[0]?.querySelector('td:nth-child(2)')?.textContent ?? null,
    lastRowStt: bodyRows.at(-1)?.querySelector('td:nth-child(2)')?.textContent ?? null,
    columnCount: table?.querySelectorAll('thead th').length ?? 0,
    colgroupCount: table?.querySelectorAll('colgroup col').length ?? 0,
    cellsPerFirstRow: bodyRows[0]?.querySelectorAll('td').length ?? 0,
    cardHeight: section.getBoundingClientRect().height,
    page: {
      documentScrollHeight: scrolling.scrollHeight,
      documentClientHeight: scrolling.clientHeight,
      documentScrollWidth: scrolling.scrollWidth,
      documentClientWidth: scrolling.clientWidth,
      mainScrollHeight: main.scrollHeight,
      mainClientHeight: main.clientHeight,
      mainScrollTop: main.scrollTop,
      mainScrollWidth: main.scrollWidth,
      mainClientWidth: main.clientWidth,
    },
    horizontalOverflow: {
      document: scrolling.scrollWidth - scrolling.clientWidth,
      main: main.scrollWidth - main.clientWidth,
      list: viewport.scrollWidth - viewport.clientWidth,
    },
  }
}

async function openList(rows, viewport = DESKTOP) {
  const page = await browser.newPage()
  await page.setViewport(viewport)
  const errors = []
  page.on('pageerror', (error) => errors.push(String(error)))
  await page.goto(`${baseUrl}/tests/layout/report-list-harness.html?rows=${rows}`, { waitUntil: 'load' })
  await page.waitForFunction((selector, title, expected) => {
    if (document.body.innerText.includes('Chưa kết nối ứng dụng Python')) return false
    const section = Array.from(document.querySelectorAll('section.card'))
      .find((element) => element.querySelector('h2')?.textContent === title)
    if (!section?.querySelector(selector)) return false
    const rendered = section.querySelectorAll('tbody tr').length
    return expected === 0 ? !section.querySelector('table') : rendered === expected
  }, { timeout: 30_000 }, VIEWPORT, SECTION_TITLE, rows)
  // One settle pass so both the 750 ms poll and layout have landed.
  await page.evaluate(() => new Promise((resolve) => {
    requestAnimationFrame(() => requestAnimationFrame(() => resolve()))
  }))
  await delay(120)
  assert.deepEqual(errors, [], 'the report list must render without runtime errors')
  return page
}

async function measure(page) {
  const value = await page.evaluate(measureInPage, VIEWPORT, SECTION_TITLE)
  assert.ok(value, 'the report list viewport must exist')
  return value
}

/** Set the React-controlled search box through the native value setter (typing alone cannot clear it). */
async function setSearch(page, value) {
  await page.$eval(SEARCH_INPUT, (element, next) => {
    const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set
    setter.call(element, next)
    element.dispatchEvent(new Event('input', { bubbles: true }))
  }, value)
}

/** Wait until the rendered report list body shows exactly `count` rows. */
async function waitForRows(page, count) {
  await page.waitForFunction((selector, title, expected) => {
    const section = Array.from(document.querySelectorAll('section.card'))
      .find((element) => element.querySelector('h2')?.textContent === title)
    if (!section?.querySelector(selector)) return false
    const rendered = section.querySelectorAll('tbody tr').length
    return expected === 0 ? !section.querySelector('table') : rendered === expected
  }, { timeout: 10_000 }, VIEWPORT, SECTION_TITLE, count)
}

function assertNoHorizontalScrollbar(measured, label) {
  assert.equal(measured.horizontalOverflow.list, 0, `${label}: the report list must not scroll horizontally`)
  assert.equal(measured.horizontalOverflow.main, 0, `${label}: the workspace column must not scroll horizontally`)
  assert.equal(measured.horizontalOverflow.document, 0, `${label}: the window must not scroll horizontally`)
}

test('PROMPT-034R · 0 rows: no oversized empty report area', async (t) => {
  if (browserUnavailable(t)) return
  const page = await openList(0)
  try {
    const measured = await measure(page)
    assert.equal(measured.rowCount, 0)
    assert.equal(measured.head, null, 'an empty list renders no table header')
    // Pre-fix this box was a fixed 300px minimum inside a fixed 430px column (rendered height 298px).
    assert.ok(measured.viewport.clientHeight < 260,
      `0 rows must not reserve a blank report area (viewport was ${measured.viewport.clientHeight}px)`)
    assert.ok(measured.cardHeight < 350,
      `the report card must shrink for an empty list (card was ${measured.cardHeight}px)`)
    assert.equal(measured.viewport.scrollHeight - measured.viewport.clientHeight, 0,
      'an empty list must not scroll')
    assertNoHorizontalScrollbar(measured, '0 rows')
  } finally {
    await page.close()
  }
})

test('PROMPT-034R · 1–6 rows: natural compact height, no reserved blank space', async (t) => {
  if (browserUnavailable(t)) return
  const heights = {}
  for (const rows of [1, 2, 3, 6]) {
    const page = await openList(rows)
    try {
      const measured = await measure(page)
      assert.equal(measured.rowCount, rows)
      assert.ok(measured.rowHeight && Math.abs(measured.rowHeight - ROW_HEIGHT) <= ROW_TOLERANCE,
        `unexpected row height for ${rows} rows: ${measured.rowHeight}`)
      const expected = measured.head.height + rows * measured.rowHeight + 2 // + the 2px box border
      assert.ok(Math.abs(measured.viewport.clientHeight - expected) <= 2,
        `${rows} rows must keep natural height (got ${measured.viewport.clientHeight}px, expected ${expected}px)`)
      assert.ok(measured.viewport.clientHeight < measured.viewport.capPx + 1,
        `${rows} rows must stay under the 6-row cap`)
      assert.equal(measured.viewport.scrollHeight - measured.viewport.clientHeight, 0,
        `${rows} rows must not create an internal scrollbar`)
      assert.equal(measured.viewport.overflowY, 'auto', 'the list box stays the scroll container')
      assert.equal(measured.visibleRows, rows, 'every row is visible')
      assertNoHorizontalScrollbar(measured, `${rows} rows`)
      heights[rows] = measured.viewport.clientHeight
    } finally {
      await page.close()
    }
  }
  assert.ok(heights[1] < 150, `1 row must stay compact — no fixed ~300px box (was ${heights[1]}px)`)
  assert.ok(heights[1] < heights[2] && heights[2] < heights[3] && heights[3] < heights[6],
    `height must grow with the row count (${JSON.stringify(heights)})`)
})

test('PROMPT-034R · >6 rows: bounded viewport, internal vertical scrolling, sticky header', async (t) => {
  if (browserUnavailable(t)) return
  const baselines = {}
  for (const rows of [7, 30, 1000]) {
    const page = await openList(rows)
    try {
      const measured = await measure(page)
      assert.ok(Math.abs(measured.rowHeight - ROW_HEIGHT) <= ROW_TOLERANCE,
        `unexpected row height: ${measured.rowHeight}`)
      // The cap is the documented “6 rows + sticky header + the box frame” rule, and the rendered box
      // must actually respect it — this is what stops 1000 reports from stretching the page.
      assert.equal(measured.viewport.rowsVar, 6, `${rows} rows: the list is specified for six visible rows`)
      const rule = measured.viewport.rowsVar * measured.viewport.rowVar
        + measured.viewport.headVar + measured.viewport.frameVar
      assert.ok(Math.abs(measured.viewport.capPx - rule) <= 1,
        `${rows} rows: the cap must be six rows plus the header (cap ${measured.viewport.capPx}px, rule ${rule}px)`)
      assert.ok(measured.viewport.clientHeight <= measured.viewport.capPx + 1,
        `${rows} rows: rendered height ${measured.viewport.clientHeight}px must respect the cap`)
      assert.ok(Math.abs(measured.viewport.clientHeight - rule) <= 3,
        `${rows} rows: the list should be filled to its cap (got ${measured.viewport.clientHeight}px, rule ${rule}px)`)
      assert.ok(Math.abs(measured.rowHeight - measured.viewport.rowVar) <= 2,
        `${rows} rows: rendered row height ${measured.rowHeight}px must match the documented row height`)
      assert.ok(measured.viewport.scrollHeight > measured.viewport.clientHeight,
        `${rows} rows: the rows must scroll inside the list`)
      assert.equal(measured.viewport.overflowY, 'auto')
      assert.ok(measured.visibleRows >= 6 && measured.visibleRows <= 7,
        `${rows} rows: about six rows stay visible (got ${measured.visibleRows})`)
      assert.equal(measured.head.position, 'sticky', 'the column header must be sticky')
      assert.equal(measured.rowCount, rows, 'every report stays in the list — no pagination, no truncation')
      assertNoHorizontalScrollbar(measured, `${rows} rows`)
      baselines[rows] = measured
    } finally {
      await page.close()
    }
  }

  // Requirement 4: hundreds of reports must not stretch the application page.
  assert.equal(baselines[1000].page.mainScrollHeight, baselines[7].page.mainScrollHeight,
    'the workspace height must not depend on the number of reports')
  assert.equal(baselines[1000].viewport.clientHeight, baselines[7].viewport.clientHeight,
    'the list viewport height must not grow with the number of reports')
  assert.equal(baselines[1000].page.documentScrollHeight, baselines[7].page.documentScrollHeight,
    'the window height must not grow with the number of reports')
})

test('PROMPT-034R · sticky header stays pinned and every report stays reachable by scrolling', async (t) => {
  if (browserUnavailable(t)) return
  const page = await openList(1000)
  try {
    const before = await measure(page)
    assert.equal(before.viewport.scrollTop, 0)
    assert.ok(Math.abs(before.head.top - before.viewport.top) <= 1, 'the header starts at the top of the list')

    const observed = []
    for (const target of [1, 12, 250, Number.MAX_SAFE_INTEGER]) {
      const state = await page.evaluate((selector, title, scrollTarget) => {
        const section = Array.from(document.querySelectorAll('section.card'))
          .find((element) => element.querySelector('h2')?.textContent === title)
        const viewport = section.querySelector(selector)
        const table = section.querySelector('table')
        viewport.scrollTop = Math.min(scrollTarget, viewport.scrollHeight - viewport.clientHeight)
        const box = viewport.getBoundingClientRect()
        const head = table.querySelector('thead').getBoundingClientRect()
        const rows = Array.from(table.querySelectorAll('tbody tr'))
        const visible = rows.filter((row) => {
          const rect = row.getBoundingClientRect()
          return rect.top >= box.top - 1 && rect.bottom <= box.bottom + 1
        })
        const headerHit = document.elementFromPoint(head.left + head.width / 2, head.top + head.height / 2)
        return {
          scrollTop: viewport.scrollTop,
          headerOffset: Math.round(head.top - box.top),
          headerVisible: head.bottom > box.top && head.top < box.bottom,
          headerFullyInside: head.top >= box.top - 1 && head.bottom <= box.bottom + 1,
          headerHitTest: headerHit?.tagName ?? null,
          visibleStt: visible.map((row) => row.querySelector('td:nth-child(2)').textContent),
          lastRowVisible: rows.at(-1).getBoundingClientRect().bottom <= box.bottom + 1,
          pageScrollTop: document.querySelector('main').scrollTop,
          documentScrollTop: document.scrollingElement.scrollTop,
        }
      }, VIEWPORT, SECTION_TITLE, target)
      observed.push(state)
    }

    for (const state of observed) {
      assert.ok(Math.abs(state.headerOffset) <= 1,
        `the header must stay pinned to the top of the list while scrolling (offset ${state.headerOffset}px)`)
      assert.ok(state.headerVisible && state.headerFullyInside, 'the header must remain fully visible')
      assert.equal(state.headerHitTest, 'TH', 'the sticky header must stay above the scrolling rows')
      assert.equal(state.pageScrollTop, 0, 'scrolling the list must not scroll the workspace')
      assert.equal(state.documentScrollTop, 0, 'scrolling the list must not scroll the window')
      assert.equal(state.visibleStt.length, 6, 'about six rows stay visible while scrolling')
    }

    assert.deepEqual(observed.map((state) => state.scrollTop),
      [1, 12, 250, before.viewport.scrollHeight - before.viewport.clientHeight],
      'the list scrolls to the requested offsets and stops at its last row')

    const firstVisible = observed.map((state) => Number(state.visibleStt[0]))
    assert.equal(firstVisible[0], 1, 'the first report is visible at the top of the list')
    for (let index = 1; index < firstVisible.length; index += 1) {
      assert.ok(firstVisible[index] >= firstVisible[index - 1],
        `scrolling must only move forward through the dataset (${firstVisible.join(' → ')})`)
    }

    const bottom = observed.at(-1)
    assert.equal(bottom.lastRowVisible, true, 'the 1000th report must be reachable')
    assert.equal(bottom.visibleStt.at(-1), '1000', 'no report is dropped or truncated by the bounded viewport')
    assert.deepEqual(bottom.visibleStt, ['995', '996', '997', '998', '999', '1000'],
      'the bottom of the list shows exactly the tail of the dataset')
    assert.equal(before.firstRowStt, '1')
  } finally {
    await page.close()
  }
})

test('PROMPT-034R · large fixture stays structurally valid (columns, cells, order)', async (t) => {
  if (browserUnavailable(t)) return
  const page = await openList(1200)
  try {
    const measured = await measure(page)
    assert.equal(measured.rowCount, 1200, 'the 1200-report fixture must be rendered in full')
    assert.equal(measured.columnCount, 8, 'all eight columns stay present')
    assert.equal(measured.colgroupCount, 8, 'the fixed column widths stay defined')
    assert.equal(measured.cellsPerFirstRow, 8, 'each row keeps all eight cells')
    assert.ok(measured.viewport.scrollHeight > measured.viewport.clientHeight * 100,
      'the 1200 rows live behind the bounded viewport scroll')
    assert.equal(measured.visibleRows, 6)
    assert.equal(measured.head.position, 'sticky')
    assert.equal(measured.viewport.scrollHeight - measured.viewport.clientHeight,
      measured.page.documentScrollHeight < 0 ? 0 : measured.viewport.scrollHeight - measured.viewport.clientHeight)
    assertNoHorizontalScrollbar(measured, '1200 rows')
  } finally {
    await page.close()
  }
})

test('PROMPT-034R · desktop widths never gain a horizontal scrollbar', async (t) => {
  if (browserUnavailable(t)) return
  for (const width of DESKTOP_WIDTHS) {
    const page = await openList(1000, { width, height: 940 })
    try {
      const measured = await measure(page)
      assert.equal(measured.rowCount, 1000)
      assertNoHorizontalScrollbar(measured, `${width}px`)
      assert.ok(measured.viewport.scrollWidth <= measured.viewport.clientWidth,
        `${width}px: the report rows must fit the column width`)
      assert.ok(measured.visibleRows >= 5, `${width}px: the viewport still shows about six rows`)
    } finally {
      await page.close()
    }
  }
})

test('PROMPT-034R · search and status filter still work inside the bounded list', async (t) => {
  if (browserUnavailable(t)) return
  const page = await openList(30)
  try {
    assert.equal((await measure(page)).rowCount, 30)

    // Search: exactly one Management Number matches.
    await setSearch(page, '260900012')
    await waitForRows(page, 1)
    const searched = await measure(page)
    assert.equal(searched.firstRowStt, '12', 'the search narrows the list to the matching report')
    assert.ok(searched.viewport.clientHeight < 150,
      `a one-row search result must not keep the full list height (got ${searched.viewport.clientHeight}px)`)
    assert.equal(searched.viewport.scrollHeight - searched.viewport.clientHeight, 0,
      'a one-row search result must not need an internal scrollbar')

    // A search with no match keeps the report area compact — no blank box, no oversized empty state.
    await setSearch(page, 'khong-co-bao-cao-nao')
    await page.waitForFunction(() => document.body.innerText.includes('Không có báo cáo khớp bộ lọc'),
      { timeout: 10_000 })
    const empty = await measure(page)
    assert.equal(empty.rowCount, 0)
    assert.ok(empty.cardHeight < 350, `an empty filter result must not reserve a blank list (${empty.cardHeight}px)`)

    // Clear the search and filter by status instead.
    await setSearch(page, '')
    await waitForRows(page, 30)

    await page.select('select', 'File lỗi')
    await waitForRows(page, 6)
    const filtered = await measure(page)
    assert.equal(filtered.rowCount, 6, 'the status filter keeps only the six error reports')
    const chipLabels = await page.$$eval(`${VIEWPORT_LIST} tbody tr td:nth-child(7)`,
      (cells) => cells.map((cell) => cell.textContent.trim()))
    assert.deepEqual(chipLabels, Array(6).fill('Lỗi'), 'the status column still reports the filtered state')
    assert.equal(filtered.viewport.scrollHeight - filtered.viewport.clientHeight, 0,
      'six filtered rows fit exactly — no scrollbar, no blank space')

    await page.select('select', 'Tất cả file đã quét')
    await waitForRows(page, 30)
    assert.equal((await measure(page)).rowCount, 30, 'clearing the filter restores the full list')
  } finally {
    await page.close()
  }
})

test('PROMPT-034R · selection and select-all still cover the whole dataset', async (t) => {
  if (browserUnavailable(t)) return
  const page = await openList(30)
  try {
    // The report list box itself, found the same way for the pre-fix and post-fix markup.
    const rack = await page.$(`${VIEWPORT_LIST}, div.overflow-auto`)
    assert.ok(rack, 'the report list box must be present')
    const CHECKBOX = 'tbody button[aria-label^="Chọn "]'
    const checkedCount = async () => {
      const boxes = await rack.$$(CHECKBOX)
      const states = await Promise.all(boxes.map((box) => box.evaluate((el) => el.className.includes('bg-brand-500'))))
      return states.filter(Boolean).length
    }
    const waitForChecked = (want) => page.waitForFunction((selectorList, title, expected) => {
      const section = Array.from(document.querySelectorAll('section.card'))
        .find((element) => element.querySelector('h2')?.textContent === title)
      const viewport = section?.querySelector(selectorList)
      const boxes = Array.from(viewport?.querySelectorAll('tbody button[aria-label^="Chọn "]') ?? [])
      if (boxes.length !== 30) return false
      return boxes.filter((box) => box.className.includes('bg-brand-500')).length === expected
    }, { timeout: 10_000 }, VIEWPORT, SECTION_TITLE, want)

    assert.equal(await checkedCount(), 0)
    await (await rack.$('thead button[aria-label="Chọn tất cả"]')).click()
    await waitForChecked(30)
    assert.equal(await checkedCount(), 30,
      'select-all must select every one of the 30 reports, not only the six visible rows')

    await (await rack.$('thead button[aria-label="Chọn tất cả"]')).click()
    await waitForChecked(0)
    assert.equal(await checkedCount(), 0, 'select-all must toggle off again')

    await (await rack.$('tbody tr:nth-child(4) button[aria-label^="Chọn "]')).click()
    await waitForChecked(1)
    assert.equal(await checkedCount(), 1, 'a single row can still be selected')
    const rows = await rack.$$('tbody tr')
    const outlines = await Promise.all(rows.map((row) => row.evaluate((el) => el.className.includes('outline-brand-400'))))
    assert.ok(outlines.filter(Boolean).length <= 1, 'the active-row highlight stays exclusive')
  } finally {
    await page.close()
  }
})

test('PROMPT-034R · incremental scan updates still render into the bounded list', async (t) => {
  if (browserUnavailable(t)) return
  const page = await openList(8)
  try {
    const start = await measure(page)
    assert.ok(start.viewport.scrollHeight > start.viewport.clientHeight, 'eight rows already scroll')

    await page.evaluate(() => window.__tnpLayoutHarness.appendIncrementalRows(3))
    await waitForRows(page, 11)
    const grown = await measure(page)
    assert.equal(grown.rowCount, 11, 'a completed scan leaf adds its new reports incrementally')
    assert.equal(grown.viewport.clientHeight, start.viewport.clientHeight,
      'incremental updates must not enlarge the bounded viewport')
    assert.equal(grown.lastRowStt, '11')

    const rack = await page.$(`${VIEWPORT_LIST}, div.overflow-auto`)
    const statusChips = async () => {
      const cells = await rack.$$('tbody tr td:nth-child(7)')
      return Promise.all(cells.map((cell) => cell.evaluate((el) => el.textContent.trim())))
    }

    await page.evaluate(() => window.__tnpLayoutHarness.setStatus(0, 'processing'))
    await page.waitForFunction((selectorList, title) => {
      const section = Array.from(document.querySelectorAll('section.card'))
        .find((element) => element.querySelector('h2')?.textContent === title)
      const cell = section?.querySelector(selectorList)?.querySelector('tbody tr td:nth-child(7)')
      return Boolean(cell?.textContent.trim().startsWith('Đang xử lý'))
    }, { timeout: 10_000 }, VIEWPORT, SECTION_TITLE)
    const processing = await statusChips()
    assert.equal(processing[0].startsWith('Đang xử lý'), true, 'processing state still renders per row')
    assert.equal(processing[1], 'Hoàn thành', 'other rows keep their own status')

    await page.evaluate(() => window.__tnpLayoutHarness.setStatus(2, 'error'))
    await page.waitForFunction((selectorList, title) => {
      const section = Array.from(document.querySelectorAll('section.card'))
        .find((element) => element.querySelector('h2')?.textContent === title)
      const cell = section?.querySelector(selectorList)?.querySelector('tbody tr:nth-child(3) td:nth-child(7)')
      return cell?.textContent.trim() === 'Lỗi'
    }, { timeout: 10_000 }, VIEWPORT, SECTION_TITLE)
    assert.equal((await measure(page)).viewport.clientHeight, start.viewport.clientHeight,
      'status updates never resize the list viewport')
  } finally {
    await page.close()
  }
})

test('PROMPT-034R · a 1000-report list adds no per-row native bridge traffic', async (t) => {
  if (browserUnavailable(t)) return
  const trafficWindow = 3_200

  /** Native traffic during one window; a per-row implementation would explode these numbers. */
  const traffic = async (page) => {
    const before = await page.evaluate(() => window.__tnpLayoutHarness.bridgeCalls())
    await delay(trafficWindow)
    const after = await page.evaluate(() => window.__tnpLayoutHarness.bridgeCalls())
    const delta = {}
    for (const [method, count] of Object.entries(after)) delta[method] = count - (before[method] ?? 0)
    const polls = delta.get_dashboard_state ?? 0
    return { delta, polls, total: Object.values(delta).reduce((sum, value) => sum + value, 0) }
  }

  const bigPage = await openList(1000)
  let big
  try {
    big = await traffic(bigPage)
  } finally {
    await bigPage.close()
  }
  const smallPage = await openList(6)
  let small
  try {
    small = await traffic(smallPage)
  } finally {
    await smallPage.close()
  }

  for (const measured of [big, small]) {
    assert.ok(measured.polls >= 2, `the dashboard poll must keep running (saw ${measured.polls} polls)`)
    // The only per-report call in the app fetches the ACTIVE report's detail, and the dashboard snapshot
    // refreshes it on each poll: two per poll is the ceiling. The row count is not a factor.
    assert.ok((measured.delta.get_report_details ?? 0) <= 2 * measured.polls + 2,
      `detail fetches must stay bounded by the poll cadence: ${JSON.stringify(measured.delta)}`)
    assert.ok(measured.total <= 4 * measured.polls + 4,
      `no bridge call may fan out per report: ${JSON.stringify(measured.delta)}`)
    assert.ok(measured.total < 40, `bridge traffic must stay small: ${JSON.stringify(measured.delta)}`)
    assert.equal(measured.delta.ping ?? 0, 0, 'the list must not trigger bridge reconnects')
  }
  assert.ok(Math.abs(big.total - small.total) <= 4,
    `1000 rows must cost the same native traffic as 6 rows `
    + `(1000: ${JSON.stringify(big.delta)}, 6: ${JSON.stringify(small.delta)})`)
})
