import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { fileURLToPath } from 'node:url'
import { createServer } from 'vite'

const frontendRoot = fileURLToPath(new URL('..', import.meta.url))
let server
let normalizeUpdate

before(async () => {
  server = await createServer({
    configFile: `${frontendRoot}/vite.config.ts`,
    root: frontendRoot,
    appType: 'custom',
    server: { middlewareMode: true, hmr: false },
  })
  ;({ normalizeUpdate } = await server.ssrLoadModule('/src/state/store.tsx'))
})

after(async () => {
  await server?.close()
})

const available = (overrides = {}) => ({
  path: 'configured-update', autoCheck: true, status: 'available',
  message: 'Có phiên bản mới: 1.3.6 — Build 019',
  available: { version: '1.3.6', build: '019', package: 'ReportExtractor_1.3.6.zip', size: 1_234_567 },
  progress: 0, progressStage: '', installEnabled: true, canInstall: true,
  installDisabledReason: '', error: '', ...overrides,
})

test('a valid production update DTO preserves backend install permission and package metadata', () => {
  const result = normalizeUpdate(available())
  assert.equal(result.status, 'available')
  assert.equal(result.installEnabled, true)
  assert.equal(result.canInstall, true)
  assert.equal(result.available.build, '019')
  assert.equal(result.available.package, 'ReportExtractor_1.3.6.zip')
})

test('a development runtime can discover an update but stays explicitly non-installable', () => {
  const reason = 'Chỉ có thể cài đặt cập nhật từ ứng dụng Portable đã đóng gói.'
  const result = normalizeUpdate(available({ installEnabled: false, canInstall: false, installDisabledReason: reason }))
  assert.equal(result.status, 'available')
  assert.equal(result.installEnabled, false)
  assert.equal(result.canInstall, false)
  assert.equal(result.installDisabledReason, reason)
})

test('incomplete or contradictory update DTOs fail closed and never authorize installation', () => {
  for (const raw of [
    (() => { const { canInstall, ...legacy } = available(); return legacy })(),
    available({ installEnabled: true, canInstall: false }),
    available({ status: 'available', available: null }),
    available({ available: { version: '1.3.6', build: '../019', package: 'x.zip', size: 1 } }),
  ]) {
    const result = normalizeUpdate(raw)
    assert.equal(result.status, 'failed')
    assert.equal(result.installEnabled, false)
    assert.equal(result.canInstall, false)
  }
})

test('busy staging progress is valid state but never becomes a second install authorization', () => {
  const result = normalizeUpdate(available({ status: 'staging', installEnabled: false,
    progress: 37, progressStage: 'COPYING', installDisabledReason: 'Đang chuẩn bị cập nhật.' }))
  assert.equal(result.status, 'staging')
  assert.equal(result.progress, 37)
  assert.equal(result.installEnabled, false)
})
