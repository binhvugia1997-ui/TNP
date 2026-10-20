/**
 * PROMPT-041W — the reduced-fidelity notice must not claim a process fact it cannot prove.
 *
 * The Windows acceptance run of PR #17 showed this exact sentence with `Get-Process POWERPNT` returning ZERO rows:
 *
 *     "Một phiên PowerPoint đang chạy trên máy này nên app không tạo được phiên riêng.
 *      Đóng PowerPoint rồi chọn lại ảnh."
 *
 * That text is produced by exactly one slug, `powerpoint-registered-instance`, and on that machine the slug came
 * from the Running Object Table — which still holds an entry for a PowerPoint that has already exited — not from
 * the process table, which agreed with `Get-Process`. Python now:
 *   • no longer turns a ROT entry into a process claim (the process table is the only source of process facts);
 *   • emits a NEW slug, `powerpoint-session-unproven`, for "a PowerPoint was seen but our own Windows session
 *     could not be read" — deliberately distinct from `powerpoint-user-session-running`;
 *   • keeps `powerpoint-registered-instance` only as a fail-closed refusal for previews persisted before the fix.
 *
 * So the wording has to change too: the legacy slug may no longer tell the reviewer "a PowerPoint session is
 * running on this machine", because that is exactly the lie the reviewer disproved with `Get-Process`.
 */
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { test } from 'node:test'
import { createServer } from 'vite'
import { fileURLToPath } from 'node:url'

const frontendRoot = fileURLToPath(new URL('..', import.meta.url))
const repoRoot = fileURLToPath(new URL('../..', import.meta.url))
const LEARNING_TAB = `${frontendRoot}/src/tabs/LearningTab.tsx`

/** Every slug Python can put in `slidePreviewReason`, read from the authoritative Python source. */
function pythonReasonSlugs() {
  const source = readFileSync(`${repoRoot}/app/qpn_renderer.py`, 'utf8')
  const assignments = new Map()
  for (const match of source.matchAll(/^REASON_[A-Z_]+ = "([a-z0-9-]+)"$/gm)) {
    assignments.set(match[0].split(' ')[0], match[1])
  }
  const present = source.match(/POWERPOINT_PRESENT_REASONS = frozenset\(\{([^}]*)\}\)/s)
  assert.ok(present, 'POWERPOINT_PRESENT_REASONS must exist in app/qpn_renderer.py')
  const names = [...present[1].matchAll(/REASON_[A-Z_]+/g)].map((m) => m[0])
  return { all: [...assignments.values()], present: names.map((n) => assignments.get(n)).filter(Boolean) }
}

async function loadFidelityHint() {
  const server = await createServer({
    configFile: `${frontendRoot}/vite.config.ts`, root: frontendRoot,
    appType: 'custom', server: { middlewareMode: true, hmr: false },
  })
  try {
    const module = await server.ssrLoadModule('/src/tabs/LearningTab.tsx')
    return module.fidelityHint
  } finally {
    await server.close()
  }
}

/* ------------------------------------------------------- the false positive itself */
test('the "a PowerPoint session is running on this machine" claim is gone', async () => {
  const tsx = readFileSync(LEARNING_TAB, 'utf8')
  assert.doesNotMatch(
    tsx,
    /Một phiên PowerPoint đang chạy trên máy này/,
    'no slug may assert a machine-wide running PowerPoint — `Get-Process POWERPNT` disproved that on Windows',
  )
  // NOTE: "Đóng PowerPoint" on its own stays — it is the CORRECT remedy for `powerpoint-user-session-running`,
  // where the reviewer's own PowerPoint really is open.  Only the machine-wide *claim* above is a lie.

  const fidelityHint = await loadFidelityHint()
  const legacy = fidelityHint('powerpoint-registered-instance')
  assert.ok(legacy.length > 10, 'the legacy slug keeps wording: previews persisted before the fix still carry it')
  assert.doesNotMatch(legacy, /đang chạy trên máy này/, 'it may not claim a process is running')
  assert.doesNotMatch(legacy, /Đóng PowerPoint/, 'it may not tell the reviewer to close something that is not open')
  assert.match(legacy, /thử render lại/i, 'the honest remedy is to retry, not to hunt for a phantom process')
})

/* ------------------------------------------------------- the new distinct slug */
test('powerpoint-session-unproven has its own wording, distinct from powerpoint-user-session-running', async () => {
  const fidelityHint = await loadFidelityHint()
  const unproven = fidelityHint('powerpoint-session-unproven')
  const userSession = fidelityHint('powerpoint-user-session-running')

  assert.ok(unproven.length > 10, 'the new slug must produce an explanation')
  assert.notEqual(unproven, userSession, 'contract D: "session unprovable" is not "your session is in the way"')
  assert.match(unproven, /phiên Windows/i, 'it must name what could not be read: our own Windows session')
  assert.doesNotMatch(unproven, /Đóng PowerPoint/, 'nothing to close: we could not even prove a session')
  assert.match(userSession, /Đóng PowerPoint/, '...while the proven same-session case still says close it')
})

/* ------------------------------------------------------- the two sides cannot drift */
test('every reason slug Python can emit has wording, and none is left unmapped', async () => {
  const fidelityHint = await loadFidelityHint()
  const { all, present } = pythonReasonSlugs()
  assert.ok(all.includes('powerpoint-session-unproven'), 'the new slug must exist in Python')
  assert.ok(present.includes('powerpoint-session-unproven'), '...and must count as a PowerPoint-blocked fallback')

  for (const slug of all) {
    if (slug === 'ok') continue
    assert.ok(fidelityHint(slug).length > 10, `${slug} must produce an explanation`)
  }
  assert.equal(fidelityHint(''), '', 'no reason → no extra sentence')
  assert.equal(fidelityHint('some-future-code'), '', 'an unknown slug degrades to the generic notice')

  const tsx = readFileSync(LEARNING_TAB, 'utf8')
  for (const slug of all) {
    assert.ok(tsx.includes(`'${slug}'`), `LearningTab.tsx has no case for ${slug}`)
  }
})
