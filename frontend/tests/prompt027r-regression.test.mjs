/**
 * PROMPT-027R — Learning preview fidelity and overlay truth (source-level guards).
 *
 * These checks read the production sources, so they run without a browser and fail loudly if a later change:
 *   • restyles the slide bitmap (blur / brightness / contrast / saturation / opacity / dark overlays / clip copies)
 *     — the authored render must reach the reviewer unchanged (PROMPT-027R §18/§19);
 *   • recomputes overlay geometry in React instead of using the Python DTO (§15/§23/§25/§26);
 *   • uses the picture-membership list for anything but display/debug (§24);
 *   • moves a hook after an early return (React #310 protection, §44).
 * The rendered counterpart lives in prompt027-overlays.test.mjs (one unmodified <img>, outline overlays only).
 */
import assert from 'node:assert/strict'
import { test } from 'node:test'
import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'

const frontendRoot = fileURLToPath(new URL('..', import.meta.url))
const learningTab = await readFile(`${frontendRoot}/src/tabs/LearningTab.tsx`, 'utf8')
const types = await readFile(`${frontendRoot}/src/types.ts`, 'utf8')

/** Source with line and block comments removed, so documentation can name the forbidden tokens safely. */
function codeOnly(source) {
  return source
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .split('\n')
    .map((line) => line.replace(/\/\/.*$/, ''))
    .join('\n')
}

/** The body of a top-level function declaration in LearningTab.tsx (from its signature to the next top-level one). */
function functionBody(source, name) {
  const start = source.indexOf(`function ${name}(`)
  assert.ok(start >= 0, `function ${name} must exist`)
  const next = source.slice(start + 1).search(/\n(?:function |export function |const [A-Z_]+ =)/)
  return next < 0 ? source.slice(start) : source.slice(start, start + 1 + next)
}

const FORBIDDEN_BITMAP_EFFECTS = [
  /filter\s*:/, /brightness\s*\(/, /contrast\s*\(/, /saturate\s*\(/, /blur\s*\(/, /grayscale\s*\(/,
  /backdrop-filter/, /mix-blend-mode/, /clipPath/, /opacity\s*:/, /bg-black\/\d/, /bg-black\b/,
]

test('the slide bitmap is never restyled: no filter, dimming, opacity or clipped copy in the preview (§18/§19)', () => {
  const code = codeOnly(learningTab)
  for (const pattern of FORBIDDEN_BITMAP_EFFECTS) {
    assert.doesNotMatch(code, pattern, `LearningTab must not use ${pattern} on the review surface`)
  }
  // exactly one slide bitmap element in the preview
  const bitmapImages = code.match(/<img[\s\S]*?data-testid="slide-preview-bitmap"/g) || []
  assert.equal(bitmapImages.length, 1, 'one unmodified slide bitmap')
  const previewImgs = code.match(/src=\{cand\.slidePreview\}/g) || []
  assert.equal(previewImgs.length, 1, 'the slide preview src is used exactly once')
})

test('overlay geometry comes from the Python DTO percentages — React never derives crop geometry (§15/§23/§25)', () => {
  const evidence = functionBody(learningTab, 'evidenceRect')
  assert.match(evidence, /c\.evidenceRegionBboxPct/, 'evidence overlay must use the production region percentages')
  assert.doesNotMatch(codeOnly(evidence), /bounds|targetBbox|slidePreview|pictureIds|evidenceRegionPictureIds/,
    'the evidence region must never be derived from the picture candidate')

  const target = functionBody(learningTab, 'targetRect')
  assert.match(target, /c\.targetBboxPct/, 'candidate overlay must use the picture target percentages')

  const code = codeOnly(learningTab)
  assert.match(code, /data-testid="overlay-evidence-region"/)
  assert.match(code, /data-testid="overlay-picture-candidate"/)
  // the evidence overlay is positioned from the evidence rectangle, the candidate overlay from the target rectangle
  assert.match(code, /overlay-evidence-region[\s\S]*?left: `\$\{evidence\.x\}%`/)
  assert.match(code, /overlay-picture-candidate[\s\S]*?left: `\$\{rect\.x\}%`/)
  assert.doesNotMatch(code, /evidenceRegionBbox\.(x|y|width|height)\s*\*/, 'no pixel conversion of the EMU box in React')
})

test('evidenceRegionPictureIds is display/debug data only (§24)', () => {
  const code = codeOnly(learningTab)
  const uses = code.match(/evidenceRegionPictureIds/g) || []
  assert.ok(uses.length >= 1, 'the membership list is shown for acceptance')
  // the two geometry helpers never read the membership list
  for (const name of ['evidenceRect', 'targetRect']) {
    assert.doesNotMatch(codeOnly(functionBody(learningTab, name)), /evidenceRegionPictureIds/,
      `${name} must not read the membership list`)
  }
  // and it never feeds an overlay style
  const overlayStyles = code.match(/style=\{\{[^}]*evidenceRegionPictureIds[^}]*\}\}/g) || []
  assert.equal(overlayStyles.length, 0)
})

test('the DTO exposes the membership fields with the documented names (§24)', () => {
  for (const field of ['evidenceRegionPictureIds', 'evidenceRegionId', 'evidenceRegionBlockIndex',
    'evidenceRegionBoundarySource', 'evidenceRegionNextHeadingTop']) {
    assert.match(types, new RegExp(`${field}\\??:`), `types.ts must declare ${field}`)
  }
})

test('the preview keeps the honest backend notice and never hides a fallback (§6/§41)', () => {
  const code = codeOnly(learningTab)
  assert.match(code, /Bản xem trước đơn giản — bố cục có thể khác PowerPoint/)
  assert.match(code, /showFidelityNotice = Boolean\(previewBackend\) && !previewFaithful/)
})

test('hooks in ImageReview stay before every conditional early return (React #310 protection, §44)', () => {
  const body = functionBody(learningTab, 'ImageReview')
  const earlyReturn = body.indexOf('if (!cand) return')
  assert.ok(earlyReturn > 0, 'the empty-state early return must exist')
  const hookCalls = [...body.matchAll(/\b(useState|useEffect|useMemo|useRef|useCallback|useLayoutEffect)\(/g)]
  assert.ok(hookCalls.length >= 6, 'the review keeps its hooks')
  for (const hook of hookCalls) {
    assert.ok(hook.index < earlyReturn, `${hook[1]} is declared after the early return`)
  }
})

test('no new hook or conditional return was introduced after the candidate guard', () => {
  const body = functionBody(learningTab, 'ImageReview')
  const afterGuard = body.slice(body.indexOf('if (!cand) return'))
  assert.doesNotMatch(afterGuard, /\buse[A-Z][A-Za-z]*\(/, 'hooks must not appear after the candidate early return')
})
