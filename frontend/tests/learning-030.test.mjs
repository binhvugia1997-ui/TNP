import assert from 'node:assert/strict'
import { test } from 'node:test'
import { JSDOM } from 'jsdom'
import { createServer } from 'vite'
import * as React from 'react'
import { fileURLToPath } from 'node:url'

const frontendRoot = fileURLToPath(new URL('..', import.meta.url))
const delay = (ms) => new Promise((resolve) => globalThis.setTimeout(resolve, ms))

function emptyLearningState(images = [], contents = []) {
  return {
    images, contents,
    counts: { image: { total: images.length, labeled: 0 }, content: { total: 0, labeled: 0 } },
    modelStatus: { image: 'Ready', content: 'Ready' },
    excelPending: 0, excelLastResult: { kind: 'none', message: '' }, training: false,
  }
}
function imageCandidate(id, slide, opts = {}) {
  const { pictureId = `p${slide}`, note = '', userLabel = 'UNLABELED', slidePreview = 'data:image/jpeg;base64,AA==' } = opts
  return {
    id, sourceFile: 'report.pptx', managementNumber: 'TNP-001', slide, pictureId,
    src: 'data:image/png;base64,BB==',
    bounds: { x: 10, y: 10, w: 20, h: 20 }, slideWidth: 12192000, slideHeight: 6858000,
    targetBbox: { x: 100, y: 100, width: 200, height: 200 },
    targetBboxPct: { x: 10, y: 10, w: 20, h: 20 }, targetKind: 'picture',
    slidePreview, slidePreviewWidth: 1600, slidePreviewHeight: 900,
    itemId: `item-${slide}`, itemIndex: opts.itemIndex ?? (slide - 1), itemHeading: `Heading ${slide}`,
    decision: 'review', confidenceBand: 'medium', confidence: 0.5, evidence: ['ev'], excelEligible: false, eligibilityReason: '',
    nearbyText: '', userLabel, labelPending: false, note,
  }
}
async function makeHarness(learningResponse) {
  const dom = new JSDOM('<!doctype html><html><body><div id="root"></div></body></html>', { url: 'http://localhost/' })
  const prior = new Map(['window','document','navigator','HTMLElement','Node','Event','MouseEvent','IS_REACT_ACT_ENVIRONMENT','ResizeObserver'].map(n=>[n, Object.getOwnPropertyDescriptor(globalThis, n)]))
  const def = (n,v)=> Object.defineProperty(globalThis, n, { configurable:true, writable:true, value:v })
  def('window', dom.window); def('document', dom.window.document); def('navigator', dom.window.navigator)
  def('HTMLElement', dom.window.HTMLElement); def('Node', dom.window.Node)
  def('Event', dom.window.Event); def('MouseEvent', dom.window.MouseEvent)
  def('IS_REACT_ACT_ENVIRONMENT', true)
  class Obs { observe(){} unobserve(){} disconnect(){} }
  def('ResizeObserver', Obs); dom.window.ResizeObserver = Obs
  dom.window.HTMLElement.prototype.scrollIntoView = () => {}
  dom.window.matchMedia = (m) => ({ matches: m==='(min-width: 1280px)', media:m, onchange:null, addEventListener(){}, removeEventListener(){}, addListener(){}, removeListener(){}, dispatchEvent(){return true}})
  const server = await createServer({ configFile: `${frontendRoot}/vite.config.ts`, root: frontendRoot, appType: 'custom', server: { middlewareMode: true, hmr: false } })
  const [{ StoreProvider, useStore }, { LearningTab }] = await Promise.all([
    server.ssrLoadModule('/src/state/store.tsx'),
    server.ssrLoadModule('/src/tabs/LearningTab.tsx'),
  ])
  const { createRoot } = await import('react-dom/client')
  const config = { paths:{reportFolder:'',template:'',output:''}, period:{mode:'auto',month:'',year:'',from:'',to:''}, forceReprocess:false, ollama:{host:'',port:'11434',model:'',checked:'unchecked',message:'',models:[]}, update:{path:'',autoCheck:true} }
  const dashboard = { app:{name:'Report Extractor',title:'Report Extractor v1.3.2',version:'1.3.2',build:'015',buildNumber:15}, config, scan:{scanned:false,message:'No reports'}, reports:[], job:{status:'idle',queue:[],index:0,doneCount:0,stage:'waiting',percent:0,currentFile:'',elapsedSec:0,remainSec:null,startedAt:null,finishedAt:null,hasSamples:false}, logs:[], ollama:{host:'',port:'11434',model:'',checked:'unchecked',message:'',models:[],aiStatus:'',discovering:false,discoveryMessage:'',discoveryResults:[],discoveryChecked:0,discoveryTotal:0,serverApplyRunning:false,serverApplyMessage:''}, update:{path:'',autoCheck:true,status:'idle',message:'',progress:0,progressStage:''}, diagnostics:{running:false,rows:[],error:''} }
  const env = (d)=> ({ ok:true, data:d })
  let current = learningResponse
  let saveCount = 0
  const pywebview = {
    ping: async()=> env('Python bridge OK'),
    get_app_version: async()=> env({name:'Report Extractor',title:'Report Extractor v1.3.2',version:'1.3.2',build:'015'}),
    get_current_config: async()=> env(config),
    get_dashboard_state: async()=> env(dashboard),
    get_learning_state: async()=> env(current),
    set_learning_label: async(kind,id,label, note='')=>{ current = {...current, images: current.images.map(c=> c.id===id? {...c, userLabel: label, labelPending:true}:c)}; return env(current) },
    save_learning_labels: async(kind, notes)=>{ saveCount+=1; current = {...current, images: current.images.map(c=> ({...c, labelPending:false}))}; return env({message:'Đã lưu', state: current, saveCount}) },
  }
  function Harness(){
    const s = useStore()
    return React.createElement(React.Fragment, null,
      React.createElement('output', {'data-testid':'bridge-state'}, s.connected?'ready':'disconnected'),
      React.createElement('button', {'data-testid':'refresh-learning', onClick:()=>{ void s.refreshLearning()}}, 'Refresh'),
      React.createElement(LearningTab),
    )
  }
  return { dom, server, StoreProvider, Harness, createRoot, def, prior, getCurrent:()=>current, getSaveCount:()=>saveCount, setCurrent:(v)=>{current=v}, dashboard, config, pywebview }
}
async function waitFor(pred, desc, timeoutMs=3000){ const d=Date.now()+timeoutMs; while(Date.now()<d){ if(pred()) return; await React.act(async()=>{ await delay(10) }) } assert.fail(`Timed out: ${desc}`) }

test('label controls rendered near top of right panel', async () => {
  const h = await makeHarness(emptyLearningState([imageCandidate('a',1), imageCandidate('b',2)]))
  let root
  try{
    const { dom, server, StoreProvider, Harness, createRoot, pywebview } = h
    root = createRoot(dom.window.document.getElementById('root'))
    await React.act(async()=>{ root.render(React.createElement(React.StrictMode, null, React.createElement(StoreProvider, null, React.createElement(Harness)))) })
    dom.window.pywebview = { api: pywebview }
    await React.act(async()=>{ dom.window.dispatchEvent(new dom.window.Event('pywebviewready')); await delay(60) })
    await waitFor(()=> h.dom.window.document.querySelector('[data-testid="bridge-state"]')?.textContent==='ready', 'bridge ready')
    await React.act(async()=>{ dom.window.document.querySelector('[data-testid="refresh-learning"]').dispatchEvent(new dom.window.MouseEvent('click',{bubbles:true})); await delay(30) })
    await delay(120)
    const container = dom.window.document.getElementById('root')
    const btn = [...container.querySelectorAll('aside nav button')].find(b=> b.textContent.trim().includes('Kiểm tra ảnh cải tiến'))
    assert.ok(btn)
    await React.act(async()=>{ btn.dispatchEvent(new dom.window.MouseEvent('click',{bubbles:true})); await delay(10) })
    await waitFor(()=> container.querySelector('section[aria-label="Thông tin và nhãn xác nhận"]'), 'right panel')
    const right = container.querySelector('section[aria-label="Thông tin và nhãn xác nhận"]')
    const html = right.innerHTML
    const labelIdx = html.indexOf('Chọn nhãn xác nhận')
    const srcIdx = html.indexOf('Report nguồn')
    assert.ok(labelIdx!==-1 && srcIdx!==-1 && labelIdx < srcIdx, `Label (${labelIdx}) must be before Report nguồn (${srcIdx})`)
  } finally { if(root) await React.act(async()=>{root.unmount()}); if(h.server) await h.server.close(); h.dom.window.close(); for(const [k,v] of h.prior){ if(v) Object.defineProperty(globalThis,k,v); else delete globalThis[k] } }
})

test('single-click selects without confirmation', async () => {
  const h = await makeHarness(emptyLearningState([imageCandidate('a',1), imageCandidate('b',1)]))
  let root
  try{
    const { dom, server, StoreProvider, Harness, createRoot, pywebview } = h
    root = createRoot(dom.window.document.getElementById('root'))
    await React.act(async()=>{ root.render(React.createElement(React.StrictMode, null, React.createElement(StoreProvider, null, React.createElement(Harness)))) })
    dom.window.pywebview = { api: pywebview }
    await React.act(async()=>{ dom.window.dispatchEvent(new dom.window.Event('pywebviewready')); await delay(60) })
    await waitFor(()=> dom.window.document.querySelector('[data-testid="bridge-state"]')?.textContent==='ready', 'bridge ready')
    await React.act(async()=>{ dom.window.document.querySelector('[data-testid="refresh-learning"]').dispatchEvent(new dom.window.MouseEvent('click',{bubbles:true})); await delay(30) })
    await delay(100)
    const container = dom.window.document.getElementById('root')
    const btn = [...container.querySelectorAll('aside nav button')].find(b=> b.textContent.trim().includes('Kiểm tra ảnh cải tiến'))
    await React.act(async()=>{ btn.dispatchEvent(new dom.window.MouseEvent('click',{bubbles:true})); await delay(10) })
    await waitFor(()=> container.querySelectorAll('section[aria-label="Đối tượng ảnh"] button').length===2, '2 items')
    const before = h.getSaveCount()
    const leftBtns = [...container.querySelectorAll('section[aria-label="Đối tượng ảnh"] button')]
    await React.act(async()=>{ leftBtns[1].dispatchEvent(new dom.window.MouseEvent('click',{bubbles:true})); await delay(80) })
    assert.equal(h.getSaveCount(), before, 'single click must not trigger save')
  } finally { if(root) await React.act(async()=>{root.unmount()}); if(h.server) await h.server.close(); h.dom.window.close(); for(const [k,v] of h.prior){ if(v) Object.defineProperty(globalThis,k,v); else delete globalThis[k] } }
})

test('double-click without active label does not guess', async () => {
  const h = await makeHarness(emptyLearningState([imageCandidate('a',1)]))
  let root
  try{
    const { dom, server, StoreProvider, Harness, createRoot, pywebview } = h
    root = createRoot(dom.window.document.getElementById('root'))
    await React.act(async()=>{ root.render(React.createElement(React.StrictMode, null, React.createElement(StoreProvider, null, React.createElement(Harness)))) })
    dom.window.pywebview = { api: pywebview }
    await React.act(async()=>{ dom.window.dispatchEvent(new dom.window.Event('pywebviewready')); await delay(60) })
    await waitFor(()=> dom.window.document.querySelector('[data-testid="bridge-state"]')?.textContent==='ready', 'bridge ready')
    await React.act(async()=>{ dom.window.document.querySelector('[data-testid="refresh-learning"]').dispatchEvent(new dom.window.MouseEvent('click',{bubbles:true})); await delay(30) })
    await delay(100)
    const container = dom.window.document.getElementById('root')
    const nav = [...container.querySelectorAll('aside nav button')].find(b=> b.textContent.trim().includes('Kiểm tra ảnh cải tiến'))
    await React.act(async()=>{ nav.dispatchEvent(new dom.window.MouseEvent('click',{bubbles:true})); await delay(10) })
    await waitFor(()=> container.querySelector('section[aria-label="Đối tượng ảnh"] button'), 'left button')
    const leftBtn = container.querySelector('section[aria-label="Đối tượng ảnh"] button')
    const before = h.getSaveCount()
    await React.act(async()=>{ leftBtn.dispatchEvent(new dom.window.MouseEvent('dblclick',{bubbles:true})); await delay(250) })
    assert.equal(h.getSaveCount(), before, 'double-click without label must not save')
    assert.ok(container.textContent.includes('Vui lòng chọn nhãn'), `should hint, got ${container.textContent.slice(0,400)}`)
  } finally { if(root) await React.act(async()=>{root.unmount()}); if(h.server) await h.server.close(); h.dom.window.close(); for(const [k,v] of h.prior){ if(v) Object.defineProperty(globalThis,k,v); else delete globalThis[k] } }
})

test('Ctrl multi and Shift range within same slide, bulk indicator visible', async () => {
  const h = await makeHarness(emptyLearningState([
    imageCandidate('s5-1',5,{itemIndex:0}), imageCandidate('s5-2',5,{itemIndex:1}), imageCandidate('s5-3',5,{itemIndex:2}),
  ]))
  let root
  try{
    const { dom, server, StoreProvider, Harness, createRoot, pywebview } = h
    root = createRoot(dom.window.document.getElementById('root'))
    await React.act(async()=>{ root.render(React.createElement(React.StrictMode, null, React.createElement(StoreProvider, null, React.createElement(Harness)))) })
    dom.window.pywebview = { api: pywebview }
    await React.act(async()=>{ dom.window.dispatchEvent(new dom.window.Event('pywebviewready')); await delay(60) })
    await waitFor(()=> dom.window.document.querySelector('[data-testid="bridge-state"]')?.textContent==='ready', 'bridge ready')
    await React.act(async()=>{ dom.window.document.querySelector('[data-testid="refresh-learning"]').dispatchEvent(new dom.window.MouseEvent('click',{bubbles:true})); await delay(30) })
    await delay(100)
    const container = dom.window.document.getElementById('root')
    const nav = [...container.querySelectorAll('aside nav button')].find(b=> b.textContent.trim().includes('Kiểm tra ảnh cải tiến'))
    await React.act(async()=>{ nav.dispatchEvent(new dom.window.MouseEvent('click',{bubbles:true})); await delay(10) })
    await waitFor(()=> container.querySelectorAll('section[aria-label="Đối tượng ảnh"] button').length===3, '3 items')
    const btns = ()=> [...container.querySelectorAll('section[aria-label="Đối tượng ảnh"] button')]
    await React.act(async()=>{ btns()[1].dispatchEvent(new dom.window.MouseEvent('click',{bubbles:true, ctrlKey:true})); await delay(30) })
    await React.act(async()=>{ btns()[2].dispatchEvent(new dom.window.MouseEvent('click',{bubbles:true, ctrlKey:true})); await delay(30) })
    let sel = container.querySelectorAll('[data-selected="true"]')
    assert.equal(sel.length, 2, `Ctrl multi should be 2, got ${sel.length}`)
    // Shift range
    await React.act(async()=>{ btns()[0].dispatchEvent(new dom.window.MouseEvent('click',{bubbles:true})); await delay(30) })
    await React.act(async()=>{ btns()[2].dispatchEvent(new dom.window.MouseEvent('click',{bubbles:true, shiftKey:true})); await delay(30) })
    sel = container.querySelectorAll('[data-selected="true"]')
    assert.equal(sel.length, 3, `Shift range same slide should be 3, got ${sel.length}`)
  } finally { if(root) await React.act(async()=>{root.unmount()}); if(h.server) await h.server.close(); h.dom.window.close(); for(const [k,v] of h.prior){ if(v) Object.defineProperty(globalThis,k,v); else delete globalThis[k] } }
})

test('double-click with active label confirms exactly once, cross-slide rejected', async () => {
  const h = await makeHarness(emptyLearningState([
    imageCandidate('a',1,{itemIndex:0}), imageCandidate('b',1,{itemIndex:1}), imageCandidate('c',2),
  ]))
  let root
  try{
    const { dom, server, StoreProvider, Harness, createRoot, pywebview } = h
    root = createRoot(dom.window.document.getElementById('root'))
    await React.act(async()=>{ root.render(React.createElement(React.StrictMode, null, React.createElement(StoreProvider, null, React.createElement(Harness)))) })
    dom.window.pywebview = { api: pywebview }
    await React.act(async()=>{ dom.window.dispatchEvent(new dom.window.Event('pywebviewready')); await delay(60) })
    await waitFor(()=> dom.window.document.querySelector('[data-testid="bridge-state"]')?.textContent==='ready', 'bridge ready')
    await React.act(async()=>{ dom.window.document.querySelector('[data-testid="refresh-learning"]').dispatchEvent(new dom.window.MouseEvent('click',{bubbles:true})); await delay(30) })
    await delay(100)
    const container = dom.window.document.getElementById('root')
    const nav = [...container.querySelectorAll('aside nav button')].find(b=> b.textContent.trim().includes('Kiểm tra ảnh cải tiến'))
    await React.act(async()=>{ nav.dispatchEvent(new dom.window.MouseEvent('click',{bubbles:true})); await delay(10) })
    await waitFor(()=> container.querySelector('section[aria-label="Thông tin và nhãn xác nhận"]'), 'right panel')
    // active label
    const labelBtn = [...container.querySelectorAll('section[aria-label="Thông tin và nhãn xác nhận"] button')].find(b=> b.textContent.includes('Sau cải tiến'))
    assert.ok(labelBtn)
    await React.act(async()=>{ labelBtn.dispatchEvent(new dom.window.MouseEvent('click',{bubbles:true})); await delay(80) })
    const leftBtns = [...container.querySelectorAll('section[aria-label="Đối tượng ảnh"] button')]
    const before = h.getSaveCount()
    // double-click first item should save once
    await React.act(async()=>{ leftBtns[0].dispatchEvent(new dom.window.MouseEvent('dblclick',{bubbles:true})); await delay(350) })
    // allow async save: store calls save_learning_labels → getSaveCount bumps
    let tries=0; while(h.getSaveCount()===before && tries<20){ await delay(40); tries++}
    assert.equal(h.getSaveCount(), before+1, `double-click with label should save once, got ${h.getSaveCount()-before}`)
    // cross-slide multi: selecting across slides with Ctrl should not multi-hold
    await React.act(async()=>{ leftBtns[0].dispatchEvent(new dom.window.MouseEvent('click',{bubbles:true})); await delay(30) })
    await React.act(async()=>{ leftBtns[2].dispatchEvent(new dom.window.MouseEvent('click',{bubbles:true, ctrlKey:true})); await delay(50) })
    const sel = container.querySelectorAll('[data-selected="true"]')
    // cross-slide: our implementation collapses to single latest when slide differs
    assert.equal(sel.length, 1, `cross-slide must be 1, got ${sel.length}`)
  } finally { if(root) await React.act(async()=>{root.unmount()}); if(h.server) await h.server.close(); h.dom.window.close(); for(const [k,v] of h.prior){ if(v) Object.defineProperty(globalThis,k,v); else delete globalThis[k] } }
})
