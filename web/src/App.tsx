import { useEffect, useMemo, useState } from 'react'
import { FormulaEngine, parseCellReference } from '@ricsam/formula-engine'
import { FormulaWorkbook } from '@ricsam/react-spreadsheets'
import '@ricsam/react-spreadsheets/styles.css'
import './App.css'

const DOCS_URL = 'https://formula-engine.mintlify.site'
const GITHUB_URL = 'https://github.com/ricsam/formula-engine'
const GRID_URL = 'https://github.com/ricsam/react-spreadsheets'
const WORKBOOK = 'Studio'

const sheet = (sheetName: string) => ({ workbookName: WORKBOOK, sheetName })

function finiteRange(startCol: number, startRow: number, endCol: number, endRow: number) {
  return {
    start: { col: startCol, row: startRow },
    end: {
      col: { type: 'number' as const, value: endCol },
      row: { type: 'number' as const, value: endRow },
    },
  }
}

function populateEngine(engine: FormulaEngine) {
  engine.setSheetContent(
    sheet('Revenue Model'),
    new Map<string, string | number | boolean | undefined>([
      ['A1', 'PRODUCT'], ['B1', 'UNITS'], ['C1', 'UNIT PRICE'], ['D1', 'REVENUE'], ['E1', 'MIX'],
      ['A2', 'Starter'], ['B2', 120], ['C2', 29], ['D2', '=B2*C2'], ['E2', '=D2/$D$6'],
      ['A3', 'Pro'], ['B3', 65], ['C3', 99], ['D3', '=B3*C3'], ['E3', '=D3/$D$6'],
      ['A4', 'Team'], ['B4', 28], ['C4', 249], ['D4', '=B4*C4'], ['E4', '=D4/$D$6'],
      ['A5', 'Enterprise'], ['B5', 8], ['C5', 1200], ['D5', '=B5*C5'], ['E5', '=D5/$D$6'],
      ['A6', 'Total'], ['B6', '=SUM(B2:B5)'], ['D6', '=SUM(D2:D5)'], ['E6', '=SUM(E2:E5)'],
      ['A8', 'GROWTH SCENARIO'], ['B8', 'Multiplier'], ['C8', '=Inputs!B2'], ['D8', 'Next period'], ['E8', '=D6*C8'],
      ['A10', '4-PERIOD FORECAST'], ['B10', '=SEQUENCE(4,1,1,1)'], ['C10', '=SEQUENCE(4,1,D6,D6*0.15)'],
    ]),
  )

  engine.setSheetContent(
    sheet('Inputs'),
    new Map<string, string | number | boolean | undefined>([
      ['A1', 'ASSUMPTION'], ['B1', 'VALUE'], ['C1', 'NOTES'],
      ['A2', 'Growth multiplier'], ['B2', 1.2], ['C2', 'Edit this and return to Revenue Model'],
      ['A3', 'Target margin'], ['B3', 0.72], ['C3', 'Workbook-scoped input'],
      ['A5', 'Named expression'], ['B5', '=B3*100'], ['C5', 'Live dependent formula'],
    ]),
  )

  engine.setSheetContent(
    sheet('Lookup Lab'),
    new Map<string, string | number | boolean | undefined>([
      ['A1', 'REGION'], ['B1', 'PIPELINE'], ['C1', 'OWNER'],
      ['A2', 'North'], ['B2', 125000], ['C2', 'Amelia'],
      ['A3', 'South'], ['B3', 98000], ['C3', 'Malik'],
      ['A4', 'East'], ['B4', 142000], ['C4', 'Noor'],
      ['A5', 'West'], ['B5', 117000], ['C5', 'Theo'],
      ['E1', 'Find region'], ['F1', 'East'],
      ['E2', 'Pipeline'], ['F2', '=XLOOKUP(F1,A2:A5,B2:B5,"Not found")'],
      ['E3', 'Owner'], ['F3', '=XLOOKUP(F1,A2:A5,C2:C5,"Not found")'],
      ['E5', 'North + West'], ['F5', '=SUMIF(A2:A5,"North",B2:B5)+SUMIF(A2:A5,"West",B2:B5)'],
    ]),
  )
}

function createDemoEngine() {
  const engine = FormulaEngine.buildEmpty()
  engine.addWorkbook(WORKBOOK)
  engine.addSheet(sheet('Revenue Model'))
  engine.addSheet(sheet('Inputs'))
  engine.addSheet(sheet('Lookup Lab'))

  engine.addNamedExpression({ expressionName: 'ModelName', expression: '"SaaS Plan"', workbookName: WORKBOOK })
  populateEngine(engine)

  const styleArea = (sheetName: string, startCol: number, startRow: number, endCol: number, endRow: number) => ({
    workbookName: WORKBOOK,
    sheetName,
    range: finiteRange(startCol, startRow, endCol, endRow),
  })

  engine.addCellStyle({
    areas: [styleArea('Revenue Model', 0, 0, 4, 0), styleArea('Inputs', 0, 0, 2, 0), styleArea('Lookup Lab', 0, 0, 2, 0)],
    style: { bold: true, backgroundColor: '#eef2ff', color: '#3730a3' },
  })
  engine.addCellStyle({
    areas: [styleArea('Revenue Model', 0, 5, 4, 5)],
    style: { bold: true, backgroundColor: '#f1f5f9', borderColor: '#94a3b8', borderSides: { top: true } },
  })
  engine.addCellStyle({
    areas: [styleArea('Revenue Model', 0, 7, 4, 7), styleArea('Revenue Model', 0, 9, 2, 9)],
    style: { bold: true, backgroundColor: '#ecfdf5', color: '#047857' },
  })
  engine.clearUndoRedoHistory()
  return engine
}

function Icon({ name, size = 18 }: { name: 'arrow' | 'github' | 'spark' | 'code' | 'grid' | 'bolt' | 'check' | 'copy' | 'undo'; size?: number }) {
  const paths: Record<typeof name, React.ReactNode> = {
    arrow: <><path d="M5 12h14"/><path d="m13 6 6 6-6 6"/></>,
    github: <path d="M15 22v-4a4.8 4.8 0 0 0-1-3.5c3.3-.4 6.8-1.6 6.8-7A5.5 5.5 0 0 0 19.3 4 5.1 5.1 0 0 0 19.2.5S18 0 15 2a13.4 13.4 0 0 0-7 0C5 .1 3.8.5 3.8.5A5.1 5.1 0 0 0 3.7 4a5.5 5.5 0 0 0-1.5 3.8c0 5.4 3.5 6.6 6.8 7A4.8 4.8 0 0 0 8 18v4"/>,
    spark: <><path d="m12 3-1.3 4.2a5 5 0 0 1-3.3 3.3L3 12l4.4 1.5a5 5 0 0 1 3.3 3.3L12 21l1.3-4.2a5 5 0 0 1 3.3-3.3L21 12l-4.4-1.5a5 5 0 0 1-3.3-3.3L12 3Z"/></>,
    code: <><path d="m8 9-4 3 4 3"/><path d="m16 9 4 3-4 3"/><path d="m14 5-4 14"/></>,
    grid: <><rect x="3" y="3" width="18" height="18" rx="1"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18"/></>,
    bolt: <path d="m13 2-9 12h8l-1 8 9-12h-8l1-8Z"/>,
    check: <path d="m5 12 4 4L19 6"/>,
    copy: <><rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></>,
    undo: <><path d="M9 14 4 9l5-5"/><path d="M4 9h10a6 6 0 0 1 6 6v1"/></>,
  }
  return <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{paths[name]}</svg>
}

function formatNumber(value: unknown) {
  return typeof value === 'number' ? new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 }).format(value) : String(value ?? '')
}

function LiveWorkbook() {
  const engine = useMemo(createDemoEngine, [])
  const [showFormulas, setShowFormulas] = useState(false)
  const [activeSheet, setActiveSheet] = useState('Revenue Model')
  const [revision, setRevision] = useState(0)

  useEffect(() => engine.onUpdate(() => setRevision((value) => value + 1)), [engine])

  const read = (ref: string, sheetName = 'Revenue Model') => {
    const { colIndex, rowIndex } = parseCellReference(ref)
    return engine.getCellValue({ workbookName: WORKBOOK, sheetName, colIndex, rowIndex })
  }

  const reset = () => {
    populateEngine(engine)
    engine.clearUndoRedoHistory()
    setActiveSheet('Revenue Model')
  }

  const totalRevenue = read('D6')
  const nextRevenue = read('E8')
  void revision

  return (
    <section className="playground-section" id="playground">
      <div className="section-heading split-heading">
        <div>
          <div className="eyebrow"><span>Live playground</span></div>
          <h2>This is the product, not a screenshot.</h2>
        </div>
        <p>Edit a price, drag a fill handle, add a sheet, paste a block. The grid runs the real engine, selection manager and React component package in your browser.</p>
      </div>

      <div className="workbench-shell">
        <div className="workbench-topbar">
          <div className="traffic"><span/><span/><span/></div>
          <div className="workbench-title"><span className="status-dot"/>Studio / Revenue forecast</div>
          <div className="workbench-actions">
            <button type="button" onClick={() => engine.undo()} disabled={!engine.canUndo()}><Icon name="undo" size={15}/>Undo</button>
            <button type="button" onClick={() => setShowFormulas((value) => !value)} className={showFormulas ? 'active' : ''}><Icon name="code" size={15}/>{showFormulas ? 'Values' : 'Formulas'}</button>
            <button type="button" onClick={reset}>Reset</button>
          </div>
        </div>
        <div className="metric-strip">
          <div><span>Current revenue</span><strong>${formatNumber(totalRevenue)}</strong></div>
          <div><span>Next period</span><strong>${formatNumber(nextRevenue)}</strong></div>
          <div><span>Functions loaded</span><strong>39</strong></div>
          <div className="recalc"><span>Calculation</span><strong><i/> Live</strong></div>
        </div>
        <div className="workbook-stage">
          <FormulaWorkbook
            engine={engine}
            workbookName={WORKBOOK}
            activeSheet={activeSheet}
            onActiveSheetChange={setActiveSheet}
            showFormulas={showFormulas}
            isSelected
          />
        </div>
        <div className="workbench-hint"><span>Tip</span> Double-click any cell to edit. Try changing <code>Inputs!B2</code>, then return to the model.</div>
      </div>
    </section>
  )
}

const featurePanels = [
  {
    id: 'dependencies', label: 'Dependency graph', icon: 'bolt' as const,
    title: 'Recalculate only what changed.',
    copy: 'Edits patch the dependency graph. Values are evaluated on read, cached, and invalidation stops as soon as an output stays the same.',
  },
  {
    id: 'arrays', label: 'Dynamic arrays', icon: 'grid' as const,
    title: 'Arrays that know where to go.',
    copy: 'Range arithmetic and SEQUENCE spill across rows and columns. Every spilled value remains referenceable by ordinary formulas.',
  },
  {
    id: 'modeling', label: 'Modeling primitives', icon: 'spark' as const,
    title: 'Names and tables, built in.',
    copy: 'Workbook and sheet-scoped names, structured table references, cross-sheet formulas, search, replace, metadata, undo and redo.',
  },
]

function AnimatedCapabilities() {
  const [active, setActive] = useState(0)
  useEffect(() => {
    const timer = window.setInterval(() => setActive((value) => (value + 1) % featurePanels.length), 4200)
    return () => window.clearInterval(timer)
  }, [])
  const panel = featurePanels[active]!

  return (
    <section className="capabilities" id="capabilities">
      <div className="section-heading centered-heading">
        <div className="eyebrow"><span>Calculation infrastructure</span></div>
        <h2>A spreadsheet engine that behaves like an engine.</h2>
        <p>Not a bag of formula functions. A complete model runtime for products that need spreadsheet semantics without inheriting a spreadsheet UI.</p>
      </div>
      <div className="capability-layout">
        <div className="capability-tabs">
          {featurePanels.map((item, index) => (
            <button key={item.id} type="button" className={active === index ? 'active' : ''} onClick={() => setActive(index)}>
              <span className="cap-icon"><Icon name={item.icon}/></span>
              <span><strong>{item.label}</strong><small>{item.title}</small></span>
              <i className="progress"/>
            </button>
          ))}
        </div>
        <div className={`capability-visual visual-${panel.id}`}>
          {panel.id === 'dependencies' && (
            <div className="graph-demo" aria-hidden="true">
              <div className="graph-formula"><span>F12</span>=SUM(B2:B10)*TaxRate</div>
              <svg viewBox="0 0 620 280"><path d="M98 65 C190 65 175 140 282 140"/><path d="M98 210 C190 210 175 140 282 140"/><path d="M350 140 C440 140 420 82 530 82"/><path d="M350 140 C440 140 420 218 530 218"/></svg>
              <div className="graph-node node-a"><small>Range</small><strong>B2:B10</strong><em>9 values</em></div>
              <div className="graph-node node-b"><small>Name</small><strong>TaxRate</strong><em>0.25</em></div>
              <div className="graph-node node-c active"><small>Formula</small><strong>F12</strong><em>evaluating</em></div>
              <div className="graph-node node-d"><small>Value</small><strong>18,450</strong><em>cached</em></div>
              <div className="graph-node node-e"><small>Dependent</small><strong>Dashboard</strong><em>updated</em></div>
            </div>
          )}
          {panel.id === 'arrays' && (
            <div className="spill-demo" aria-hidden="true">
              <div className="formula-chip"><span>H2</span>=SEQUENCE(4,3,10,10)</div>
              <div className="spill-grid">
                {Array.from({ length: 12 }, (_, index) => <span key={index} style={{ '--i': index } as React.CSSProperties}>{(index + 1) * 10}</span>)}
              </div>
              <div className="spill-label">12 cells · 1 formula</div>
            </div>
          )}
          {panel.id === 'modeling' && (
            <div className="model-demo" aria-hidden="true">
              <div className="model-table">
                <div className="model-header"><span>Product</span><span>Qty</span><span>Revenue</span></div>
                <div><span>Starter</span><span>120</span><span>3,480</span></div>
                <div><span>Pro</span><span>65</span><span>6,435</span></div>
                <div><span>Team</span><span>28</span><span>6,972</span></div>
              </div>
              <div className="structured-ref"><small>Structured reference</small><code>=SUM(Sales[Revenue])</code><strong>16,887</strong></div>
              <div className="scope-pill">Workbook name · <strong>TaxRate</strong></div>
            </div>
          )}
          <div className="visual-copy"><span>0{active + 1}</span><div><h3>{panel.title}</h3><p>{panel.copy}</p></div></div>
        </div>
      </div>
    </section>
  )
}

function CopyInstall() {
  const [copied, setCopied] = useState(false)
  const command = 'bun add @ricsam/formula-engine'
  const copy = async () => {
    await navigator.clipboard.writeText(command)
    setCopied(true)
    window.setTimeout(() => setCopied(false), 1800)
  }
  return <button type="button" className="install-command" onClick={copy}><code>{command}</code><span><Icon name={copied ? 'check' : 'copy'} size={16}/>{copied ? 'Copied' : 'Copy'}</span></button>
}

function App() {
  return (
    <div className="site-shell">
      <header className="nav-wrap">
        <nav className="nav container">
          <a className="brand" href="#top" aria-label="FormulaEngine home"><span className="brand-mark">ƒx</span><span>Formula<span>Engine</span></span></a>
          <div className="nav-links"><a href="#capabilities">Capabilities</a><a href="#playground">Playground</a><a href={GRID_URL}>React grid</a><a href={DOCS_URL}>Docs</a></div>
          <a className="nav-github" href={GITHUB_URL}><Icon name="github" size={17}/><span>GitHub</span></a>
        </nav>
      </header>

      <main id="top">
        <section className="hero-section container">
          <div className="hero-grid">
            <div className="hero-copy">
              <div className="release-pill"><span>v0.2.15</span> Headless spreadsheet runtime <Icon name="arrow" size={14}/></div>
              <h1>The calculation layer for <em>serious spreadsheets.</em></h1>
              <p className="hero-lede">A TypeScript formula engine with lazy evaluation, dynamic arrays, structured references and undo-ready mutations. Bring your own UI—or use ours.</p>
              <div className="hero-actions"><a className="button primary" href="#playground">Try the live model <Icon name="arrow"/></a><a className="button secondary" href={DOCS_URL}>Read the docs</a></div>
              <CopyInstall />
              <div className="hero-proof"><span><Icon name="check" size={15}/>39 built-in functions</span><span><Icon name="check" size={15}/>1,985 tests</span><span><Icon name="check" size={15}/>Zero DOM dependencies</span></div>
            </div>
            <div className="hero-code-card" aria-label="FormulaEngine code example">
              <div className="code-card-bar"><div><span/><span/><span/></div><span>forecast.ts</span><small>TypeScript</small></div>
              <pre><code><span className="c-purple">import</span> {'{ '}<span className="c-blue">FormulaEngine</span>{' }'} <span className="c-purple">from</span>{' '}<span className="c-green">"@ricsam/formula-engine"</span>{'\n\n'}<span className="c-purple">const</span> engine = FormulaEngine.<span className="c-blue">buildEmpty</span>(){'\n'}engine.<span className="c-blue">addWorkbook</span>(<span className="c-green">"Forecast"</span>){'\n'}engine.<span className="c-blue">addSheet</span>({'{'} workbookName: <span className="c-green">"Forecast"</span>,{'\n  '}sheetName: <span className="c-green">"Model"</span> {'}'}){'\n\n'}engine.<span className="c-blue">setSheetContent</span>(sheet, <span className="c-purple">new</span> Map([{'\n  '}[<span className="c-green">"A1"</span>, <span className="c-orange">120</span>],{'\n  '}[<span className="c-green">"A2"</span>, <span className="c-orange">65</span>],{'\n  '}[<span className="c-green">"A3"</span>, <span className="c-green">"=SUM(A1:A2)"</span>],{'\n'}])){'\n\n'}engine.<span className="c-blue">getCellValue</span>(a3) <span className="c-muted">// 185</span></code></pre>
              <div className="code-result"><span className="status-dot"/>Evaluated lazily in <strong>&lt;1 ms</strong></div>
              <div className="code-orbit orbit-one">#SPILL!</div><div className="code-orbit orbit-two">Sales[Qty]</div><div className="code-orbit orbit-three">Sheet2!A1</div>
            </div>
          </div>
          <div className="hero-trust"><span>BUILT FOR</span><div>Financial models</div><i/> <div>Planning tools</div><i/> <div>Data grids</div><i/> <div>Workflow products</div></div>
        </section>

        <AnimatedCapabilities />
        <LiveWorkbook />

        <section className="architecture-section container">
          <div className="section-heading split-heading">
            <div><div className="eyebrow"><span>Composable by design</span></div><h2>Use the layer you need.</h2></div>
            <p>FormulaEngine stays headless. Selection stays independent. The React grid is optional. Adopt one package or the complete spreadsheet stack.</p>
          </div>
          <div className="package-stack">
            <a href={GITHUB_URL} className="package-card engine-card"><span className="package-icon"><Icon name="bolt"/></span><small>CORE RUNTIME</small><h3>@ricsam/formula-engine</h3><p>Parsing, evaluation, dependency tracking, spills, tables, names, search, copy, fill and history.</p><div><span>TypeScript</span><span>Headless</span><Icon name="arrow"/></div></a>
            <a href={GRID_URL} className="package-card grid-card"><span className="package-icon"><Icon name="grid"/></span><small>REACT SURFACE</small><h3>@ricsam/react-spreadsheets</h3><p>An infinitely scrollable grid with editing, resizing, clipboard, fill handles, workbooks and overlays.</p><div><span>React 18 / 19</span><span>CSS variables</span><Icon name="arrow"/></div></a>
            <a href="https://www.npmjs.com/package/@ricsam/selection-manager" className="package-card selection-card"><span className="package-icon"><Icon name="spark"/></span><small>INTERACTION MODEL</small><h3>@ricsam/selection-manager</h3><p>Framework-independent cell selection, keyboard navigation, copy, paste and multi-area selection semantics.</p><div><span>Independent</span><span>Composable</span><Icon name="arrow"/></div></a>
          </div>
        </section>

        <section className="feature-ledger container">
          <div className="ledger-row ledger-head"><span>What ships today</span><span>Not a roadmap. In the package.</span></div>
          {[
            ['Evaluation', 'Lazy, dependency-aware recalculation with cached results', 'bolt'],
            ['Modeling', 'Workbook + sheet scope, named expressions, structured tables', 'spark'],
            ['Arrays', '2D spills, range arithmetic, scalar broadcast, infinite ranges', 'grid'],
            ['Editing', 'Formula-aware copy, cut, paste, move, fill, search and replace', 'copy'],
            ['Application state', 'Styles, metadata, reference handles, transactions and history', 'code'],
          ].map(([title, copy, icon], index) => <div className="ledger-row" key={title}><span className="ledger-num">0{index + 1}</span><h3><span className="ledger-icon"><Icon name={icon as 'bolt'}/></span>{title}</h3><p>{copy}</p><Icon name="check"/></div>)}
        </section>

        <section className="cta-section container">
          <div className="cta-card">
            <div><div className="eyebrow light"><span>Ready to calculate?</span></div><h2>Start with a formula.<br/>Build the whole product.</h2><p>Open source, typed end to end, and small enough to understand.</p></div>
            <div className="cta-actions"><CopyInstall/><div><a className="button white" href={DOCS_URL}>Get started <Icon name="arrow"/></a><a className="button ghost" href={GITHUB_URL}><Icon name="github"/>View source</a></div></div>
          </div>
        </section>
      </main>

      <footer><div className="container footer-inner"><a className="brand" href="#top"><span className="brand-mark">ƒx</span><span>Formula<span>Engine</span></span></a><p>Headless spreadsheet infrastructure for TypeScript.</p><div><a href={DOCS_URL}>Docs</a><a href={GITHUB_URL}>GitHub</a><a href={GRID_URL}>React grid</a><a href="https://www.npmjs.com/package/@ricsam/formula-engine">npm</a></div></div></footer>
    </div>
  )
}

export default App
