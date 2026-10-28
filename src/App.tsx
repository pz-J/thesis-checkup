import { useEffect, useMemo, useRef, useState } from 'react'
import { parseDocx, type ParsedDoc } from './lib/docx'
import { applyFixes, downloadBlob, fixedFileName, type FixResult } from './lib/fix'
import { downloadReport } from './lib/exportReport'
import { downloadShareCard } from './lib/shareCard'
import { GENERIC_PROFILE, PROFILES, TJU_PROFILE, parseProfile, type RuleProfile } from './lib/profiles'
import { RULE_LIST, evaluate, type CategoryId, type Report } from './lib/rules'
import { Header } from './components/Header'
import { Landing } from './components/Landing'
import { ScorePanel } from './components/ScorePanel'
import { FindingCard } from './components/FindingCard'
import { DocumentDrawer } from './components/DocumentDrawer'
import { FixDialog } from './components/FixDialog'
import { FixResultModal } from './components/FixResultModal'
import { RuleModal } from './components/RuleModal'
import { RuleEditor } from './components/RuleEditor'
import { AlertIcon, CheckIcon, CloseIcon } from './components/Icons'

type View = 'landing' | 'report'
type Filter = 'all' | 'error' | 'warn' | 'info' | 'fixable'

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
const wait = (ms: number) => new Promise((r) => window.setTimeout(r, ms))

interface FixState {
  before: Report
  after: Report
  result: FixResult
  afterDoc: ParsedDoc
}

export default function App() {
  const [profileId, setProfileId] = useState(() => localStorage.getItem('thesis-checkup-profile') ?? TJU_PROFILE.id)
  const [customProfiles, setCustomProfiles] = useState<RuleProfile[]>(() => {
    try {
      const raw = localStorage.getItem('thesis-checkup-custom-profiles')
      if (!raw) return []
      const list = JSON.parse(raw)
      if (!Array.isArray(list)) return []
      return list
        .map((item) => parseProfile(item).profile)
        .filter((item): item is RuleProfile => !!item)
    } catch {
      return []
    }
  })
  const allProfiles = useMemo(() => [...PROFILES, ...customProfiles], [customProfiles])
  const profile = allProfiles.find((p) => p.id === profileId) ?? TJU_PROFILE
  const [view, setView] = useState<View>('landing')
  const [doc, setDoc] = useState<ParsedDoc | null>(null)
  const [report, setReport] = useState<Report | null>(null)
  const [busy, setBusy] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [filter, setFilter] = useState<Filter>('all')
  const [category, setCategory] = useState<CategoryId | null>(null)
  const [drawer, setDrawer] = useState<{ open: boolean; focus?: number }>({ open: false })
  const [fixOpen, setFixOpen] = useState(false)
  const [fixInitial, setFixInitial] = useState<string[] | null>(null)
  const [applying, setApplying] = useState(false)
  const [fixState, setFixState] = useState<FixState | null>(null)
  const [rulesOpen, setRulesOpen] = useState(false)
  const [editorOpen, setEditorOpen] = useState(false)
  const [toast, setToast] = useState<string | null>(null)
  const toastTimer = useRef<number | null>(null)

  useEffect(() => {
    localStorage.setItem('thesis-checkup-profile', profileId)
  }, [profileId])

  useEffect(() => {
    localStorage.setItem('thesis-checkup-custom-profiles', JSON.stringify(customProfiles))
  }, [customProfiles])

  useEffect(() => {
    if (!toast) return
    if (toastTimer.current) window.clearTimeout(toastTimer.current)
    toastTimer.current = window.setTimeout(() => setToast(null), 4200)
    return () => {
      if (toastTimer.current) window.clearTimeout(toastTimer.current)
    }
  }, [toast])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return
      if (editorOpen) {
        setEditorOpen(false)
        return
      }
      if (rulesOpen) {
        setRulesOpen(false)
        return
      }
      if (fixOpen) {
        setFixOpen(false)
        return
      }
      if (fixState) {
        setFixState(null)
        return
      }
      if (drawer.open) setDrawer({ open: false })
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [editorOpen, rulesOpen, fixOpen, fixState, drawer.open])

  const showToast = (msg: string) => setToast(msg)

  async function process(buffer: ArrayBuffer, fileName: string) {
    setError(null)
    setBusy('正在读取 OOXML 结构…')
    try {
      await wait(320)
      const parsed = await parseDocx(buffer, fileName)
      setBusy('正在还原样式继承…')
      await wait(320)
      setBusy('正在逐项比对规则…')
      const nextReport = evaluate(parsed, profile)
      await wait(220)
      setDoc(parsed)
      setReport(nextReport)
      setView('report')
      setFilter('all')
      setCategory(null)
      setDrawer({ open: false })
      setFixState(null)
    } catch (e) {
      setError(e instanceof Error ? e.message : '解析失败，请确认文件是有效的 .docx。')
    } finally {
      setBusy(null)
    }
  }

  async function loadFile(file: File) {
    if (!/\.docx$/i.test(file.name)) {
      setError('只支持 .docx 文件。旧版 .doc 请先在 Word 中「另存为」.docx 后再试。')
      return
    }
    if (file.size > 80 * 1024 * 1024) {
      setError('文件超过 80 MB，浏览器可能无法稳定解析。')
      return
    }
    const buffer = await file.arrayBuffer()
    await process(buffer, file.name)
  }

  async function loadSample(kind: 'issues' | 'clean' = 'issues') {
    try {
      setBusy(kind === 'clean' ? '正在加载满分样例…' : '正在加载示例论文…')
      const res = await fetch(kind === 'clean' ? './sample-clean.docx' : './sample-thesis.docx')
      if (!res.ok) throw new Error('示例文件加载失败')
      const buffer = await res.arrayBuffer()
      await process(buffer, kind === 'clean' ? '满分样例_已合规.docx' : '示例论文_待体检.docx')
    } catch (e) {
      setBusy(null)
      setError(e instanceof Error ? e.message : '示例文件加载失败')
    }
  }

  function reset() {
    setView('landing')
    setDoc(null)
    setReport(null)
    setError(null)
    setDrawer({ open: false })
    setFixState(null)
    setFilter('all')
    setCategory(null)
  }

  function changeProfile(id: string) {
    setProfileId(id)
    if (doc && view === 'report') {
      const next = allProfiles.find((p) => p.id === id) ?? TJU_PROFILE
      const nextReport = evaluate(doc, next)
      setReport(nextReport)
      setFilter('all')
      setCategory(null)
      showToast(`已切换到「${next.shortName}」规则包`)
    }
  }

  function importProfile(next: RuleProfile) {
    setCustomProfiles((prev) => [...prev.filter((p) => p.id !== next.id), next])
    setProfileId(next.id)
    if (doc && view === 'report') {
      setReport(evaluate(doc, next))
      setFilter('all')
      setCategory(null)
    }
    setEditorOpen(false)
    setRulesOpen(false)
    showToast(`已应用规则包「${next.name}」`)
  }

  function deleteProfile() {
    if (!customProfiles.some((p) => p.id === profileId)) return
    setCustomProfiles((prev) => prev.filter((p) => p.id !== profileId))
    setProfileId(TJU_PROFILE.id)
    if (doc && view === 'report') {
      setReport(evaluate(doc, TJU_PROFILE))
      setFilter('all')
      setCategory(null)
    }
    setEditorOpen(false)
    showToast('已删除自定义规则包，并切换回天大本科规则')
  }

  function openFix(ruleId?: string) {
    setFixInitial(ruleId ? [ruleId] : null)
    setFixOpen(true)
  }

  async function applySelected(rules: string[]) {
    if (!doc || !report) return
    setApplying(true)
    try {
      const result = await applyFixes(doc.buffer, doc, profile, rules)
      const afterDoc = await parseDocx(result.buffer, doc.fileName)
      const after = evaluate(afterDoc, profile)
      setFixState({ before: report, after, result, afterDoc })
      setFixOpen(false)
      showToast(`修复版已生成：${report.score} → ${after.score} 分`)
    } catch (e) {
      showToast(e instanceof Error ? `修复失败：${e.message}` : '修复失败')
    } finally {
      setApplying(false)
    }
  }

  function downloadFixed() {
    if (!fixState || !doc) return
    const blob = new Blob([fixState.result.buffer], { type: DOCX_MIME })
    downloadBlob(blob, fixedFileName(doc.fileName))
    showToast('已下载修复版。原文件没有被修改。')
  }

  function viewFixed() {
    if (!fixState) return
    setDoc(fixState.afterDoc)
    setReport(fixState.after)
    setFixState(null)
    setFilter('all')
    setCategory(null)
    showToast('已切换到修复后的复检报告')
  }

  const filtered = useMemo(() => {
    if (!report) return []
    let list = report.findings
    if (filter === 'fixable') list = list.filter((f) => f.fixable)
    else if (filter !== 'all') list = list.filter((f) => f.severity === filter)
    if (category) list = list.filter((f) => f.category === category)
    return list
  }, [report, filter, category])

  const tabCounts = useMemo(() => {
    if (!report) return { all: 0, error: 0, warn: 0, info: 0, fixable: 0 }
    return {
      all: report.findings.length,
      error: report.findings.filter((f) => f.severity === 'error').length,
      warn: report.findings.filter((f) => f.severity === 'warn').length,
      info: report.findings.filter((f) => f.severity === 'info').length,
      fixable: report.findings.filter((f) => f.fixable).length,
    }
  }, [report])

  return (
    <div className="app">
      <Header
        profiles={allProfiles}
        profileId={profile.id}
        onProfileChange={changeProfile}
        onAddProfile={() => setEditorOpen(true)}
        onRules={() => setRulesOpen(true)}
      />

      {view === 'landing' || !report || !doc ? (
        <Landing
          profile={profile}
          ruleCount={RULE_LIST.length}
          busy={busy}
          error={error}
          onFile={loadFile}
          onSample={loadSample}
          onEditRules={() => setEditorOpen(true)}
        />
      ) : (
        <main className="report">
          <div className="report-head">
            <div>
              <span className="eyebrow">CHECKUP REPORT · {new Date(report.generatedAt).toLocaleDateString('zh-CN')}</span>
              <h1>
                体检报告
                {fixState && <span className="report-badge">已生成修复版</span>}
              </h1>
              <p>
                {profile.name} · 共 {report.totalRules} 项规则，通过 {report.passedRules} 项
              </p>
            </div>
            <div className="report-head-actions">
              {category && (
                <button className="chip-button active" onClick={() => setCategory(null)}>
                  分类：{category === 'page' ? '页面' : category === 'body' ? '正文' : category === 'heading' ? '标题' : category === 'caption' ? '题注' : category === 'reference' ? '文献' : '结构'}
                  <CloseIcon width={12} height={12} />
                </button>
              )}
              <button className="chip-button" onClick={reset}>
                换一份论文
              </button>
            </div>
          </div>

          <div className="report-grid">
            <ScorePanel
              report={report}
              profile={profile}
              activeCategory={category}
              onCategory={setCategory}
              onFix={() => openFix()}
              onExport={() => {
                downloadReport(report, profile)
                showToast('体检报告已导出为 HTML，可直接打印或分享')
              }}
              onShare={() => {
                downloadShareCard(report, profile)
                showToast('体检卡已生成，可保存后分享')
              }}
              onOpenDoc={(index) => setDrawer({ open: true, focus: index })}
            />

            <section className="findings-panel">
              <div className="findings-toolbar">
                <div className="filter-tabs" role="tablist">
                  {(
                    [
                      ['all', '全部', tabCounts.all],
                      ['error', '必改', tabCounts.error],
                      ['warn', '建议', tabCounts.warn],
                      ['info', '提示', tabCounts.info],
                      ['fixable', '可修复', tabCounts.fixable],
                    ] as const
                  ).map(([key, label, count]) => (
                    <button
                      key={key}
                      className={`filter-tab ${filter === key ? 'active' : ''} ${key !== 'all' && key !== 'fixable' ? `tab-${key}` : ''}`}
                      onClick={() => setFilter(key)}
                      role="tab"
                      aria-selected={filter === key}
                    >
                      {label}
                      <span>{count}</span>
                    </button>
                  ))}
                </div>
                <div className="findings-meta">按严重程度排序 · 点击「定位」查看段落</div>
              </div>

              {filtered.length === 0 ? (
                <div className="empty-state">
                  <span className="empty-mark">
                    <CheckIcon width={22} height={22} />
                  </span>
                  <h3>这个筛选下没有问题</h3>
                  <p>换个筛选条件，或看看其他分类。</p>
                  <button
                    className="chip-button"
                    onClick={() => {
                      setFilter('all')
                      setCategory(null)
                    }}
                  >
                    查看全部问题
                  </button>
                </div>
              ) : (
                <div className="findings-list">
                  {filtered.map((f) => (
                    <FindingCard
                      key={f.ruleId}
                      finding={f}
                      onLocate={(index) => setDrawer({ open: true, focus: index })}
                      onFixRule={openFix}
                    />
                  ))}
                </div>
              )}

              <div className="findings-foot">
                <AlertIcon width={14} height={14} />
                报告只用于提交前自查，不能替代学院正式审查。规则包可在顶部切换，也可以按学院模板扩展。
              </div>
            </section>
          </div>
        </main>
      )}

      <footer className="site-footer">
        <div>
          <strong>论文体检</strong> v0.3 · 本地运行，文件不上传
        </div>
        <div>
          {RULE_LIST.length} 项规则 · {GENERIC_PROFILE.shortName} / {TJU_PROFILE.shortName} · 结果仅供自查
        </div>
      </footer>

      {drawer.open && doc && report && (
        <DocumentDrawer
          doc={doc}
          report={report}
          focusIndex={drawer.focus}
          onClose={() => setDrawer({ open: false })}
        />
      )}

      {fixOpen && report && (
        <FixDialog
          report={report}
          initialSelected={fixInitial}
          applying={applying}
          onClose={() => setFixOpen(false)}
          onApply={applySelected}
        />
      )}

      {fixState && (
        <FixResultModal
          before={fixState.before}
          after={fixState.after}
          applied={fixState.result.applied}
          onDownload={downloadFixed}
          onViewAfter={viewFixed}
          onClose={() => setFixState(null)}
        />
      )}

      {rulesOpen && (
        <RuleModal
          profile={profile}
          onClose={() => setRulesOpen(false)}
          onEdit={() => {
            setRulesOpen(false)
            setEditorOpen(true)
          }}
        />
      )}

      {editorOpen && (
        <RuleEditor
          profile={profile}
          custom={customProfiles.some((p) => p.id === profile.id)}
          onClose={() => setEditorOpen(false)}
          onImport={importProfile}
          onDelete={customProfiles.some((p) => p.id === profile.id) ? deleteProfile : undefined}
        />
      )}

      {toast && (
        <div className="toast" role="status">
          <CheckIcon width={14} height={14} />
          {toast}
        </div>
      )}
    </div>
  )
}
