import type { Report, CategoryId } from '../lib/rules'
import type { RuleProfile } from '../lib/profiles'
import { CATEGORIES } from '../lib/rules'
import { formatNumber } from '../lib/docx'
import { buildParaIssueMap, SEVERITY_LABEL } from '../lib/view'
import { DownloadIcon, FileIcon, LocateIcon, ScanIcon } from './Icons'

interface Props {
  report: Report
  profile: RuleProfile
  activeCategory: CategoryId | null
  onCategory: (c: CategoryId | null) => void
  onFix: () => void
  onExport: () => void
  onShare: () => void
  onOpenDoc: (index?: number) => void
}

export function ScorePanel({ report, profile, activeCategory, onCategory, onFix, onExport, onShare, onOpenDoc }: Props) {
  const issueMap = buildParaIssueMap(report)
  const sizeKb = `${formatNumber(report.doc.fileSize / 1024)} KB`
  const circumference = 2 * Math.PI * 54
  const offset = circumference * (1 - report.score / 100)
  const tone = report.grade.tone
  const fontCount = report.doc.stats.distinctBodyFonts.length
  const sizeCount = report.doc.stats.distinctBodySizes.length
  const spacingCount = report.doc.stats.distinctBodySpacings.length
  const visibleParas = report.doc.paragraphs.filter((p) => !p.empty || p.hasDrawing)

  return (
    <aside className="score-panel">
      <div className="file-card">
        <FileIcon width={18} height={18} />
        <div>
          <strong title={report.doc.fileName}>{report.doc.fileName}</strong>
          <span>
            {report.doc.stats.paragraphCount} 段 · {report.doc.sections.length} 节 · {sizeKb}
          </span>
        </div>
      </div>

      <div className={`score-card tone-${tone}`}>
        <div className="score-top">
          <div className="ring">
            <svg viewBox="0 0 120 120" aria-hidden="true">
              <circle className="ring-track" cx="60" cy="60" r="54" />
              <circle
                className="ring-value"
                cx="60"
                cy="60"
                r="54"
                strokeDasharray={circumference}
                strokeDashoffset={offset}
              />
            </svg>
            <div className="ring-label">
              <strong>{report.score}</strong>
              <span>/ 100</span>
            </div>
          </div>
          <div className="score-meta">
            <span className="grade-stamp">{report.grade.label}</span>
            <p>{report.grade.description}</p>
          </div>
        </div>
        <div className="score-counts">
          <div>
            <b className="sev-error">{report.counts.error}</b>
            <span>必改</span>
          </div>
          <div>
            <b className="sev-warn">{report.counts.warn}</b>
            <span>建议</span>
          </div>
          <div>
            <b className="sev-info">{report.counts.info}</b>
            <span>提示</span>
          </div>
          <div>
            <b className="sev-good">{report.counts.fixable}</b>
            <span>可修复</span>
          </div>
        </div>
        <div className="score-foot">
          通过 {report.passedRules} / {report.totalRules} 项规则 · 共 {report.counts.occurrences} 处
        </div>
      </div>

      <section className="panel-block">
        <div className="panel-title">
          <h3>分类得分</h3>
          <span>{activeCategory ? '已筛选' : '点击筛选'}</span>
        </div>
        <div className="cat-list">
          {CATEGORIES.map((c) => {
            const score = report.categoryScores[c.id]
            const active = activeCategory === c.id
            return (
              <button
                key={c.id}
                className={`cat-row ${active ? 'active' : ''}`}
                onClick={() => onCategory(active ? null : c.id)}
              >
                <span className="cat-name">{c.name}</span>
                <span className="cat-bar">
                  <i style={{ width: `${score}%` }} />
                </span>
                <span className="cat-score">{score}</span>
              </button>
            )
          })}
        </div>
      </section>

      <section className="panel-block">
        <div className="panel-title">
          <h3>格式指纹</h3>
          <span>{report.doc.stats.bodyParagraphCount} 个正文段</span>
        </div>
        <div className="fp-table">
          <FingerprintRow label="字体" items={report.doc.stats.distinctBodyFonts} noisy={fontCount > 1} />
          <FingerprintRow label="字号" items={report.doc.stats.distinctBodySizes} noisy={sizeCount > 1} />
          <FingerprintRow label="行距" items={report.doc.stats.distinctBodySpacings} noisy={spacingCount > 1} />
        </div>
        {(fontCount > 1 || sizeCount > 1 || spacingCount > 1) && (
          <p className="fp-note">
            正文里出现了 {fontCount} 种字体、{sizeCount} 种字号、{spacingCount} 种行距。肉眼看不出来，但 Word 知道。
          </p>
        )}
      </section>

      <section className="panel-block">
        <div className="panel-title">
          <h3>格式 X 光</h3>
          <span>点击定位</span>
        </div>
        <div className="xray" aria-label="段落格式问题分布">
          {visibleParas.map((p) => {
            const issue = issueMap.get(p.index)
            const sev = issue?.severity
            const base = sev === 'error' ? 46 : sev === 'warn' ? 32 : sev === 'info' ? 22 : 10
            const bump = (p.headingLevel ? 10 : 0) + Math.min(6, Math.round(p.charCount / 60))
            return (
              <button
                key={p.index}
                className={`xray-bar ${sev ? `sev-${sev}` : 'pass'} ${p.headingLevel ? 'heading' : ''}`}
                style={{ height: `${base + bump}px` }}
                title={`第 ${p.index + 1} 段${issue ? ` · ${issue.titles.join(' / ')}` : ' · 通过'}`}
                onClick={() => onOpenDoc(p.index)}
              />
            )
          })}
        </div>
        <div className="xray-legend">
          <span className="sev-error">必改</span>
          <span className="sev-warn">建议</span>
          <span className="sev-info">提示</span>
          <span>通过</span>
        </div>
      </section>

      <div className="panel-actions">
        <button className="primary-button full" onClick={onFix} disabled={report.counts.fixable === 0}>
          <DownloadIcon width={16} height={16} />
          生成修复版
          <span className="button-badge">{report.counts.fixable}</span>
        </button>
        <div className="action-row">
          <button className="chip-button" onClick={onExport}>
            导出体检报告
          </button>
          <button className="chip-button" onClick={onShare}>
            生成体检卡
          </button>
        </div>
        <button className="doc-link" onClick={() => onOpenDoc()}>
          <ScanIcon width={14} height={14} />
          查看逐段 X 光
          <LocateIcon width={13} height={13} />
        </button>
        <p className="action-note">
          修复只调整字体、字号、行距、缩进与页边距等确定项，不改文字内容；生成的是副本，原文件不会被修改。新增文字仍可能继承原样式。
        </p>
      </div>
    </aside>
  )
}

function FingerprintRow({
  label,
  items,
  noisy,
}: {
  label: string
  items: { value: string; count: number }[]
  noisy: boolean
}) {
  return (
    <div className={`fp-row ${noisy ? 'noisy' : ''}`}>
      <span className="fp-label">{label}</span>
      <div className="fp-values">
        {items.length === 0 && <em>未检测到</em>}
        {items.slice(0, 4).map((it) => (
          <span className="fp-chip" key={it.value}>
            {it.value}
            <i>{it.count}</i>
          </span>
        ))}
        {items.length > 4 && <span className="fp-chip more">+{items.length - 4}</span>}
      </div>
    </div>
  )
}
