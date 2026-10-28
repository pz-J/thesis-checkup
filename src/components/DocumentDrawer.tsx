import { useEffect, useState } from 'react'
import {
  alignmentText,
  displayFont,
  formatSize,
  indentText,
  lineSpacingText,
  majorityFont,
  majoritySize,
  type ParsedDoc,
} from '../lib/docx'
import type { Report } from '../lib/rules'
import { buildParaIssueMap, findingsForParagraph, SEVERITY_LABEL } from '../lib/view'
import { CloseIcon, FileIcon } from './Icons'

interface Props {
  doc: ParsedDoc
  report: Report
  focusIndex?: number
  onClose: () => void
}

export function DocumentDrawer({ doc, report, focusIndex, onClose }: Props) {
  const [expanded, setExpanded] = useState<number | null>(focusIndex ?? null)
  const [onlyIssues, setOnlyIssues] = useState(false)
  const issueMap = buildParaIssueMap(report)
  const rows = doc.paragraphs.filter((p) => {
    if (!onlyIssues) return true
    return issueMap.has(p.index)
  })

  useEffect(() => {
    setExpanded(focusIndex ?? null)
    if (focusIndex !== undefined) {
      const t = window.setTimeout(() => {
        document.getElementById(`doc-row-${focusIndex}`)?.scrollIntoView({ block: 'center', behavior: 'smooth' })
      }, 80)
      return () => window.clearTimeout(t)
    }
  }, [focusIndex])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKey)
    const prev = document.body.style.overflow
    document.body.style.overflow = 'hidden'
    return () => {
      window.removeEventListener('keydown', onKey)
      document.body.style.overflow = prev
    }
  }, [onClose])

  return (
    <div className="drawer-overlay" onClick={onClose}>
      <div className="drawer" role="dialog" aria-modal="true" aria-label="逐段 X 光" onClick={(e) => e.stopPropagation()}>
        <div className="drawer-head">
          <div>
            <span className="drawer-kicker">DOCUMENT X-RAY</span>
            <h2>逐段 X 光</h2>
            <p>
              <FileIcon width={13} height={13} />
              {doc.fileName}
            </p>
          </div>
          <div className="drawer-head-actions">
            <label className="switch">
              <input type="checkbox" checked={onlyIssues} onChange={(e) => setOnlyIssues(e.target.checked)} />
              <span>只看有问题</span>
            </label>
            <button className="icon-button large" onClick={onClose} aria-label="关闭">
              <CloseIcon width={18} height={18} />
            </button>
          </div>
        </div>

        <div className="drawer-summary">
          <span>{doc.stats.paragraphCount} 段</span>
          <span>{doc.stats.bodyParagraphCount} 个正文段</span>
          <span>{report.counts.occurrences} 处问题</span>
          <span>{doc.sections.length} 节</span>
        </div>

        <div className="doc-list">
          {rows.map((p) => {
            const issue = issueMap.get(p.index)
            const isOpen = expanded === p.index
            const findings = findingsForParagraph(report, p.index)
            return (
              <div
                key={p.index}
                id={`doc-row-${p.index}`}
                className={`doc-row ${issue ? `sev-${issue.severity}` : ''} ${isOpen ? 'open' : ''}`}
              >
                <button className="doc-row-main" onClick={() => setExpanded(isOpen ? null : p.index)}>
                  <span className="doc-index">{String(p.index + 1).padStart(2, '0')}</span>
                  <span className={`doc-text ${p.empty && !p.hasDrawing ? 'is-empty' : ''}`}>
                    {p.hasDrawing ? '［图片 / 图形］' : p.empty ? '空段' : p.text}
                  </span>
                  <span className="doc-badges">
                    {p.headingLevel && <span className="mini-badge">H{p.headingLevel}</span>}
                    {p.captionKind && <span className="mini-badge">{p.captionKind === 'figure' ? '图题' : '表题'}</span>}
                    {issue && <span className={`mini-badge sev-${issue.severity}`}>{findings.length}</span>}
                  </span>
                </button>
                {isOpen && (
                  <div className="doc-detail">
                    <dl className="format-grid">
                      <div>
                        <dt>样式</dt>
                        <dd>{p.styleName}</dd>
                      </div>
                      <div>
                        <dt>中文字体</dt>
                        <dd>{displayFont(majorityFont(p))}</dd>
                      </div>
                      <div>
                        <dt>字号</dt>
                        <dd>{formatSize(majoritySize(p))}</dd>
                      </div>
                      <div>
                        <dt>行距</dt>
                        <dd>{lineSpacingText(p.format)}</dd>
                      </div>
                      <div>
                        <dt>首行</dt>
                        <dd>{indentText(p.format)}</dd>
                      </div>
                      <div>
                        <dt>对齐</dt>
                        <dd>{alignmentText(p.format.alignment)}</dd>
                      </div>
                    </dl>
                    {findings.length > 0 && (
                      <div className="doc-issues">
                        {findings.map((f) => (
                          <div key={f.ruleId} className={`doc-issue sev-${f.severity}`}>
                            <span className={`sev-tag sev-${f.severity}`}>{SEVERITY_LABEL[f.severity]}</span>
                            <span>{f.title}</span>
                          </div>
                        ))}
                      </div>
                    )}
                    {findings.length === 0 && <div className="doc-pass">这一段没有发现规则问题</div>}
                  </div>
                )}
              </div>
            )
          })}
        </div>
      </div>
    </div>
  )
}
