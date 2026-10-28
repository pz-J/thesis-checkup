import { useState } from 'react'
import type { Finding } from '../lib/rules'
import { SEVERITY_LABEL, uniqueActuals } from '../lib/view'
import { AlertIcon, ChevronIcon, InfoIcon, LocateIcon } from './Icons'

interface Props {
  finding: Finding
  onLocate: (paragraphIndex?: number) => void
  onFixRule: (ruleId: string) => void
}

export function FindingCard({ finding, onLocate, onFixRule }: Props) {
  const [open, setOpen] = useState(false)
  const first = finding.occurrences[0]
  const actuals = uniqueActuals(finding)
  const shown = open ? finding.occurrences : finding.occurrences.slice(0, 3)
  const hiddenCount = finding.occurrences.length - shown.length

  return (
    <article className={`finding-card sev-${finding.severity}`}>
      <div className="finding-head">
        <span className={`sev-dot sev-${finding.severity}`} />
        <span className={`sev-tag sev-${finding.severity}`}>{SEVERITY_LABEL[finding.severity]}</span>
        <h3>{finding.title}</h3>
        <span className="finding-count">{finding.occurrences.length} 处</span>
        {finding.fixable && <span className="fix-badge">可自动修复</span>}
        <span className="rule-id">{finding.ruleId}</span>
      </div>

      <p className="finding-hint">{finding.hint}</p>

      {actuals.length > 0 && (
        <div className="actual-line">
          <span>实际</span>
          {actuals.slice(0, 3).map((a) => (
            <b key={a}>{a}</b>
          ))}
          {actuals.length > 3 && <em>等 {actuals.length} 种</em>}
        </div>
      )}

      <div className="occ-list">
        {shown.map((o, i) => (
          <div className="occ-row" key={`${finding.ruleId}-${i}`}>
            <span className="occ-loc">
              {o.paragraphIndex !== undefined
                ? `第 ${o.paragraphIndex + 1} 段`
                : o.sectionIndex !== undefined
                  ? `第 ${o.sectionIndex + 1} 节`
                  : '全文'}
            </span>
            <div className="occ-body">
              {o.text && <p className="occ-text">{o.text}</p>}
              {o.actual && <p className="occ-actual">{o.actual}</p>}
              {o.detail && <p className="occ-detail">{o.detail}</p>}
            </div>
            {o.paragraphIndex !== undefined && (
              <button className="icon-button" title="在文档中定位" onClick={() => onLocate(o.paragraphIndex)}>
                <LocateIcon width={14} height={14} />
              </button>
            )}
          </div>
        ))}
      </div>

      <div className="finding-actions">
        {hiddenCount > 0 && (
          <button className="text-button small" onClick={() => setOpen(true)}>
            <ChevronIcon width={13} height={13} />
            展开其余 {hiddenCount} 处
          </button>
        )}
        {open && finding.occurrences.length > 3 && (
          <button className="text-button small" onClick={() => setOpen(false)}>
            收起
          </button>
        )}
        {first?.paragraphIndex !== undefined && (
          <button className="chip-button small" onClick={() => onLocate(first.paragraphIndex)}>
            定位第一处
          </button>
        )}
        {finding.fixable && (
          <button className="chip-button small accent" onClick={() => onFixRule(finding.ruleId)}>
            修复此类
          </button>
        )}
      </div>
    </article>
  )
}
