import type { Report } from '../lib/rules'
import { RULE_META } from '../lib/rules'
import { CheckIcon, CloseIcon, DownloadIcon } from './Icons'

interface Props {
  before: Report
  after: Report
  applied: string[]
  onDownload: () => void
  onViewAfter: () => void
  onClose: () => void
}

export function FixResultModal({ before, after, applied, onDownload, onViewAfter, onClose }: Props) {
  const delta = after.score - before.score
  const remaining = after.counts.error + after.counts.warn + after.counts.info

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal result-modal" role="dialog" aria-modal="true" aria-label="修复版已生成" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <div>
            <span className="modal-kicker">REPAIRED COPY</span>
            <h2>修复版已生成</h2>
            <p>原文件保持不变。下载的是一个新副本，建议在 Word 中打开复核。</p>
          </div>
          <button className="icon-button large" onClick={onClose} aria-label="关闭">
            <CloseIcon width={18} height={18} />
          </button>
        </div>

        <div className="result-score">
          <div className="result-side">
            <span>修复前</span>
            <strong>{before.score}</strong>
            <em>{before.grade.label}</em>
          </div>
          <div className="result-arrow">
            <span>+{delta}</span>
            <svg viewBox="0 0 80 24" width="80" height="24" aria-hidden="true">
              <path d="M2 12h68" stroke="currentColor" strokeWidth="1.2" strokeDasharray="3 4" />
              <path d="m64 6 8 6-8 6" fill="none" stroke="currentColor" strokeWidth="1.2" />
            </svg>
          </div>
          <div className="result-side after">
            <span>修复后</span>
            <strong>{after.score}</strong>
            <em>{after.grade.label}</em>
          </div>
        </div>

        <div className="applied-list">
          {applied.map((id) => (
            <span key={id} className="applied-chip">
              <CheckIcon width={12} height={12} />
              {RULE_META[id]?.title ?? id}
            </span>
          ))}
        </div>

        <div className="result-note">
          仍有 {remaining} 类问题需要手动处理，包括结构、文献编号和页码等。修复版只调整了可直接确定的格式项。
        </div>

        <div className="modal-actions spread">
          <button className="chip-button" onClick={onViewAfter}>
            查看修复后报告
          </button>
          <div className="modal-actions">
            <button className="chip-button" onClick={onClose}>
              留在这里
            </button>
            <button className="primary-button" onClick={onDownload}>
              <DownloadIcon width={16} height={16} />
              下载修复版 .docx
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}
