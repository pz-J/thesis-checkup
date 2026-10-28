import { useMemo, useState } from 'react'
import { RULE_META, type Report } from '../lib/rules'
import { CloseIcon, DownloadIcon } from './Icons'

interface Props {
  report: Report
  initialSelected: string[] | null
  applying: boolean
  onClose: () => void
  onApply: (rules: string[]) => void
}

export function FixDialog({ report, initialSelected, applying, onClose, onApply }: Props) {
  const fixable = useMemo(() => report.findings.filter((f) => f.fixable), [report])
  const defaultFixable = useMemo(
    () => fixable.filter((f) => !RULE_META[f.ruleId]?.defaultOff),
    [fixable],
  )
  const [selected, setSelected] = useState<Set<string>>(
    () => new Set(initialSelected ?? defaultFixable.map((f) => f.ruleId)),
  )

  const toggle = (id: string) => {
    setSelected((prev) => {
      const next = new Set(prev)
      if (next.has(id)) next.delete(id)
      else next.add(id)
      return next
    })
  }

  const all = selected.size === fixable.length && fixable.length > 0
  const count = fixable
    .filter((f) => selected.has(f.ruleId))
    .reduce((a, f) => a + f.occurrences.length, 0)

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal fix-modal" role="dialog" aria-modal="true" aria-label="生成修复版" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <div>
            <span className="modal-kicker">REPAIR PLAN</span>
            <h2>生成修复版</h2>
            <p>只处理可以确定的格式项，不改文字内容，原文件不会被修改。</p>
          </div>
          <button className="icon-button large" onClick={onClose} aria-label="关闭">
            <CloseIcon width={18} height={18} />
          </button>
        </div>

        <div className="fix-toolbar">
          <span>
            已选 <b>{selected.size}</b> 类，约 <b>{count}</b> 处
          </span>
          <button
            className="text-button small"
            onClick={() => setSelected(all ? new Set() : new Set(fixable.map((f) => f.ruleId)))}
          >
            {all ? '全部取消' : '全选可修复项'}
          </button>
        </div>

        <div className="fix-list">
          {fixable.map((f) => (
            <label key={f.ruleId} className={`fix-row ${selected.has(f.ruleId) ? 'checked' : ''}`}>
              <input
                type="checkbox"
                checked={selected.has(f.ruleId)}
                onChange={() => toggle(f.ruleId)}
              />
              <span className="fix-box" aria-hidden="true">
                {selected.has(f.ruleId) && (
                  <svg viewBox="0 0 16 16" width="12" height="12">
                    <path d="m3 8 3.2 3.2L13 4.5" fill="none" stroke="currentColor" strokeWidth="2" />
                  </svg>
                )}
              </span>
              <span className="fix-copy">
                <strong>{f.title}</strong>
                <small>{f.occurrences.length} 处 · {f.hint}</small>
              </span>
              {RULE_META[f.ruleId]?.defaultOff && <span className="fix-flag">需确认</span>}
              <span className={`sev-tag sev-${f.severity}`}>{f.severity === 'error' ? '必改' : f.severity === 'warn' ? '建议' : '提示'}</span>
            </label>
          ))}
        </div>

        <div className="fix-note">
          修复会直接写入一份新的 .docx 副本：字体、字号、行距、缩进、对齐、页边距、题注与标题样式中的确定项会被调整；目录、页码、文献编号等仍需手动处理。
          修复的是当前文字的直接格式，不会重写样式表，之后新增的文字仍可能继承原样式。
          目录重建默认不勾选：它会移除现有静态目录条目并插入 TOC 域，需要在 Word 中右键更新。
        </div>

        <div className="modal-actions">
          <button className="chip-button" onClick={onClose} disabled={applying}>
            取消
          </button>
          <button
            className="primary-button"
            disabled={selected.size === 0 || applying}
            onClick={() => onApply(Array.from(selected))}
          >
            <DownloadIcon width={16} height={16} />
            {applying ? '正在生成…' : '生成修复版'}
          </button>
        </div>
      </div>
    </div>
  )
}
