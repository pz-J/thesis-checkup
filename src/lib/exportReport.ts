import { CATEGORIES, type Finding, type Report } from './rules'
import type { RuleProfile } from './profiles'

const esc = (s: string) =>
  s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')

const severityLabel = { error: '必改', warn: '建议', info: '提示' } as const

function findingHtml(f: Finding): string {
  const occurrences = f.occurrences
    .map((o) => {
      const loc = o.paragraphIndex !== undefined ? `第 ${o.paragraphIndex + 1} 段` : o.sectionIndex !== undefined ? `第 ${o.sectionIndex + 1} 节` : ''
      const text = o.text ? `<div class="occ-text">${esc(o.text)}</div>` : ''
      const actual = o.actual ? `<div class="occ-actual">实际：${esc(o.actual)}</div>` : ''
      const detail = o.detail ? `<div class="occ-detail">${esc(o.detail)}</div>` : ''
      return `<div class="occ"><span class="occ-loc">${esc(loc)}</span>${text}${actual}${detail}</div>`
    })
    .join('')
  return `<section class="finding">
    <div class="finding-head">
      <span class="sev sev-${f.severity}">${severityLabel[f.severity]}</span>
      <h3>${esc(f.title)}</h3>
      <span class="count">${f.occurrences.length} 处</span>
      ${f.fixable ? '<span class="fixable">可自动修复</span>' : ''}
    </div>
    <p class="hint">${esc(f.hint)}</p>
    ${occurrences}
  </section>`
}

export function buildReportHtml(report: Report, profile: RuleProfile): string {
  const date = new Date(report.generatedAt)
  const stamp = `${date.getFullYear()} 年 ${date.getMonth() + 1} 月 ${date.getDate()} 日 ${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`
  const categories = CATEGORIES.map((c) => {
    const score = report.categoryScores[c.id]
    return `<div class="cat">
      <div class="cat-top"><span>${c.name}</span><strong>${score}</strong></div>
      <div class="bar"><i style="width:${score}%"></i></div>
    </div>`
  }).join('')
  const findings = report.findings.map(findingHtml).join('')
  const empty = report.findings.length
    ? ''
    : '<div class="empty">所有检查项均通过。提交前仍建议人工通读一遍。</div>'

  return `<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>论文体检报告 · ${esc(report.doc.fileName)}</title>
<style>
  :root{--paper:#f7f4ed;--ink:#1b1a18;--muted:#6f6b62;--line:#d9d3c6;--accent:#a63a2b;--good:#2f6b4f;--warn:#a9761e}
  *{box-sizing:border-box}
  body{margin:0;background:var(--paper);color:var(--ink);font-family:-apple-system,"Segoe UI","Microsoft YaHei",sans-serif}
  .wrap{max-width:900px;margin:0 auto;padding:56px 32px 72px}
  header{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:2px solid var(--ink);padding-bottom:20px}
  .brand{font-family:"Songti SC","SimSun",serif;font-size:22px;letter-spacing:.12em}
  .brand small{display:block;font-family:inherit;font-size:10px;letter-spacing:.28em;color:var(--muted);margin-top:4px}
  .meta{text-align:right;font-size:12px;color:var(--muted);line-height:1.8}
  .hero{display:flex;gap:32px;align-items:center;margin:36px 0 8px}
  .score{font-size:96px;line-height:1;font-weight:600;font-variant-numeric:tabular-nums}
  .score small{font-size:18px;color:var(--muted);margin-left:6px}
  .grade{display:inline-block;margin-top:10px;border:1.5px solid var(--accent);color:var(--accent);padding:4px 12px;border-radius:3px;font-size:13px;letter-spacing:.16em;transform:rotate(-2deg)}
  .hero-note{flex:1;font-size:13px;line-height:1.9;color:var(--muted)}
  .hero-note strong{color:var(--ink);font-weight:500}
  .cats{display:grid;grid-template-columns:1fr 1fr;gap:14px 32px;margin:28px 0 36px}
  .cat-top{display:flex;justify-content:space-between;font-size:12px;margin-bottom:5px}
  .bar{height:4px;background:#e5dfd2;border-radius:2px;overflow:hidden}
  .bar i{display:block;height:100%;background:var(--ink)}
  .summary{border:1px solid var(--line);padding:14px 18px;font-size:12px;color:var(--muted);display:flex;gap:18px;flex-wrap:wrap;margin-bottom:34px}
  .summary b{color:var(--ink);font-weight:500}
  h2{font-family:"Songti SC","SimSun",serif;font-size:20px;letter-spacing:.08em;margin:0 0 16px;border-bottom:1px solid var(--line);padding-bottom:10px}
  .finding{padding:16px 0;border-bottom:1px solid var(--line)}
  .finding-head{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
  .finding h3{font-size:15px;margin:0;font-weight:600}
  .sev{font-size:10px;letter-spacing:.08em;padding:2px 6px;border-radius:2px;border:1px solid}
  .sev-error{color:var(--accent);border-color:var(--accent)}
  .sev-warn{color:var(--warn);border-color:var(--warn)}
  .sev-info{color:var(--muted);border-color:var(--line)}
  .count{font-size:11px;color:var(--muted);margin-left:auto}
  .fixable{font-size:10px;color:var(--good);border:1px solid var(--good);padding:2px 6px;border-radius:2px}
  .hint{margin:7px 0 8px;font-size:12px;color:var(--muted);line-height:1.7}
  .occ{border-left:2px solid var(--line);padding:6px 0 6px 12px;margin:6px 0;font-size:12px}
  .occ-loc{font-family:Consolas,monospace;font-size:10px;color:var(--muted);margin-right:8px}
  .occ-text{color:var(--ink);line-height:1.6;margin-top:3px}
  .occ-actual{color:var(--accent);font-size:11px;margin-top:3px}
  .occ-detail{color:var(--muted);font-size:11px;margin-top:2px}
  .empty{padding:40px 0;text-align:center;color:var(--muted);font-size:13px}
  footer{margin-top:48px;border-top:1px solid var(--line);padding-top:16px;font-size:11px;line-height:1.8;color:var(--muted)}
  @media print{body{background:#fff}.wrap{padding:20px}.finding{break-inside:avoid}}
</style></head>
<body><div class="wrap">
  <header>
    <div class="brand">论文体检<small>THESIS CHECKUP</small></div>
    <div class="meta">${esc(report.doc.fileName)}<br/>${stamp}<br/>规则包：${esc(profile.name)}</div>
  </header>
  <div class="hero">
    <div><div class="score">${report.score}<small>/100</small></div><div class="grade">${esc(report.grade.label)}</div></div>
    <div class="hero-note"><strong>${esc(report.grade.description)}</strong><br/>
      共检查 ${report.totalRules} 项规则，${report.passedRules} 项通过。发现 ${report.counts.error} 类必改问题、${report.counts.warn} 类建议问题和 ${report.counts.info} 类提示，合计 ${report.counts.occurrences} 处；其中 ${report.counts.fixable} 类可由工具自动修复。<br/>
      本报告在浏览器本地生成，论文文件没有被上传。</div>
  </div>
  <div class="cats">${categories}</div>
  <div class="summary"><span>段落 <b>${report.doc.stats.paragraphCount}</b></span><span>正文有效段落 <b>${report.doc.stats.bodyParagraphCount}</b></span><span>参考文献 <b>${report.doc.stats.referenceCount}</b> 条</span><span>图表题注 <b>${report.doc.stats.figureCaptionCount + report.doc.stats.tableCaptionCount}</b> 处</span></div>
  <h2>问题清单</h2>
  ${findings}${empty}
  <footer>规则来源：${esc(profile.source)}。<br/>${esc(profile.note)}<br/>本报告仅用于提交前自查，不能替代学院正式审查。生成工具：论文体检 · 本地运行。</footer>
</div></body></html>`
}

export function downloadReport(report: Report, profile: RuleProfile) {
  const html = buildReportHtml(report, profile)
  const blob = new Blob([html], { type: 'text/html;charset=utf-8' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  const base = report.doc.fileName.replace(/\.docx$/i, '')
  a.href = url
  a.download = `${base}_体检报告.html`
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 4000)
}
