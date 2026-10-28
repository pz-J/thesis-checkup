import { CATEGORIES, type Report } from './rules'
import type { RuleProfile } from './profiles'

const W = 1080
const H = 1440
const PAPER = '#f6f3ec'
const INK = '#1b1a17'
const MUTED = '#77736a'
const LINE = '#dcd5c7'
const ACCENT = '#a63a2b'
const GOOD = '#2f6b4f'
const WARN = '#a9761e'

function roundRect(ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number, r: number) {
  ctx.beginPath()
  ctx.moveTo(x + r, y)
  ctx.arcTo(x + w, y, x + w, y + h, r)
  ctx.arcTo(x + w, y + h, x, y + h, r)
  ctx.arcTo(x, y + h, x, y, r)
  ctx.arcTo(x, y, x + w, y, r)
  ctx.closePath()
}

function wrapText(
  ctx: CanvasRenderingContext2D,
  text: string,
  x: number,
  y: number,
  maxWidth: number,
  lineHeight: number,
  maxLines = 4,
): number {
  const chars = Array.from(text)
  let line = ''
  let lines = 0
  for (const ch of chars) {
    const test = line + ch
    if (ctx.measureText(test).width > maxWidth && line) {
      ctx.fillText(line, x, y + lines * lineHeight)
      lines++
      line = ch
      if (lines >= maxLines) break
    } else {
      line = test
    }
  }
  if (lines < maxLines && line) ctx.fillText(line, x, y + lines * lineHeight)
  return y + (lines + 1) * lineHeight
}

export function downloadShareCard(report: Report, profile: RuleProfile) {
  const canvas = document.createElement('canvas')
  canvas.width = W
  canvas.height = H
  const ctx = canvas.getContext('2d')
  if (!ctx) return

  // paper
  ctx.fillStyle = PAPER
  ctx.fillRect(0, 0, W, H)

  // faint grid
  ctx.strokeStyle = 'rgba(27,26,23,0.035)'
  ctx.lineWidth = 1
  for (let x = 0; x <= W; x += 36) {
    ctx.beginPath()
    ctx.moveTo(x, 0)
    ctx.lineTo(x, H)
    ctx.stroke()
  }
  for (let y = 0; y <= H; y += 36) {
    ctx.beginPath()
    ctx.moveTo(0, y)
    ctx.lineTo(W, y)
    ctx.stroke()
  }

  // brand margin line
  ctx.fillStyle = ACCENT
  ctx.fillRect(58, 74, 3, 54)

  ctx.fillStyle = INK
  ctx.font = '600 42px "Songti SC", "SimSun", serif'
  ctx.textAlign = 'left'
  ctx.textBaseline = 'alphabetic'
  ctx.fillText('论文体检', 80, 112)

  ctx.fillStyle = MUTED
  ctx.font = '15px "Microsoft YaHei", sans-serif'
  ctx.fillText('T H E S I S   C H E C K U P', 81, 142)

  const fileName = report.doc.fileName.length > 24 ? `${report.doc.fileName.slice(0, 23)}…` : report.doc.fileName
  const date = new Date(report.generatedAt)
  const dateText = `${date.getFullYear()}.${String(date.getMonth() + 1).padStart(2, '0')}.${String(date.getDate()).padStart(2, '0')}`
  ctx.textAlign = 'right'
  ctx.fillStyle = MUTED
  ctx.font = '17px "Microsoft YaHei", sans-serif'
  ctx.fillText(fileName, 1000, 104)
  ctx.font = '15px "Microsoft YaHei", sans-serif'
  ctx.fillText(dateText, 1000, 134)

  ctx.strokeStyle = LINE
  ctx.lineWidth = 1
  ctx.beginPath()
  ctx.moveTo(80, 180)
  ctx.lineTo(1000, 180)
  ctx.stroke()

  // score
  const scoreColor = report.score >= 90 ? GOOD : report.score >= 80 ? WARN : ACCENT
  ctx.textAlign = 'left'
  ctx.fillStyle = INK
  ctx.font = '600 250px "Songti SC", "SimSun", serif'
  const scoreText = String(report.score)
  ctx.fillText(scoreText, 78, 440)
  const scoreWidth = ctx.measureText(scoreText).width
  ctx.fillStyle = MUTED
  ctx.font = '24px "Microsoft YaHei", sans-serif'
  ctx.fillText('/ 100', 82 + scoreWidth + 14, 430)

  // grade stamp
  ctx.save()
  ctx.translate(790, 300)
  ctx.rotate(-0.035)
  ctx.strokeStyle = scoreColor
  ctx.lineWidth = 2
  roundRect(ctx, -130, -40, 260, 80, 4)
  ctx.stroke()
  ctx.fillStyle = scoreColor
  ctx.font = '600 26px "Microsoft YaHei", sans-serif'
  ctx.textAlign = 'center'
  ctx.fillText(report.grade.label, 0, 9)
  ctx.restore()

  // description
  ctx.textAlign = 'left'
  ctx.fillStyle = MUTED
  ctx.font = '24px "Microsoft YaHei", sans-serif'
  const afterDesc = wrapText(ctx, report.grade.description, 80, 520, 920, 38, 2)

  // summary line
  ctx.fillStyle = INK
  ctx.font = '600 22px "Microsoft YaHei", sans-serif'
  ctx.fillText(
    `${report.counts.error} 类必改 · ${report.counts.warn} 类建议 · ${report.counts.info} 类提示`,
    80,
    afterDesc + 22,
  )
  ctx.fillStyle = MUTED
  ctx.font = '16px "Microsoft YaHei", sans-serif'
  ctx.fillText(`共 ${report.totalRules} 项规则，通过 ${report.passedRules} 项 · 其中 ${report.counts.fixable} 类可自动修复`, 80, afterDesc + 54)

  // categories
  const catTop = 690
  ctx.fillStyle = INK
  ctx.font = '600 20px "Microsoft YaHei", sans-serif'
  ctx.fillText('分类得分', 80, catTop)
  ctx.fillStyle = MUTED
  ctx.font = '14px "Microsoft YaHei", sans-serif'
  ctx.textAlign = 'right'
  ctx.fillText(profile.shortName, 1000, catTop)

  CATEGORIES.forEach((cat, i) => {
    const y = catTop + 52 + i * 66
    const score = report.categoryScores[cat.id]
    ctx.textAlign = 'left'
    ctx.fillStyle = INK
    ctx.font = '18px "Microsoft YaHei", sans-serif'
    ctx.fillText(cat.name, 80, y)
    ctx.fillStyle = '#e7e1d4'
    roundRect(ctx, 250, y - 12, 600, 10, 5)
    ctx.fill()
    ctx.fillStyle = score >= 90 ? INK : score >= 70 ? WARN : ACCENT
    if (score > 0) {
      roundRect(ctx, 250, y - 12, 600 * (score / 100), 10, 5)
      ctx.fill()
    }
    ctx.fillStyle = MUTED
    ctx.font = '15px "Consolas", monospace'
    ctx.textAlign = 'right'
    ctx.fillText(String(score), 1000, y)
  })

  // top findings
  const findTop = 1152
  ctx.textAlign = 'left'
  ctx.fillStyle = INK
  ctx.font = '600 20px "Microsoft YaHei", sans-serif'
  ctx.fillText('优先处理', 80, findTop)

  const list = report.findings.slice(0, 3)
  if (list.length === 0) {
    ctx.fillStyle = GOOD
    ctx.font = '20px "Microsoft YaHei", sans-serif'
    ctx.fillText('所有检查项通过。', 80, findTop + 44)
  } else {
    list.forEach((f, i) => {
      const y = findTop + 44 + i * 42
      const color = f.severity === 'error' ? ACCENT : f.severity === 'warn' ? WARN : '#5c6b7a'
      ctx.fillStyle = color
      ctx.beginPath()
      ctx.arc(88, y - 7, 5, 0, Math.PI * 2)
      ctx.fill()
      ctx.fillStyle = INK
      ctx.font = '19px "Microsoft YaHei", sans-serif'
      const title = f.title.length > 26 ? `${f.title.slice(0, 25)}…` : f.title
      ctx.fillText(title, 106, y)
      ctx.fillStyle = MUTED
      ctx.font = '14px "Microsoft YaHei", sans-serif'
      ctx.textAlign = 'right'
      ctx.fillText(`${f.occurrences.length} 处`, 1000, y)
      ctx.textAlign = 'left'
    })
    if (report.findings.length > 3) {
      ctx.fillStyle = MUTED
      ctx.font = '14px "Microsoft YaHei", sans-serif'
      ctx.fillText(`另有 ${report.findings.length - 3} 类问题，详见体检报告`, 80, findTop + 44 + 3 * 42)
    }
  }

  // footer
  ctx.strokeStyle = LINE
  ctx.beginPath()
  ctx.moveTo(80, 1330)
  ctx.lineTo(1000, 1330)
  ctx.stroke()
  ctx.fillStyle = MUTED
  ctx.font = '15px "Microsoft YaHei", sans-serif'
  ctx.fillText('本地运行 · 文件未上传 · 结果仅供提交前自查', 80, 1368)
  ctx.textAlign = 'right'
  ctx.fillText(profile.shortName, 1000, 1368)

  canvas.toBlob((blob) => {
    if (!blob) return
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `${report.doc.fileName.replace(/\.docx$/i, '')}_体检卡.png`
    document.body.appendChild(a)
    a.click()
    a.remove()
    setTimeout(() => URL.revokeObjectURL(url), 4000)
  }, 'image/png')
}
