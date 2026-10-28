import {
  alignmentText,
  boldRatio,
  displayFont,
  fontMix,
  formatNumber,
  formatSize,
  indentText,
  lineSpacingText,
  majorityAsciiFont,
  majorityFont,
  majoritySize,
  normalizeFontName,
  sizeMix,
  twipsToCmValue,
  type ParaInfo,
  type ParsedDoc,
  type Severity,
} from './docx'
import { cmToTwips, type RuleProfile } from './profiles'

export type CategoryId = 'page' | 'body' | 'heading' | 'caption' | 'reference' | 'structure'

export interface CategoryMeta {
  id: CategoryId
  name: string
  description: string
}

export const CATEGORIES: CategoryMeta[] = [
  { id: 'page', name: '页面与页码', description: '纸张、页边距、页码' },
  { id: 'body', name: '正文', description: '字体、字号、行距、缩进' },
  { id: 'heading', name: '标题层级', description: '标题样式、字号、层级连续性' },
  { id: 'caption', name: '图表题注', description: '图题、表题与编号' },
  { id: 'reference', name: '参考文献', description: '著录格式、编号与类型标识' },
  { id: 'structure', name: '结构完整', description: '摘要、目录、致谢等必要部分' },
]

export interface RuleMeta {
  id: string
  category: CategoryId
  title: string
  weight: number
  severity: Severity
  fixable: boolean
  defaultOff?: boolean
  hint: string
}

const RULES: RuleMeta[] = [
  { id: 'page.size', category: 'page', title: '纸张不是 A4', weight: 6, severity: 'error', fixable: true, hint: '页面设置 → 纸张大小 → A4（21 × 29.7 cm）' },
  { id: 'page.margins', category: 'page', title: '页边距不符合规则', weight: 8, severity: 'error', fixable: true, hint: '页面设置 → 页边距，按规则包调整四边距离' },
  { id: 'page.pageNumber', category: 'page', title: '没有检测到页码字段', weight: 7, severity: 'error', fixable: true, hint: '自动插入居中的 PAGE 页码域；若已有页脚，会在原页脚中追加页码' },

  { id: 'body.font', category: 'body', title: '正文中文字体不一致', weight: 10, severity: 'error', fixable: true, hint: '正文中文字体应为宋体，避免主题字体或直接格式混用' },
  { id: 'body.ascii', category: 'body', title: '正文西文字体不是 Times New Roman', weight: 3, severity: 'warn', fixable: true, hint: '数字与英文使用 Times New Roman，可保留局部符号字体' },
  { id: 'body.size', category: 'body', title: '正文字号不一致', weight: 10, severity: 'error', fixable: true, hint: '正文应为小四（12pt）；脚注、图表题注除外' },
  { id: 'body.spacing', category: 'body', title: '正文行距不符合规则', weight: 10, severity: 'error', fixable: true, hint: '天大规则包要求固定值 20 磅，通用包要求 1.5 倍行距' },
  { id: 'body.indent', category: 'body', title: '正文缺少首行缩进', weight: 7, severity: 'error', fixable: true, hint: '每段首行缩进 2 字符，不要用空格手动缩进' },
  { id: 'body.align', category: 'body', title: '正文未两端对齐', weight: 5, severity: 'warn', fixable: true, hint: '段落 → 对齐方式 → 两端对齐' },
  { id: 'body.beforeAfter', category: 'body', title: '正文段前段后间距不为 0', weight: 2, severity: 'info', fixable: true, hint: '段落 → 间距 → 段前 0、段后 0' },
  { id: 'body.bold', category: 'body', title: '正段落体整段加粗', weight: 2, severity: 'warn', fixable: true, hint: '仅对需要强调的词加粗，正文段落保持常规' },
  { id: 'body.leadingSpace', category: 'body', title: '段首用空格充当缩进', weight: 2, severity: 'warn', fixable: true, hint: '删除段首全角空格，改用段落首行缩进' },
  { id: 'body.consistency', category: 'body', title: '同一段内字体或字号混用', weight: 5, severity: 'warn', fixable: false, hint: '复制粘贴常带入局部格式，建议用格式刷统一' },

  { id: 'heading.style', category: 'heading', title: '标题没有使用标题样式', weight: 4, severity: 'warn', fixable: false, hint: '套用「标题 1/2/3」样式，自动目录才能正确生成' },
  { id: 'heading.font', category: 'heading', title: '标题字体不符合规则', weight: 6, severity: 'error', fixable: true, hint: '一级、二级标题用黑体，三级标题用楷体' },
  { id: 'heading.size', category: 'heading', title: '标题字号不符合规则', weight: 6, severity: 'error', fixable: true, hint: '按规则包设置各级标题字号' },
  { id: 'heading.jump', category: 'heading', title: '标题层级出现跳级', weight: 3, severity: 'warn', fixable: false, hint: '一级标题下应先出现二级标题，再出现三级标题' },
  { id: 'heading.levelMismatch', category: 'heading', title: '标题级别与编号深度不一致', weight: 2, severity: 'warn', fixable: false, hint: '例如「1.1」应使用二级标题样式，而不是一级标题' },
  { id: 'heading.duplicateNumber', category: 'heading', title: '标题编号重复', weight: 2, severity: 'warn', fixable: false, hint: '同一编号出现两次，检查是否复制了标题' },
  { id: 'heading.pageBreak', category: 'heading', title: '章标题可能没有另起一页', weight: 2, severity: 'info', fixable: true, hint: '在章标题前插入分页符，或勾选「段前分页」' },

  { id: 'caption.format', category: 'caption', title: '图表题注格式不一致', weight: 5, severity: 'warn', fixable: true, hint: '题注居中，宋体五号（10.5pt）；图题在图下方，表题在表上方' },
  { id: 'caption.duplicate', category: 'caption', title: '图表编号重复', weight: 2, severity: 'warn', fixable: false, hint: '同一编号出现多次，检查是否复制了题注' },
  { id: 'caption.sequence', category: 'caption', title: '图表编号可能断档', weight: 2, severity: 'warn', fixable: false, hint: '同一组图表编号应连续；连续编号或分章编号都可以' },
  { id: 'caption.position', category: 'caption', title: '图题可能放在图的上方', weight: 2, severity: 'info', fixable: false, hint: '图序与图名应置于图的下方' },
  { id: 'caption.tablePosition', category: 'caption', title: '表题位置可能不对', weight: 2, severity: 'warn', fixable: false, hint: '表序与表名应置于表的上方' },

  { id: 'reference.exists', category: 'reference', title: '参考文献缺失或为空', weight: 5, severity: 'error', fixable: false, hint: '正文后另起一页，列出正文中实际引用过的文献' },
  { id: 'reference.title', category: 'reference', title: '「参考文献」标题格式不对', weight: 3, severity: 'warn', fixable: true, hint: '参考文献标题使用黑体四号，顶格，另起一页' },
  { id: 'reference.font', category: 'reference', title: '参考文献条目字体或字号不一致', weight: 3, severity: 'warn', fixable: true, hint: '中文条目五号宋体，外文条目五号 Times New Roman' },
  { id: 'reference.numbering', category: 'reference', title: '参考文献编号不连续', weight: 3, severity: 'warn', fixable: false, hint: '使用 [1] [2] [3] 连续编号，与正文引用一一对应' },
  { id: 'reference.type', category: 'reference', title: '文献缺少类型标识', weight: 3, severity: 'warn', fixable: false, hint: '期刊 [J]、专著 [M]、学位论文 [D]、会议 [C]、专利 [P]、标准 [S]' },
  { id: 'reference.punctuation', category: 'reference', title: '文献条目末尾缺少句点', weight: 2, severity: 'info', fixable: false, hint: '每条著录以「.」或「。」结束' },
  { id: 'reference.citationMapping', category: 'reference', title: '正文引用编号没有对应文献', weight: 3, severity: 'warn', fixable: false, hint: '正文中的 [n] 应与参考文献编号一一对应，常见于增删文献后忘记同步' },
  { id: 'reference.pageBreak', category: 'reference', title: '参考文献可能没有另起一页', weight: 2, severity: 'info', fixable: true, hint: '「参考文献」标题前插入分页符' },

  { id: 'structure.abstract', category: 'structure', title: '缺少中文摘要', weight: 5, severity: 'error', fixable: false, hint: '添加「摘要」标题与摘要正文' },
  { id: 'structure.abstractEn', category: 'structure', title: '缺少英文摘要', weight: 4, severity: 'error', fixable: false, hint: '添加「Abstract」标题与英文摘要' },
  { id: 'structure.keywords', category: 'structure', title: '缺少关键词', weight: 4, severity: 'error', fixable: false, hint: '中文摘要后添加「关键词：」，英文摘要后添加「Key words：」' },
  { id: 'structure.toc', category: 'structure', title: '缺少目录', weight: 5, severity: 'error', fixable: false, hint: '在摘要之后添加自动目录，至少显示两个层级' },
  { id: 'structure.tocField', category: 'structure', title: '目录可能是静态文本', weight: 3, severity: 'warn', fixable: true, defaultOff: true, hint: '可重建为 TOC 域：会移除现有静态目录条目，插入目录域，之后在 Word 中右键更新' },
  { id: 'structure.tocDepth', category: 'structure', title: '目录层级不足', weight: 3, severity: 'warn', fixable: false, hint: '目录至少显示两个层级，并注明页码' },
  { id: 'structure.keywordCount', category: 'structure', title: '关键词数量可能不合适', weight: 2, severity: 'warn', fixable: false, hint: '中文与英文关键词通常各 3–5 个，用分号分隔' },
  { id: 'structure.ack', category: 'structure', title: '缺少致谢 / 谢辞', weight: 4, severity: 'error', fixable: false, hint: '添加「致谢」或「谢辞」部分' },
  { id: 'structure.ackPageBreak', category: 'structure', title: '谢辞可能没有另起一页', weight: 2, severity: 'warn', fixable: true, hint: '「谢辞」标题前插入分页符' },
  { id: 'structure.appendixPageBreak', category: 'structure', title: '附录可能没有另起一页', weight: 1, severity: 'info', fixable: true, hint: '「附录」标题前插入分页符' },
  { id: 'structure.coverFields', category: 'structure', title: '封面信息可能不完整', weight: 2, severity: 'info', fixable: false, hint: '封面通常包含题目、学院、专业、姓名、学号、指导教师' },
  { id: 'structure.order', category: 'structure', title: '前后部分顺序可能不对', weight: 3, severity: 'warn', fixable: false, hint: '摘要 → 英文摘要 → 目录 → 正文 → 参考文献 → 谢辞 / 附录（以学院模板为准）' },
  { id: 'structure.body', category: 'structure', title: '未检测到正文内容', weight: 5, severity: 'error', fixable: false, hint: '确认文档中包含论文正文，且不是空白模板' },

  { id: 'special.abstractTitle', category: 'heading', title: '「摘要」标题格式不对', weight: 3, severity: 'warn', fixable: true, hint: '「摘要」三号黑体、居中，与正文空一行' },
  { id: 'special.enAbstractTitle', category: 'heading', title: '「Abstract」标题格式不对', weight: 2, severity: 'warn', fixable: true, hint: '「Abstract」三号 Times New Roman、加粗、居中' },
  { id: 'special.tocTitle', category: 'heading', title: '「目录」标题格式不对', weight: 2, severity: 'warn', fixable: true, hint: '「目录」三号黑体、居中' },
  { id: 'special.ackTitle', category: 'heading', title: '「谢辞」标题格式不对', weight: 2, severity: 'warn', fixable: true, hint: '「谢辞」四号黑体、居中、另起一页' },
  { id: 'special.appendixTitle', category: 'heading', title: '「附录」标题格式不对', weight: 2, severity: 'info', fixable: true, hint: '「附录」四号黑体、居中、另起一页' },
  { id: 'special.keywords', category: 'heading', title: '关键词标签格式可能不对', weight: 2, severity: 'info', fixable: false, hint: '「关键词：」标签用小四黑体，内容用小四宋体' },
]

export const RULE_META: Record<string, RuleMeta> = Object.fromEntries(RULES.map((r) => [r.id, r]))
export const RULE_LIST = RULES
export const TOTAL_WEIGHT = RULES.reduce((a, r) => a + r.weight, 0)

export interface Occurrence {
  paragraphIndex?: number
  sectionIndex?: number
  text?: string
  actual?: string
  detail?: string
}

export interface Finding {
  ruleId: string
  category: CategoryId
  title: string
  hint: string
  severity: Severity
  weight: number
  fixable: boolean
  occurrences: Occurrence[]
  penalty: number
}

export interface ReportCounts {
  error: number
  warn: number
  info: number
  fixable: number
  occurrences: number
}

export interface Grade {
  label: string
  tone: 'good' | 'ok' | 'warn' | 'bad'
  description: string
}

export interface Report {
  score: number
  grade: Grade
  findings: Finding[]
  categoryScores: Record<CategoryId, number>
  categoryWeights: Record<CategoryId, number>
  counts: ReportCounts
  passedRules: number
  totalRules: number
  totalWeight: number
  profileName: string
  generatedAt: string
  doc: ParsedDoc
}

const SEVERITY_FACTOR: Record<Severity, number> = { error: 1, warn: 0.5, info: 0.2 }

const snippet = (p: ParaInfo, max = 110): string => {
  const t = p.text.replace(/\s+/g, ' ').trim()
  return t.length > max ? `${t.slice(0, max)}…` : t
}

const cm = (twips?: number): string => (twips === undefined ? '—' : `${formatNumber(twipsToCmValue(twips))} cm`)

const fontMatches = (actual: string | undefined, expected: string): boolean => {
  const a = normalizeFontName(actual)
  const e = normalizeFontName(expected)
  if (!a) return false
  return a === e
}

const sizeMatches = (actual: number | undefined, expected: number, tolerance = 0.2): boolean =>
  actual !== undefined && Math.abs(actual - expected) <= tolerance

const lineMatches = (
  fmt: ParaInfo['format'],
  expected: RuleProfile['body']['line'],
): boolean => {
  if (fmt.lineTwips === undefined) return false
  if (expected.rule === 'exact') return fmt.lineRule === 'exact' && Math.abs(fmt.lineTwips - expected.twips) <= 2
  return fmt.lineRule !== 'exact' && Math.abs(fmt.lineTwips - expected.twips) <= 8
}

const hasIndent = (fmt: ParaInfo['format'], chars: number): boolean => {
  if (fmt.firstLineChars !== undefined && Math.abs(fmt.firstLineChars - chars) <= 0.15) return true
  if (fmt.firstLineTwips !== undefined && Math.abs(fmt.firstLineTwips - chars * 240) <= 24) return true
  return false
}

const isLeftAligned = (a?: string) => !a || a === 'left' || a === 'start'

const isJustified = (a?: string) => a === 'both' || a === 'distribute'

export function evaluate(doc: ParsedDoc, profile: RuleProfile): Report {
  const findings = new Map<string, Finding>()

  const add = (
    ruleId: string,
    occurrence: Occurrence,
    opts?: { severity?: Severity; actual?: string; detail?: string },
  ) => {
    const meta = RULE_META[ruleId]
    let f = findings.get(ruleId)
    if (!f) {
      f = {
        ruleId,
        category: meta.category,
        title: meta.title,
        hint: meta.hint,
        severity: opts?.severity ?? meta.severity,
        weight: meta.weight,
        fixable: meta.fixable,
        occurrences: [],
        penalty: 0,
      }
      findings.set(ruleId, f)
    }
    f.occurrences.push({
      ...occurrence,
      actual: opts?.actual ?? occurrence.actual,
      detail: opts?.detail ?? occurrence.detail,
    })
  }

  // ---------------- 页面与页码 ----------------
  const tol = profile.page.toleranceTwips
  for (const sec of doc.sections) {
    if (!sec.widthTwips || !sec.heightTwips) {
      add('page.size', { sectionIndex: sec.index, detail: '本节没有设置纸张尺寸，Word 会回退到默认纸张' })
    } else if (
      Math.abs(sec.widthTwips - profile.page.widthTwips) > tol ||
      Math.abs(sec.heightTwips - profile.page.heightTwips) > tol
    ) {
      add('page.size', {
        sectionIndex: sec.index,
        actual: `${formatNumber(twipsToCmValue(sec.widthTwips))} × ${formatNumber(twipsToCmValue(sec.heightTwips))} cm`,
        detail: sec.widthTwips > sec.heightTwips ? '当前为横向纸张' : '当前不是规则包要求的纵向 A4',
      })
    }

    const mismatch: string[] = []
    for (const key of ['top', 'right', 'bottom', 'left'] as const) {
      const expected = cmToTwips(profile.page.marginsCm[key])
      const actual = sec.margins[key]
      if (actual === undefined || Math.abs(actual - expected) > tol) {
        mismatch.push(`${key === 'top' ? '上' : key === 'bottom' ? '下' : key === 'left' ? '左' : '右'} ${cm(actual)}`)
      }
    }
    if (mismatch.length) {
      add('page.margins', {
        sectionIndex: sec.index,
        actual: mismatch.join(' · '),
        detail: `规则要求：上 ${profile.page.marginsCm.top} / 下 ${profile.page.marginsCm.bottom} / 左 ${profile.page.marginsCm.left} / 右 ${profile.page.marginsCm.right} cm`,
      })
    }

    if (!sec.hasPageNumberField) {
      add('page.pageNumber', {
        sectionIndex: sec.index,
        detail: sec.hasFooterReference ? '页脚存在，但没有找到自动页码字段（PAGE）' : '本节没有页脚，也没有自动页码',
      })
    }
  }

  // ---------------- 正文 ----------------
  const bodyParas = doc.paragraphs.filter(
    (p) =>
      ['body', 'abstract', 'abstract-en'].includes(p.category) &&
      !p.inTable &&
      !p.inTextBox &&
      p.charCount >= 15,
  )

  for (const p of bodyParas) {
    const chineseChars = p.text.replace(/[^\u3400-\u9fff]/g, '').length
    const font = majorityFont(p)
    if (p.category !== 'abstract-en' && chineseChars >= 5 && !fontMatches(font, profile.body.eastAsiaFont)) {
      add('body.font', { paragraphIndex: p.index, text: snippet(p), actual: displayFont(font) })
    }

    const asciiChars = p.text.replace(/[^\x00-\x7F]/g, '').replace(/\s/g, '').length
    if (asciiChars >= 5) {
      const ascii = majorityAsciiFont(p)
      if (!fontMatches(ascii, profile.body.asciiFont)) {
        add('body.ascii', {
          paragraphIndex: p.index,
          text: snippet(p),
          actual: displayFont(ascii),
        })
      }
    }

    const size = majoritySize(p)
    if (!sizeMatches(size, profile.body.sizePt)) {
      add('body.size', { paragraphIndex: p.index, text: snippet(p), actual: formatSize(size) })
    }

    if (!lineMatches(p.format, profile.body.line)) {
      add('body.spacing', {
        paragraphIndex: p.index,
        text: snippet(p),
        actual: lineSpacingText(p.format),
        detail:
          profile.body.line.rule === 'exact'
            ? '规则要求：固定值 20 磅'
            : '规则要求：1.5 倍行距',
      })
    }

    if (p.category === 'body' && !hasIndent(p.format, profile.body.firstLineChars)) {
      add('body.indent', {
        paragraphIndex: p.index,
        text: snippet(p),
        actual: indentText(p.format),
      })
    }

    if (!isJustified(p.format.alignment)) {
      add('body.align', {
        paragraphIndex: p.index,
        text: snippet(p),
        actual: alignmentText(p.format.alignment),
      })
    }

    if ((p.format.beforeTwips ?? 0) > profile.body.beforeAfterMaxTwips || (p.format.afterTwips ?? 0) > profile.body.beforeAfterMaxTwips) {
      add('body.beforeAfter', {
        paragraphIndex: p.index,
        text: snippet(p),
        actual: `段前 ${formatNumber((p.format.beforeTwips ?? 0) / 20)} 磅 / 段后 ${formatNumber((p.format.afterTwips ?? 0) / 20)} 磅`,
      })
    }

    if (boldRatio(p) > 0.8) {
      add('body.bold', { paragraphIndex: p.index, text: snippet(p), actual: '整段加粗' })
    }

    if (/^(\u3000|\s{2,})/.test(p.text)) {
      add('body.leadingSpace', {
        paragraphIndex: p.index,
        text: snippet(p),
        actual: p.text.startsWith('\u3000') ? '段首有全角空格' : '段首有多个空格',
      })
    }

    const fonts = fontMix(p)
    const sizes = sizeMix(p).filter((s) => size !== undefined && Math.abs(s - size) > 1)
    if (fonts.length > 1 || sizes.length > 0) {
      add('body.consistency', {
        paragraphIndex: p.index,
        text: snippet(p),
        actual: [
          fonts.length > 1 ? `字体：${fonts.join(' / ')}` : '',
          sizes.length ? `字号：${sizes.map((s) => formatSize(s)).join(' / ')}` : '',
        ]
          .filter(Boolean)
          .join('；'),
      })
    }
  }

  // ---------------- 标题 ----------------
  const headingParas = doc.paragraphs.filter((p) => p.category === 'heading' && p.headingLevel)
  const headingNumbers = new Map<string, ParaInfo>()
  for (const p of headingParas) {
    if (!p.isHeadingStyle) {
      add('heading.style', { paragraphIndex: p.index, text: snippet(p), actual: `样式：${p.styleName}` })
    }
    const level = p.headingLevel as 1 | 2 | 3
    const expected = profile.heading[level]
    if (level >= 1 && level <= 3) {
      const font = majorityFont(p)
      if (!fontMatches(font, expected.font)) {
        add('heading.font', {
          paragraphIndex: p.index,
          text: snippet(p),
          actual: `${displayFont(font)} · ${level} 级标题`,
        })
      }
      const size = majoritySize(p)
      if (!sizeMatches(size, expected.sizePt)) {
        add('heading.size', {
          paragraphIndex: p.index,
          text: snippet(p),
          actual: formatSize(size),
        })
      }
    }

    const numberMatch = p.text.trim().match(/^(\d+(?:\.\d+){0,2})(?=[\s　.、]|$)/)
    if (numberMatch) {
      const token = numberMatch[1]
      const depth = token.split('.').length
      if (depth !== p.headingLevel) {
        add('heading.levelMismatch', {
          paragraphIndex: p.index,
          text: snippet(p),
          actual: `编号 ${token} 有 ${depth} 层，但使用了 ${p.headingLevel} 级标题`,
        })
      }
      const first = headingNumbers.get(token)
      if (first) {
        add('heading.duplicateNumber', {
          paragraphIndex: p.index,
          text: snippet(p),
          detail: `与第 ${first.index + 1} 段「${snippet(first, 24)}」编号相同`,
        })
      } else {
        headingNumbers.set(token, p)
      }
    }
  }

  let prevHeadingLevel: number | null = null
  let firstLevel1Index: number | null = null
  for (const p of headingParas) {
    if (prevHeadingLevel !== null && p.headingLevel! > prevHeadingLevel + 1) {
      add('heading.jump', {
        paragraphIndex: p.index,
        text: snippet(p),
        detail: `从 ${prevHeadingLevel} 级直接跳到 ${p.headingLevel} 级`,
      })
    }
    prevHeadingLevel = p.headingLevel

    if (p.headingLevel === 1) {
      if (firstLevel1Index === null) firstLevel1Index = p.index
      if (firstLevel1Index !== p.index && p.category === 'heading') {
        const prev = lastNonEmptyBefore(doc, p.index)
        const startsNewPage =
          !!p.format.pageBreakBefore ||
          !!prev?.hasPageBreak ||
          !!prev?.hasSectionBreak ||
          !!p.hasSectionBreak
        if (!startsNewPage) {
          add('heading.pageBreak', {
            paragraphIndex: p.index,
            text: snippet(p),
            actual: '上一段结束时没有分页',
          })
        }
      }
    }
  }

  // ---------------- 图表题注 ----------------
  const captionParas = doc.paragraphs.filter((p) => p.category === 'caption' && p.captionNumber)
  for (const p of captionParas) {
    const font = majorityFont(p)
    const size = majoritySize(p)
    const problems: string[] = []
    if (!fontMatches(font, profile.caption.font)) problems.push(`字体 ${displayFont(font)}`)
    if (!sizeMatches(size, profile.caption.sizePt)) problems.push(`字号 ${formatSize(size)}`)
    if (!isCenter(p.format.alignment)) problems.push(`对齐 ${alignmentText(p.format.alignment)}`)
    if (problems.length) {
      add('caption.format', {
        paragraphIndex: p.index,
        text: snippet(p),
        actual: problems.join(' · '),
      })
    }

    if (p.captionKind === 'figure') {
      const next = nextNonEmpty(doc, p.index)
      const prev = lastNonEmptyBefore(doc, p.index)
      if (next?.hasDrawing) {
        add('caption.position', {
          paragraphIndex: p.index,
          text: snippet(p),
          actual: '图题在图的上方',
        })
      } else if (!prev?.hasDrawing) {
        add('caption.position', {
          paragraphIndex: p.index,
          text: snippet(p),
          actual: '未找到紧邻的图片',
          detail: '可能是图片与题注之间存在空段，或题注是手打的',
        })
      }
    }

    if (p.captionKind === 'table') {
      if (p.prevIsTable) {
        add('caption.tablePosition', {
          paragraphIndex: p.index,
          text: snippet(p),
          actual: '表题在表格下方',
        })
      } else if (!p.nextIsTable) {
        add('caption.tablePosition', {
          paragraphIndex: p.index,
          text: snippet(p),
          actual: '未找到紧邻的表格',
          detail: '表序与表名应位于表格上方',
        })
      }
    }
  }

  const captionGroups = new Map<string, ParaInfo[]>()
  for (const p of captionParas) {
    const key = `${p.captionKind}:${p.captionNumber}`
    captionGroups.set(key, [...(captionGroups.get(key) ?? []), p])
  }
  for (const [, list] of captionGroups) {
    if (list.length > 1) {
      for (const p of list) {
        add('caption.duplicate', {
          paragraphIndex: p.index,
          text: snippet(p),
          actual: `${p.captionKind === 'figure' ? '图' : '表'} ${p.captionNumber} 出现 ${list.length} 次`,
        })
      }
    }
  }

  const captionSequences = new Map<string, { num: number; p: ParaInfo }[]>()
  for (const p of captionParas) {
    const match = p.captionNumber?.match(/^(\d+)(?:[-.](\d+))?$/)
    if (!match) continue
    const prefix = match[2] ? match[1] : ''
    const num = Number(match[2] ?? match[1])
    const key = `${p.captionKind}:${prefix}`
    captionSequences.set(key, [...(captionSequences.get(key) ?? []), { num, p }])
  }
  for (const [, list] of captionSequences) {
    list.sort((a, b) => a.p.index - b.p.index)
    for (let i = 1; i < list.length; i++) {
      const prev = list[i - 1]
      const cur = list[i]
      if (cur.num === prev.num) continue
      if (cur.num !== prev.num + 1) {
        add('caption.sequence', {
          paragraphIndex: cur.p.index,
          text: snippet(cur.p),
          detail: `上一编号为 ${prev.p.captionNumber}，当前为 ${cur.p.captionNumber}`,
        })
      }
    }
  }

  // ---------------- 参考文献 ----------------
  const refTitle = doc.paragraphs.find((p) => p.category === 'reference-title')
  const refEntries = doc.paragraphs.filter((p) => p.category === 'reference' && p.charCount > 5)
  if (!refTitle || refEntries.length === 0) {
    add('reference.exists', {
      paragraphIndex: refTitle?.index,
      detail: refTitle ? '有「参考文献」标题，但没有检测到条目' : '没有找到「参考文献」部分',
    })
  }

  if (refTitle) {
    const font = majorityFont(refTitle)
    const size = majoritySize(refTitle)
    const problems: string[] = []
    if (!fontMatches(font, profile.reference.titleFont)) problems.push(`字体 ${displayFont(font)}`)
    if (!sizeMatches(size, profile.reference.titleSizePt)) problems.push(`字号 ${formatSize(size)}`)
    if (refTitle.format.alignment === 'center') problems.push('标题被居中（规则要求顶格）')
    if (problems.length) {
      add('reference.title', {
        paragraphIndex: refTitle.index,
        text: snippet(refTitle),
        actual: problems.join(' · '),
      })
    }

    const prev = lastNonEmptyBefore(doc, refTitle.index)
    const startsNewPage =
      !!refTitle.format.pageBreakBefore || !!prev?.hasPageBreak || !!refTitle.hasSectionBreak
    if (!startsNewPage) {
      add('reference.pageBreak', { paragraphIndex: refTitle.index, text: snippet(refTitle), actual: '上一段结束时没有分页' })
    }
  }

  refEntries.forEach((p, i) => {
    const font = majorityFont(p)
    const size = majoritySize(p)
    const problems: string[] = []
    if (!fontMatches(font, profile.reference.font)) problems.push(`中文字体 ${displayFont(font)}`)
    if (!sizeMatches(size, profile.reference.sizePt)) problems.push(`字号 ${formatSize(size)}`)
    if (problems.length) {
      add('reference.font', { paragraphIndex: p.index, text: snippet(p), actual: problems.join(' · ') })
    }

    const numberMatch = p.text.trim().match(/^[\[［]\s*(\d+)\s*[\]］]/)
    if (!numberMatch) {
      add('reference.numbering', { paragraphIndex: p.index, text: snippet(p), actual: '缺少 [序号]' })
    } else if (Number(numberMatch[1]) !== i + 1) {
      add('reference.numbering', {
        paragraphIndex: p.index,
        text: snippet(p),
        actual: `第 ${i + 1} 条写成了 [${numberMatch[1]}]`,
      })
    }

    if (!/\[(J|M|D|C|A|P|S|N|R|Z|EB\/OL)\]/i.test(p.text)) {
      add('reference.type', { paragraphIndex: p.index, text: snippet(p), actual: '未找到 [J]/[M]/[D] 等类型标识' })
    }

    const trimmed = p.text.trim()
    if (!/[.。．]$/.test(trimmed)) {
      add('reference.punctuation', { paragraphIndex: p.index, text: snippet(p), actual: '末尾没有句点' })
    }
  })

  const refNumbers = new Set<number>()
  for (const p of refEntries) {
    const match = p.text.trim().match(/^[\[［]\s*(\d+)\s*[\]］]/)
    if (match) refNumbers.add(Number(match[1]))
  }
  const cited = new Map<number, ParaInfo>()
  for (const p of doc.paragraphs) {
    if (!['body', 'abstract', 'abstract-en'].includes(p.category) || p.inTable || p.inTextBox) continue
    for (const match of p.text.matchAll(/[\[［]\s*(\d+(?:\s*[-–—,，]\s*\d+)*)\s*[\]］]/g)) {
      for (const n of expandCitation(match[1])) {
        if (n > 0 && n <= 999 && !cited.has(n)) cited.set(n, p)
      }
    }
  }
  if (refEntries.length > 0 && cited.size > 0) {
    const missing = Array.from(cited.keys()).filter((n) => !refNumbers.has(n))
    if (missing.length > 0) {
      const first = cited.get(missing[0])!
      add('reference.citationMapping', {
        paragraphIndex: first.index,
        text: snippet(first),
        detail: `正文引用了 [${missing.join('、')}]，但参考文献中没有这些编号`,
      })
    }
  }

  // ---------------- 结构 ----------------
  const marker = (category: ParaInfo['category']) => doc.paragraphs.find((p) => p.category === category)
  const abstract = marker('abstract-title')
  const abstractEn = marker('abstract-en-title')
  const keywords = marker('keywords')
  const toc = marker('toc-title')
  const ack = marker('ack-title')

  if (profile.structure.requireChineseAbstract && !abstract) {
    add('structure.abstract', { detail: '没有找到「摘要」标题' })
  }
  if (profile.structure.requireEnglishAbstract && !abstractEn) {
    add('structure.abstractEn', { detail: '没有找到「Abstract」标题' })
  }
  if (profile.structure.requireKeywords && !keywords) {
    add('structure.keywords', { detail: '没有找到「关键词：」或「Key words：」' })
  }
  if (profile.structure.requireToc && !toc) {
    add('structure.toc', { detail: '没有找到「目录」标题' })
  }
  if (toc && !doc.hasTocField) {
    add('structure.tocField', {
      paragraphIndex: toc.index,
      text: snippet(toc),
      detail: '目录存在，但文档中没有 TOC 域；更新页码时需要手动改',
    })
    const tocEntries = doc.paragraphs.filter((p) => p.category === 'toc' && !p.empty)
    const levels = new Set(tocEntries.map(tocLevelOf))
    if (tocEntries.length > 0 && levels.size < 2) {
      add('structure.tocDepth', {
        paragraphIndex: toc.index,
        text: snippet(toc),
        actual: `只检测到 ${levels.size} 个层级`,
        detail: '天大试行稿要求目录至少有两个层次，并注明页码',
      })
    }
  }

  for (const p of doc.paragraphs.filter((item) => item.category === 'keywords')) {
    const match = p.text.match(/^(关键词|key\s*words?)\s*[:：]?/i)
    const rest = match ? p.text.slice(match[0].length) : p.text
    const items = rest.split(/[；;，,]/).map((item) => item.trim()).filter(Boolean)
    if (items.length < 3 || items.length > 5) {
      add('structure.keywordCount', {
        paragraphIndex: p.index,
        text: snippet(p),
        actual: `${items.length} 个关键词`,
      })
    }
  }
  if (profile.structure.requireAcknowledgement && !ack) {
    add('structure.ack', { detail: `没有找到 ${profile.structure.acknowledgementTerms.join(' / ')}` })
  }
  if (ack && !startsNewPage(doc, ack)) {
    add('structure.ackPageBreak', {
      paragraphIndex: ack.index,
      text: snippet(ack),
      actual: '上一段结束时没有分页',
    })
  }
  const appendix = marker('appendix-title')
  if (appendix && !startsNewPage(doc, appendix)) {
    add('structure.appendixPageBreak', {
      paragraphIndex: appendix.index,
      text: snippet(appendix),
      actual: '上一段结束时没有分页',
    })
  }

  const coverText = doc.paragraphs
    .filter((p) => p.category === 'cover')
    .map((p) => p.text)
    .join(' ')
  const requiredCoverFields = ['题目', '学院', '专业', '姓名', '学号', '指导教师']
  const missingCoverFields = requiredCoverFields.filter((field) => !coverText.includes(field))
  if (missingCoverFields.length > 0) {
    add('structure.coverFields', {
      detail: `封面可能缺少：${missingCoverFields.join('、')}；如封面为图片或表格，可忽略此项`,
    })
  }

  const bodyHeading = doc.paragraphs.find((p) => p.category === 'heading' && p.zone === 'body')
  const bodyStart = bodyHeading?.index ?? doc.paragraphs.find((p) => p.category === 'body' && p.charCount > 30)?.index ?? Infinity
  if (bodyParas.length < 3) {
    add('structure.body', { detail: '正文有效段落少于 3 段' })
  }

  const orderIssues: string[] = []
  if (abstract && abstractEn && abstract.index > abstractEn.index) orderIssues.push('中文摘要应在英文摘要之前')
  if (toc && abstractEn && toc.index < abstractEn.index) orderIssues.push('目录应在英文摘要之后')
  if (toc && bodyStart < toc.index) orderIssues.push('目录应在正文之前')
  if (refTitle && bodyStart !== Infinity && refTitle.index < bodyStart) orderIssues.push('参考文献应在正文之后')
  if (refTitle && toc && refTitle.index < toc.index) orderIssues.push('参考文献出现在目录之前')
  if (orderIssues.length) {
    add('structure.order', { detail: orderIssues.join('；') })
  }

  // ---------------- 特殊标题 ----------------
  const checkSpecialTitle = (
    ruleId: string,
    p: ParaInfo | undefined,
    expected: { font: string; sizePt: number; alignment: 'center' },
    boldRequired = false,
  ) => {
    if (!p) return
    const problems: string[] = []
    const font = majorityFont(p)
    const size = majoritySize(p)
    if (!fontMatches(font, expected.font)) problems.push(`字体 ${displayFont(font)}`)
    if (!sizeMatches(size, expected.sizePt)) problems.push(`字号 ${formatSize(size)}`)
    if (p.format.alignment !== expected.alignment) problems.push(`对齐 ${alignmentText(p.format.alignment)}`)
    if (boldRequired && boldRatio(p) < 0.5) problems.push('未加粗')
    if (problems.length) {
      add(ruleId, { paragraphIndex: p.index, text: snippet(p), actual: problems.join(' · ') })
    }
  }

  checkSpecialTitle('special.abstractTitle', abstract, profile.special.abstractTitle)
  checkSpecialTitle('special.enAbstractTitle', abstractEn, profile.special.enAbstractTitle, profile.special.enAbstractTitle.bold)
  checkSpecialTitle('special.tocTitle', toc, profile.special.tocTitle)
  checkSpecialTitle('special.ackTitle', ack, profile.special.ackTitle)
  checkSpecialTitle('special.appendixTitle', marker('appendix-title'), profile.special.appendixTitle)

  if (keywords) {
    const labelRun = keywords.runs.find((r) => /关键词|key\s*words/i.test(r.text))
    const font = labelRun?.format.eastAsiaFont ?? keywords.baseRun.eastAsiaFont
    if (!fontMatches(font, '黑体')) {
      add('special.keywords', {
        paragraphIndex: keywords.index,
        text: snippet(keywords),
        actual: `标签字体 ${displayFont(font)}`,
      })
    }
  }

  // ---------------- 计分 ----------------
  const categoryWeights = Object.fromEntries(CATEGORIES.map((c) => [c.id, 0])) as Record<CategoryId, number>
  for (const rule of RULES) categoryWeights[rule.category] += rule.weight

  const list = Array.from(findings.values())
  for (const f of list) f.penalty = f.weight * SEVERITY_FACTOR[f.severity]

  const totalPenalty = list.reduce((a, f) => a + f.penalty, 0)
  const score = Math.max(0, Math.round(((TOTAL_WEIGHT - totalPenalty) / TOTAL_WEIGHT) * 100))

  const categoryScores = Object.fromEntries(
    CATEGORIES.map((c) => {
      const w = categoryWeights[c.id]
      const p = list.filter((f) => f.category === c.id).reduce((a, f) => a + f.penalty, 0)
      return [c.id, w ? Math.max(0, Math.round(((w - p) / w) * 100)) : 100]
    }),
  ) as Record<CategoryId, number>

  list.sort((a, b) => {
    const rank: Record<Severity, number> = { error: 0, warn: 1, info: 2 }
    return rank[a.severity] - rank[b.severity] || b.weight - a.weight
  })

  const counts: ReportCounts = {
    error: list.filter((f) => f.severity === 'error').length,
    warn: list.filter((f) => f.severity === 'warn').length,
    info: list.filter((f) => f.severity === 'info').length,
    fixable: list.filter((f) => f.fixable).length,
    occurrences: list.reduce((a, f) => a + f.occurrences.length, 0),
  }

  return {
    score,
    grade: gradeFor(score),
    findings: list,
    categoryScores,
    categoryWeights,
    counts,
    passedRules: RULES.length - list.length,
    totalRules: RULES.length,
    totalWeight: TOTAL_WEIGHT,
    profileName: profile.name,
    generatedAt: new Date().toISOString(),
    doc,
  }
}

function gradeFor(score: number): Grade {
  if (score >= 90) return { label: '基本可交', tone: 'good', description: '格式风险很低，把剩余几条提示处理完再交更稳。' }
  if (score >= 80) return { label: '小修即可', tone: 'ok', description: '主体没问题，处理几条明显问题就能交。' }
  if (score >= 70) return { label: '需要修改', tone: 'warn', description: '有几处会被老师一眼看出来的格式问题。' }
  if (score >= 60) return { label: '问题较多', tone: 'bad', description: '建议按报告逐条修，再复检一次。' }
  return { label: '建议重排', tone: 'bad', description: '格式基础较弱，先套用学校模板再逐条修。' }
}

function isCenter(a?: string) {
  return a === 'center'
}

function tocLevelOf(p: ParaInfo): number {
  const style = p.styleName.toLowerCase()
  const styleLevel = style.match(/(?:toc|目录)\s*([1-3])/)
  if (styleLevel) return Number(styleLevel[1])
  const text = p.text.trim()
  if (/^\d+\.\d+\.\d+/.test(text)) return 3
  if (/^\d+\.\d+/.test(text)) return 2
  if (/^\d+/.test(text)) return 1
  return 1
}

function startsNewPage(doc: ParsedDoc, p: ParaInfo): boolean {
  const prev = lastNonEmptyBefore(doc, p.index)
  return (
    !!p.format.pageBreakBefore ||
    !!prev?.hasPageBreak ||
    !!p.hasSectionBreak ||
    !!prev?.hasSectionBreak
  )
}

function expandCitation(raw: string): number[] {
  const out: number[] = []
  for (const part of raw.split(/[,，]/)) {
    const range = part.trim().match(/^(\d+)\s*[-–—]\s*(\d+)$/)
    if (range) {
      const a = Number(range[1])
      const b = Number(range[2])
      if (b >= a && b - a <= 30) {
        for (let n = a; n <= b; n++) out.push(n)
      }
    } else {
      const n = Number(part.trim())
      if (Number.isFinite(n)) out.push(n)
    }
  }
  return out
}

function nextNonEmpty(doc: ParsedDoc, index: number): ParaInfo | undefined {
  for (let i = index + 1; i < doc.paragraphs.length; i++) {
    if (!doc.paragraphs[i].empty || doc.paragraphs[i].hasDrawing) return doc.paragraphs[i]
  }
  return undefined
}

function lastNonEmptyBefore(doc: ParsedDoc, index: number): ParaInfo | undefined {
  for (let i = index - 1; i >= 0; i--) {
    if (!doc.paragraphs[i].empty || doc.paragraphs[i].hasDrawing) return doc.paragraphs[i]
  }
  return undefined
}
