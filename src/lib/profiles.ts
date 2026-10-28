export interface LineRule {
  rule: 'exact' | 'auto'
  twips: number
}

export interface RuleProfile {
  id: string
  name: string
  shortName: string
  source: string
  sourceUrl?: string
  note: string
  page: {
    widthTwips: number
    heightTwips: number
    marginsCm: { top: number; right: number; bottom: number; left: number }
    toleranceTwips: number
  }
  body: {
    eastAsiaFont: string
    asciiFont: string
    sizePt: number
    line: LineRule
    firstLineChars: number
    alignment: 'both'
    beforeAfterMaxTwips: number
  }
  heading: Record<1 | 2 | 3, { font: string; sizePt: number }>
  caption: { font: string; sizePt: number; alignment: 'center' }
  reference: {
    font: string
    asciiFont: string
    sizePt: number
    titleFont: string
    titleSizePt: number
  }
  special: {
    abstractTitle: { font: string; sizePt: number; alignment: 'center' }
    enAbstractTitle: { font: string; sizePt: number; bold: boolean; alignment: 'center' }
    tocTitle: { font: string; sizePt: number; alignment: 'center' }
    ackTitle: { font: string; sizePt: number; alignment: 'center' }
    appendixTitle: { font: string; sizePt: number; alignment: 'center' }
  }
  structure: {
    requireChineseAbstract: boolean
    requireEnglishAbstract: boolean
    requireKeywords: boolean
    requireToc: boolean
    requireReferences: boolean
    requireAcknowledgement: boolean
    acknowledgementTerms: string[]
  }
}

export const TJU_PROFILE: RuleProfile = {
  id: 'tju-undergrad',
  name: '天津大学本科毕业论文（参考）',
  shortName: '天大本科',
  source: '《本科生毕业设计（论文）统一格式的规定（试行稿）》',
  sourceUrl: 'https://fs.duifene.com/hw9/232686/5353960/ahw_04%E6%AF%95%E4%B8%9A%E8%AE%BE%E8%AE%A1%EF%BC%88%E8%AE%BA%E6%96%87%EF%BC%89%E6%92%B0%E5%86%99%E6%A0%BC%E5%BC%8F%E8%A6%81%E6%B1%82231027151202.pdf',
  note: '页面、正文、标题、摘要与参考文献规则取自该试行稿；题注字体、页码位置等未逐项规定的细节，采用中文论文常用规范，请以学院最新模板为准。',
  page: {
    widthTwips: 11906,
    heightTwips: 16838,
    marginsCm: { top: 2.5, right: 2.5, bottom: 2.5, left: 2.5 },
    toleranceTwips: 20,
  },
  body: {
    eastAsiaFont: '宋体',
    asciiFont: 'Times New Roman',
    sizePt: 12,
    line: { rule: 'exact', twips: 400 },
    firstLineChars: 2,
    alignment: 'both',
    beforeAfterMaxTwips: 20,
  },
  heading: {
    1: { font: '黑体', sizePt: 14 },
    2: { font: '黑体', sizePt: 12 },
    3: { font: '楷体', sizePt: 12 },
  },
  caption: { font: '宋体', sizePt: 10.5, alignment: 'center' },
  reference: {
    font: '宋体',
    asciiFont: 'Times New Roman',
    sizePt: 10.5,
    titleFont: '黑体',
    titleSizePt: 14,
  },
  special: {
    abstractTitle: { font: '黑体', sizePt: 16, alignment: 'center' },
    enAbstractTitle: { font: 'Times New Roman', sizePt: 16, bold: true, alignment: 'center' },
    tocTitle: { font: '黑体', sizePt: 16, alignment: 'center' },
    ackTitle: { font: '黑体', sizePt: 14, alignment: 'center' },
    appendixTitle: { font: '黑体', sizePt: 14, alignment: 'center' },
  },
  structure: {
    requireChineseAbstract: true,
    requireEnglishAbstract: true,
    requireKeywords: true,
    requireToc: true,
    requireReferences: true,
    requireAcknowledgement: true,
    acknowledgementTerms: ['致谢', '谢辞'],
  },
}

export const GENERIC_PROFILE: RuleProfile = {
  id: 'generic-cn',
  name: '通用中文论文（宽松）',
  shortName: '通用',
  source: '中文高校毕业论文常见规范汇总',
  note: '适合不确定学校模板时先做一轮通用体检。页边距按 2.54 cm，正文 1.5 倍行距。',
  page: {
    widthTwips: 11906,
    heightTwips: 16838,
    marginsCm: { top: 2.54, right: 2.54, bottom: 2.54, left: 2.54 },
    toleranceTwips: 20,
  },
  body: {
    eastAsiaFont: '宋体',
    asciiFont: 'Times New Roman',
    sizePt: 12,
    line: { rule: 'auto', twips: 360 },
    firstLineChars: 2,
    alignment: 'both',
    beforeAfterMaxTwips: 20,
  },
  heading: {
    1: { font: '黑体', sizePt: 16 },
    2: { font: '黑体', sizePt: 14 },
    3: { font: '黑体', sizePt: 12 },
  },
  caption: { font: '宋体', sizePt: 10.5, alignment: 'center' },
  reference: {
    font: '宋体',
    asciiFont: 'Times New Roman',
    sizePt: 10.5,
    titleFont: '黑体',
    titleSizePt: 14,
  },
  special: {
    abstractTitle: { font: '黑体', sizePt: 16, alignment: 'center' },
    enAbstractTitle: { font: 'Times New Roman', sizePt: 16, bold: true, alignment: 'center' },
    tocTitle: { font: '黑体', sizePt: 16, alignment: 'center' },
    ackTitle: { font: '黑体', sizePt: 14, alignment: 'center' },
    appendixTitle: { font: '黑体', sizePt: 14, alignment: 'center' },
  },
  structure: {
    requireChineseAbstract: true,
    requireEnglishAbstract: true,
    requireKeywords: true,
    requireToc: true,
    requireReferences: true,
    requireAcknowledgement: true,
    acknowledgementTerms: ['致谢', '谢辞'],
  },
}

export const PROFILES: RuleProfile[] = [TJU_PROFILE, GENERIC_PROFILE]

export const cmToTwips = (cm: number) => Math.round(cm * 566.929)

/**
 * 把用户粘贴的 JSON 合并到天大规则包的默认值上，做一次温和校验。
 * 只接受可以安全归一化的字段，避免一条坏规则让整个应用崩溃。
 */
export function parseProfile(raw: unknown): { profile?: RuleProfile; error?: string } {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) {
    return { error: '规则包必须是一个 JSON 对象。' }
  }
  const input = raw as Record<string, any>
  const name = typeof input.name === 'string' ? input.name.trim() : ''
  if (!name) return { error: '规则包缺少 name 字段。' }

  const base = TJU_PROFILE
  const num = (value: unknown, fallback: number, min: number, max: number): number => {
    const n = typeof value === 'number' ? value : Number(value)
    if (!Number.isFinite(n) || n < min || n > max) return fallback
    return n
  }
  const str = (value: unknown, fallback: string): string =>
    typeof value === 'string' && value.trim() ? value.trim() : fallback
  const bool = (value: unknown, fallback: boolean): boolean =>
    typeof value === 'boolean' ? value : fallback
  const terms = Array.isArray(input.structure?.acknowledgementTerms)
    ? input.structure.acknowledgementTerms.filter((x: unknown): x is string => typeof x === 'string' && !!x.trim())
    : base.structure.acknowledgementTerms

  const rawId =
    typeof input.id === 'string' ? input.id.replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 30) : ''
  const builtInIds = new Set(PROFILES.map((p) => p.id))
  const id =
    rawId && !builtInIds.has(rawId)
      ? rawId
      : `${rawId || 'custom'}-${Date.now().toString(36)}`
  const profile: RuleProfile = {
    id,
    name,
    shortName: str(input.shortName, '自定义').slice(0, 8),
    source: str(input.source, '用户导入的自定义规则包'),
    sourceUrl: typeof input.sourceUrl === 'string' ? input.sourceUrl : undefined,
    note: str(input.note, '自定义规则包，请核对后再使用。'),
    page: {
      widthTwips: num(input.page?.widthTwips, base.page.widthTwips, 5000, 20000),
      heightTwips: num(input.page?.heightTwips, base.page.heightTwips, 5000, 30000),
      marginsCm: {
        top: num(input.page?.marginsCm?.top, base.page.marginsCm.top, 0.4, 8),
        right: num(input.page?.marginsCm?.right, base.page.marginsCm.right, 0.4, 8),
        bottom: num(input.page?.marginsCm?.bottom, base.page.marginsCm.bottom, 0.4, 8),
        left: num(input.page?.marginsCm?.left, base.page.marginsCm.left, 0.4, 8),
      },
      toleranceTwips: num(input.page?.toleranceTwips, base.page.toleranceTwips, 0, 200),
    },
    body: {
      eastAsiaFont: str(input.body?.eastAsiaFont, base.body.eastAsiaFont),
      asciiFont: str(input.body?.asciiFont, base.body.asciiFont),
      sizePt: num(input.body?.sizePt, base.body.sizePt, 5, 72),
      line: {
        rule: input.body?.line?.rule === 'auto' ? 'auto' : 'exact',
        twips: num(input.body?.line?.twips, base.body.line.twips, 120, 2400),
      },
      firstLineChars: num(input.body?.firstLineChars, base.body.firstLineChars, 0, 8),
      alignment: 'both',
      beforeAfterMaxTwips: num(input.body?.beforeAfterMaxTwips, base.body.beforeAfterMaxTwips, 0, 200),
    },
    heading: {
      1: {
        font: str(input.heading?.[1]?.font, base.heading[1].font),
        sizePt: num(input.heading?.[1]?.sizePt, base.heading[1].sizePt, 5, 72),
      },
      2: {
        font: str(input.heading?.[2]?.font, base.heading[2].font),
        sizePt: num(input.heading?.[2]?.sizePt, base.heading[2].sizePt, 5, 72),
      },
      3: {
        font: str(input.heading?.[3]?.font, base.heading[3].font),
        sizePt: num(input.heading?.[3]?.sizePt, base.heading[3].sizePt, 5, 72),
      },
    },
    caption: {
      font: str(input.caption?.font, base.caption.font),
      sizePt: num(input.caption?.sizePt, base.caption.sizePt, 5, 72),
      alignment: 'center',
    },
    reference: {
      font: str(input.reference?.font, base.reference.font),
      asciiFont: str(input.reference?.asciiFont, base.reference.asciiFont),
      sizePt: num(input.reference?.sizePt, base.reference.sizePt, 5, 72),
      titleFont: str(input.reference?.titleFont, base.reference.titleFont),
      titleSizePt: num(input.reference?.titleSizePt, base.reference.titleSizePt, 5, 72),
    },
    special: {
      abstractTitle: {
        font: str(input.special?.abstractTitle?.font, base.special.abstractTitle.font),
        sizePt: num(input.special?.abstractTitle?.sizePt, base.special.abstractTitle.sizePt, 5, 72),
        alignment: 'center',
      },
      enAbstractTitle: {
        font: str(input.special?.enAbstractTitle?.font, base.special.enAbstractTitle.font),
        sizePt: num(input.special?.enAbstractTitle?.sizePt, base.special.enAbstractTitle.sizePt, 5, 72),
        bold: bool(input.special?.enAbstractTitle?.bold, base.special.enAbstractTitle.bold),
        alignment: 'center',
      },
      tocTitle: {
        font: str(input.special?.tocTitle?.font, base.special.tocTitle.font),
        sizePt: num(input.special?.tocTitle?.sizePt, base.special.tocTitle.sizePt, 5, 72),
        alignment: 'center',
      },
      ackTitle: {
        font: str(input.special?.ackTitle?.font, base.special.ackTitle.font),
        sizePt: num(input.special?.ackTitle?.sizePt, base.special.ackTitle.sizePt, 5, 72),
        alignment: 'center',
      },
      appendixTitle: {
        font: str(input.special?.appendixTitle?.font, base.special.appendixTitle.font),
        sizePt: num(input.special?.appendixTitle?.sizePt, base.special.appendixTitle.sizePt, 5, 72),
        alignment: 'center',
      },
    },
    structure: {
      requireChineseAbstract: bool(input.structure?.requireChineseAbstract, base.structure.requireChineseAbstract),
      requireEnglishAbstract: bool(input.structure?.requireEnglishAbstract, base.structure.requireEnglishAbstract),
      requireKeywords: bool(input.structure?.requireKeywords, base.structure.requireKeywords),
      requireToc: bool(input.structure?.requireToc, base.structure.requireToc),
      requireReferences: bool(input.structure?.requireReferences, base.structure.requireReferences),
      requireAcknowledgement: bool(input.structure?.requireAcknowledgement, base.structure.requireAcknowledgement),
      acknowledgementTerms: terms.length ? terms : base.structure.acknowledgementTerms,
    },
  }
  return { profile }
}
