import JSZip from 'jszip'

export const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

export type Severity = 'error' | 'warn' | 'info'

export type Alignment = 'left' | 'center' | 'right' | 'both' | 'distribute' | string

export interface RunFormat {
  eastAsiaFont?: string
  asciiFont?: string
  sizePt?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean
  color?: string
}

export interface ParaFormat {
  alignment?: Alignment
  lineTwips?: number
  lineRule?: 'auto' | 'exact' | 'atLeast' | string
  beforeTwips?: number
  afterTwips?: number
  firstLineChars?: number
  firstLineTwips?: number
  leftIndentTwips?: number
  pageBreakBefore?: boolean
  outlineLevel?: number
  keepNext?: boolean
  hasNumbering?: boolean
}

export interface RunInfo {
  text: string
  format: RunFormat
  hasDrawing: boolean
  fieldInstruction?: string
}

export interface ParaInfo {
  index: number
  text: string
  styleId: string
  styleName: string
  format: ParaFormat
  runs: RunInfo[]
  baseRun: RunFormat
  headingLevel: number | null
  isHeadingStyle: boolean
  isCaption: boolean
  captionKind: 'figure' | 'table' | null
  captionNumber: string | null
  inTable: boolean
  inTextBox: boolean
  hasDrawing: boolean
  hasPageBreak: boolean
  hasTocField: boolean
  hasSectionBreak: boolean
  prevIsTable: boolean
  nextIsTable: boolean
  sectionIndex: number
  category: ParaCategory
  zone: Zone
  empty: boolean
  charCount: number
}

export type Zone = 'front' | 'abstract' | 'abstractEn' | 'toc' | 'body' | 'references' | 'ack' | 'appendix'

export type ParaCategory =
  | 'cover'
  | 'empty'
  | 'abstract-title'
  | 'abstract'
  | 'abstract-en-title'
  | 'abstract-en'
  | 'keywords'
  | 'toc-title'
  | 'toc'
  | 'heading'
  | 'caption'
  | 'reference-title'
  | 'reference'
  | 'ack-title'
  | 'ack'
  | 'appendix-title'
  | 'appendix'
  | 'body'

export interface Margins {
  top: number
  right: number
  bottom: number
  left: number
  header: number
  footer: number
  gutter: number
}

export interface SectionInfo {
  index: number
  widthTwips?: number
  heightTwips?: number
  orientation?: string
  margins: Partial<Margins>
  startParaIndex: number
  hasPageNumberField: boolean
  hasFooterReference: boolean
  hasHeaderReference: boolean
  footerText: string
  headerText: string
}

export interface StyleDef {
  id: string
  name: string
  type: string
  basedOn?: string
  pPr: Element | null
  rPr: Element | null
}

export interface DistinctValue {
  value: string
  count: number
}

export interface DocStats {
  paragraphCount: number
  bodyParagraphCount: number
  bodyCharCount: number
  distinctBodyFonts: DistinctValue[]
  distinctBodySizes: DistinctValue[]
  distinctBodySpacings: DistinctValue[]
  figureCaptionCount: number
  tableCaptionCount: number
  referenceCount: number
  headingCounts: Record<number, number>
}

export interface CoreProps {
  title?: string
  creator?: string
  lastModifiedBy?: string
  modified?: string
}

export interface ParsedDoc {
  fileName: string
  fileSize: number
  paragraphs: ParaInfo[]
  sections: SectionInfo[]
  hasTocField: boolean
  styles: Map<string, StyleDef>
  core: CoreProps
  stats: DocStats
  warnings: string[]
  buffer: ArrayBuffer
}

const wAttr = (el: Element | null | undefined, name: string): string | null => {
  if (!el) return null
  return el.getAttributeNS(W_NS, name) ?? el.getAttribute(`w:${name}`) ?? el.getAttribute(name)
}

const rAttr = (el: Element | null | undefined, name: string): string | null => {
  if (!el) return null
  return el.getAttributeNS(R_NS, name) ?? el.getAttribute(`r:${name}`) ?? el.getAttribute(name)
}

const kids = (el: Element | null | undefined, localName: string): Element[] =>
  el ? Array.from(el.children).filter((c) => c.localName === localName) : []

const kid = (el: Element | null | undefined, localName: string): Element | null =>
  el ? Array.from(el.children).find((c) => c.localName === localName) ?? null : null

const first = (root: Document | Element, localName: string): Element | null =>
  root.getElementsByTagNameNS(W_NS, localName)[0] ?? null

const boolAttrEl = (el: Element | null): boolean | undefined => {
  if (!el) return undefined
  const v = wAttr(el, 'val')
  if (v === null) return true
  return !['0', 'false', 'off'].includes(v)
}

const twipsToCm = (twips: number) => Math.round((twips / 566.929) * 100) / 100

export const twipsToCmValue = twipsToCm

const fontKey = (name?: string): string => (name ?? '').replace(/["']/g, '').replace(/\s+/g, '').toLowerCase()

export function normalizeFontName(name?: string): string {
  const k = fontKey(name)
  if (!k) return ''
  if (['simsun', '宋体', 'songti', 'songtisc', 'stsong'].includes(k) || k.includes('simsun')) return '宋体'
  if (k.includes('nsimsun') || k.includes('新宋体')) return '新宋体'
  if (['simhei', '黑体', 'heiti', 'heitisc', 'stheiti'].includes(k) || k.includes('simhei')) return '黑体'
  if (['kaiti', '楷体', '楷体gb2312', 'kaitigb2312', 'kaitisc'].includes(k) || k.includes('kaiti')) return '楷体'
  if (['fangsong', '仿宋', '仿宋gb2312', 'fangsonggb2312'].includes(k) || k.includes('fangsong')) return '仿宋'
  if (['timesnewroman', 'times', 'tnr'].includes(k) || k.includes('timesnewroman')) return 'Times New Roman'
  return name?.trim() ?? ''
}

export function displayFont(name?: string): string {
  return normalizeFontName(name) || '未显式设置'
}

interface ThemeFonts {
  majorLatin?: string
  minorLatin?: string
  majorEa?: string
  minorEa?: string
  majorCs?: string
  minorCs?: string
}

function themeTypeface(fontEl: Element | null, tag: string): string | undefined {
  const el = kid(fontEl, tag)
  const t = el?.textContent?.trim()
  return t || undefined
}

function parseTheme(themeXml?: string): ThemeFonts {
  if (!themeXml) return {}
  const doc = new DOMParser().parseFromString(themeXml, 'application/xml')
  const scheme = doc.getElementsByTagNameNS(A_NS, 'fontScheme')[0]
  if (!scheme) return {}
  const major = kid(scheme, 'majorFont')
  const minor = kid(scheme, 'minorFont')
  return {
    majorLatin: themeTypeface(major, 'latin'),
    minorLatin: themeTypeface(minor, 'latin'),
    majorEa: themeTypeface(major, 'ea') || '等线',
    minorEa: themeTypeface(minor, 'ea') || '等线',
    majorCs: themeTypeface(major, 'cs'),
    minorCs: themeTypeface(minor, 'cs'),
  }
}

function themeFont(theme: ThemeFonts, value: string | null): string | undefined {
  if (!value) return undefined
  switch (value) {
    case 'majorHAnsi':
      return theme.majorLatin
    case 'minorHAnsi':
      return theme.minorLatin
    case 'majorEastAsia':
      return theme.majorEa
    case 'minorEastAsia':
      return theme.minorEa
    case 'majorBidi':
      return theme.majorCs
    case 'minorBidi':
      return theme.minorCs
    default:
      return undefined
  }
}

function readRunFormat(rPr: Element | null, theme: ThemeFonts): RunFormat {
  if (!rPr) return {}
  const out: RunFormat = {}
  const fonts = kid(rPr, 'rFonts')
  if (fonts) {
    const ea = wAttr(fonts, 'eastAsia') || themeFont(theme, wAttr(fonts, 'eastAsiaTheme'))
    const ascii = wAttr(fonts, 'ascii') || themeFont(theme, wAttr(fonts, 'asciiTheme'))
    if (ea) out.eastAsiaFont = ea
    if (ascii) out.asciiFont = ascii
  }
  const sz = kid(rPr, 'sz')
  if (sz) {
    const v = Number(wAttr(sz, 'val'))
    if (!Number.isNaN(v) && v > 0) out.sizePt = v / 2
  }
  const b = kid(rPr, 'b')
  if (b) out.bold = boolAttrEl(b)
  const i = kid(rPr, 'i')
  if (i) out.italic = boolAttrEl(i)
  const u = kid(rPr, 'u')
  if (u) out.underline = (wAttr(u, 'val') ?? 'single') !== 'none'
  const color = kid(rPr, 'color')
  if (color) {
    const v = wAttr(color, 'val')
    if (v && v !== 'auto') out.color = v
  }
  return out
}

function readParaFormat(pPr: Element | null): ParaFormat {
  if (!pPr) return {}
  const out: ParaFormat = {}
  const jc = kid(pPr, 'jc')
  if (jc) out.alignment = wAttr(jc, 'val') ?? undefined
  const spacing = kid(pPr, 'spacing')
  if (spacing) {
    const line = wAttr(spacing, 'line')
    const rule = wAttr(spacing, 'lineRule')
    const before = wAttr(spacing, 'before')
    const after = wAttr(spacing, 'after')
    if (line !== null) out.lineTwips = Number(line)
    if (rule) out.lineRule = rule
    if (before !== null) out.beforeTwips = Number(before)
    if (after !== null) out.afterTwips = Number(after)
  }
  const ind = kid(pPr, 'ind')
  if (ind) {
    const firstChars = wAttr(ind, 'firstLineChars')
    const firstTwips = wAttr(ind, 'firstLine')
    const left = wAttr(ind, 'left')
    if (firstChars !== null) out.firstLineChars = Number(firstChars) / 100
    if (firstTwips !== null) out.firstLineTwips = Number(firstTwips)
    if (left !== null) out.leftIndentTwips = Number(left)
  }
  const pageBreak = kid(pPr, 'pageBreakBefore')
  if (pageBreak) out.pageBreakBefore = boolAttrEl(pageBreak)
  const outline = kid(pPr, 'outlineLvl')
  if (outline) {
    const v = Number(wAttr(outline, 'val'))
    if (!Number.isNaN(v)) out.outlineLevel = v
  }
  const keepNext = kid(pPr, 'keepNext')
  if (keepNext) out.keepNext = boolAttrEl(keepNext)
  const numPr = kid(pPr, 'numPr')
  if (numPr) out.hasNumbering = true
  return out
}

function mergeRun(...parts: (RunFormat | undefined)[]): RunFormat {
  const out: RunFormat = {}
  for (const p of parts) {
    if (!p) continue
    if (p.eastAsiaFont !== undefined) out.eastAsiaFont = p.eastAsiaFont
    if (p.asciiFont !== undefined) out.asciiFont = p.asciiFont
    if (p.sizePt !== undefined) out.sizePt = p.sizePt
    if (p.bold !== undefined) out.bold = p.bold
    if (p.italic !== undefined) out.italic = p.italic
    if (p.underline !== undefined) out.underline = p.underline
    if (p.color !== undefined) out.color = p.color
  }
  return out
}

function mergePara(...parts: (ParaFormat | undefined)[]): ParaFormat {
  const out: ParaFormat = {}
  for (const p of parts) {
    if (!p) continue
    if (p.alignment !== undefined) out.alignment = p.alignment
    if (p.lineTwips !== undefined) out.lineTwips = p.lineTwips
    if (p.lineRule !== undefined) out.lineRule = p.lineRule
    if (p.beforeTwips !== undefined) out.beforeTwips = p.beforeTwips
    if (p.afterTwips !== undefined) out.afterTwips = p.afterTwips
    if (p.firstLineChars !== undefined) out.firstLineChars = p.firstLineChars
    if (p.firstLineTwips !== undefined) out.firstLineTwips = p.firstLineTwips
    if (p.leftIndentTwips !== undefined) out.leftIndentTwips = p.leftIndentTwips
    if (p.pageBreakBefore !== undefined) out.pageBreakBefore = p.pageBreakBefore
    if (p.outlineLevel !== undefined) out.outlineLevel = p.outlineLevel
    if (p.keepNext !== undefined) out.keepNext = p.keepNext
    if (p.hasNumbering !== undefined) out.hasNumbering = p.hasNumbering
  }
  return out
}

function styleChain(styles: Map<string, StyleDef>, id: string | null | undefined): StyleDef[] {
  const chain: StyleDef[] = []
  const seen = new Set<string>()
  let cur = id
  while (cur && styles.has(cur) && !seen.has(cur)) {
    seen.add(cur)
    const def = styles.get(cur)!
    chain.unshift(def)
    cur = def.basedOn
  }
  return chain
}

function styleFormat(
  styles: Map<string, StyleDef>,
  id: string | undefined,
  theme: ThemeFonts,
): { para: ParaFormat; run: RunFormat; name: string } {
  const chain = styleChain(styles, id)
  let para: ParaFormat = {}
  let run: RunFormat = {}
  for (const def of chain) {
    para = mergePara(para, readParaFormat(def.pPr))
    run = mergeRun(run, readRunFormat(def.rPr, theme))
  }
  return { para, run, name: chain.length ? chain[chain.length - 1].name : 'Normal' }
}

function textOfRun(run: Element): string {
  let out = ''
  const walk = (el: Element) => {
    for (const child of Array.from(el.children)) {
      if (child.localName === 't') out += child.textContent ?? ''
      else if (child.localName === 'tab') out += '\t'
      else if (child.localName === 'br' || child.localName === 'cr') out += '\n'
      else if (child.localName === 'noBreakHyphen') out += '-'
      else if (child.localName === 'delText') continue
      else if (child.localName === 'instrText') continue
      else walk(child)
    }
  }
  walk(run)
  return out
}

function runHasDrawing(run: Element): boolean {
  return !!(
    run.getElementsByTagNameNS(W_NS, 'drawing').length ||
    run.getElementsByTagNameNS(W_NS, 'pict').length ||
    run.getElementsByTagNameNS(W_NS, 'object').length
  )
}

function collectFieldInstructions(p: Element): string[] {
  const out: string[] = []
  const instr = p.getElementsByTagNameNS(W_NS, 'instrText')
  for (const el of Array.from(instr)) out.push(el.textContent ?? '')
  const simple = p.getElementsByTagNameNS(W_NS, 'fldSimple')
  for (const el of Array.from(simple)) {
    const v = wAttr(el, 'instr')
    if (v) out.push(v)
  }
  return out
}

const HEADING_NAME_RE = /^(heading|标题)\s*([1-9])$/i
const HEADING_ID_RE = /^(heading|标题)([1-9])$/i

function headingLevelOf(
  text: string,
  styleName: string,
  styleId: string,
  format: ParaFormat,
): number | null {
  if (format.outlineLevel !== undefined && format.outlineLevel >= 0 && format.outlineLevel <= 8) {
    return format.outlineLevel + 1
  }
  const byName = styleName.match(HEADING_NAME_RE) || styleId.match(HEADING_ID_RE)
  if (byName) return Number(byName[2])
  const t = text.trim()
  if (!t || t.length > 40) return null
  if (/^第[一二三四五六七八九十百零〇\d]+章(\s|　|$)/.test(t)) return 1
  if (/^\d+\.\d+\.\d+(\s|　|$)/.test(t)) return 3
  if (/^\d+\.\d+(\s|　|$)/.test(t)) return 2
  if (
    /^\d+[.、\s　]/.test(t) &&
    !/[。，；：]/.test(t) &&
    !/^\d+\s*(年|月|日|项|分钟|秒|小时|元|人|个|次|页|倍|%|％)/.test(t)
  ) {
    return 1
  }
  if (/^(摘\s*要|目\s*录|参考文献|致\s*谢|谢\s*辞|附\s*录|引\s*言|结\s*论)$/.test(t)) return 1
  return null
}

const CAPTION_RE = /^(图|表|Fig\.?|Table)\s*([0-9一二三四五六七八九十]+(?:[-－—.．][0-9]+)?)(\s|　|[:：]|$)/i

function captionInfo(text: string): { kind: 'figure' | 'table'; number: string } | null {
  const t = text.trim()
  const m = t.match(CAPTION_RE)
  if (!m) return null
  if (t.length > 100) return null
  const kind = /^(图|Fig)/i.test(m[1]) ? 'figure' : 'table'
  return { kind, number: m[2].replace(/[．－—]/g, (c) => (c === '．' ? '.' : '-')) }
}

const TOC_STYLE_RE = /^(toc|目录)/i

function isTocEntry(text: string, styleName: string): boolean {
  if (TOC_STYLE_RE.test(styleName)) return true
  const t = text.trim()
  if (!t) return false
  return /(\.{4,}|…{2,}|·{4,})\s*\d+$/.test(t)
}

const cleanText = (s: string) => s.replace(/\u3000/g, ' ').replace(/\s+/g, ' ').trim()

const markerOf = (text: string): ParaCategory | null => {
  const t = cleanText(text)
  if (/^摘\s*要$/.test(t)) return 'abstract-title'
  if (/^abstract$/i.test(t)) return 'abstract-en-title'
  if (/^关键词(\s*[:：]|$)/.test(t)) return 'keywords'
  if (/^key\s*words?(\s*[:：]|$)/i.test(t)) return 'keywords'
  if (/^目\s*录$/.test(t)) return 'toc-title'
  if (/^参考文献$/.test(t)) return 'reference-title'
  if (/^(致\s*谢|谢\s*辞)$/.test(t)) return 'ack-title'
  if (/^附\s*录$/.test(t)) return 'appendix-title'
  return null
}

function zoneAfterMarker(category: ParaCategory, current: Zone): Zone {
  switch (category) {
    case 'abstract-title':
      return 'abstract'
    case 'abstract-en-title':
      return 'abstractEn'
    case 'toc-title':
      return 'toc'
    case 'reference-title':
      return 'references'
    case 'ack-title':
      return 'ack'
    case 'appendix-title':
      return 'appendix'
    default:
      return current
  }
}

function parseCore(coreXml?: string): CoreProps {
  if (!coreXml) return {}
  const doc = new DOMParser().parseFromString(coreXml, 'application/xml')
  const get = (local: string) => {
    const el = doc.getElementsByTagNameNS('*', local)[0]
    return el?.textContent?.trim() || undefined
  }
  return {
    title: get('title'),
    creator: get('creator'),
    lastModifiedBy: get('lastModifiedBy'),
    modified: get('modified'),
  }
}

function parseRels(relsXml?: string): Map<string, string> {
  const out = new Map<string, string>()
  if (!relsXml) return out
  const doc = new DOMParser().parseFromString(relsXml, 'application/xml')
  for (const rel of Array.from(doc.getElementsByTagNameNS(PKG_REL_NS, 'Relationship'))) {
    const id = rel.getAttribute('Id')
    const target = rel.getAttribute('Target')
    if (id && target) out.set(id, target)
  }
  return out
}

function normalizeZipPath(target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  return `word/${target.replace(/^\.\//, '')}`
}

async function readXmlText(zip: JSZip, path: string): Promise<string | undefined> {
  const f = zip.file(path)
  if (!f) return undefined
  return f.async('string')
}

function textFromHeaderFooter(xml?: string): { text: string; hasPage: boolean } {
  if (!xml) return { text: '', hasPage: false }
  const doc = new DOMParser().parseFromString(xml, 'application/xml')
  const texts = Array.from(doc.getElementsByTagNameNS(W_NS, 't'))
    .map((el) => el.textContent ?? '')
    .join(' ')
  let hasPage = false
  for (const el of Array.from(doc.getElementsByTagNameNS(W_NS, 'instrText'))) {
    if (/\bPAGE\b/.test(el.textContent ?? '')) hasPage = true
  }
  for (const el of Array.from(doc.getElementsByTagNameNS(W_NS, 'fldSimple'))) {
    if (/\bPAGE\b/.test(wAttr(el, 'instr') ?? '')) hasPage = true
  }
  return { text: cleanText(texts), hasPage }
}

export async function parseDocx(
  buffer: ArrayBuffer,
  fileName: string,
): Promise<ParsedDoc> {
  const zip = await JSZip.loadAsync(buffer)
  const docXml = await readXmlText(zip, 'word/document.xml')
  if (!docXml) throw new Error('这不是一个有效的 .docx 文件：缺少 word/document.xml')
  const stylesXml = await readXmlText(zip, 'word/styles.xml')
  const themeXml = await readXmlText(zip, 'word/theme/theme1.xml')
  const coreXml = await readXmlText(zip, 'docProps/core.xml')
  const relsXml = await readXmlText(zip, 'word/_rels/document.xml.rels')
  const numberingXml = await readXmlText(zip, 'word/numbering.xml')

  const document = new DOMParser().parseFromString(docXml, 'application/xml')
  if (document.getElementsByTagName('parsererror').length) {
    throw new Error('无法解析 document.xml（XML 结构异常）。')
  }
  const body = document.getElementsByTagNameNS(W_NS, 'body')[0]
  if (!body) throw new Error('文档缺少正文结构。')

  const theme = parseTheme(themeXml)
  const warnings: string[] = []
  if (!stylesXml) warnings.push('文档缺少 styles.xml，样式继承信息可能不完整。')

  let docDefaultsPara: ParaFormat = {}
  let docDefaultsRun: RunFormat = {}
  let defaultParaStyleId: string | undefined
  const styles = new Map<string, StyleDef>()
  if (stylesXml) {
    const sdoc = new DOMParser().parseFromString(stylesXml, 'application/xml')
    const defaults = sdoc.getElementsByTagNameNS(W_NS, 'docDefaults')[0]
    if (defaults) {
      docDefaultsPara = readParaFormat(kid(kid(defaults, 'pPrDefault'), 'pPr'))
      docDefaultsRun = readRunFormat(kid(kid(defaults, 'rPrDefault'), 'rPr'), theme)
    }
    for (const el of Array.from(sdoc.getElementsByTagNameNS(W_NS, 'style'))) {
      const id = wAttr(el, 'styleId')
      if (!id) continue
      const name = wAttr(kid(el, 'name'), 'val') ?? id
      const type = wAttr(el, 'type') ?? 'paragraph'
      const basedOn = wAttr(kid(el, 'basedOn'), 'val') ?? undefined
      if (type === 'paragraph' && wAttr(el, 'default') === '1') defaultParaStyleId = id
      styles.set(id, {
        id,
        name,
        type,
        basedOn,
        pPr: kid(el, 'pPr'),
        rPr: kid(el, 'rPr'),
      })
    }
  }
  if (!defaultParaStyleId && styles.has('Normal')) defaultParaStyleId = 'Normal'
  if (numberingXml) {
    // numbering definitions normally override indentation; we keep the flag only.
  }

  const rels = parseRels(relsXml)
  const headerCache = new Map<string, { text: string; hasPage: boolean }>()
  const footerCache = new Map<string, { text: string; hasPage: boolean }>()
  const readPart = async (
    target: string | undefined,
    cache: Map<string, { text: string; hasPage: boolean }>,
  ) => {
    if (!target) return { text: '', hasPage: false }
    const path = normalizeZipPath(target)
    if (cache.has(path)) return cache.get(path)!
    const xml = await readXmlText(zip, path)
    const parsed = textFromHeaderFooter(xml)
    cache.set(path, parsed)
    return parsed
  }

  const sectPrs = Array.from(document.getElementsByTagNameNS(W_NS, 'sectPr'))
  const sections: SectionInfo[] = []
  let inheritedHeader: string | undefined
  let inheritedFooter: string | undefined
  let inheritedHasHeader = false
  let inheritedHasFooter = false
  for (let i = 0; i < sectPrs.length; i++) {
    const sect = sectPrs[i]
    const pgSz = kid(sect, 'pgSz')
    const pgMar = kid(sect, 'pgMar')
    const margins: Partial<Margins> = {}
    if (pgMar) {
      const map: (keyof Margins)[] = ['top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter']
      for (const key of map) {
        const v = wAttr(pgMar, key)
        if (v !== null && v !== undefined) margins[key] = Number(v)
      }
    }
    const headerRefs = kids(sect, 'headerReference')
    const footerRefs = kids(sect, 'footerReference')
    const headerIds = headerRefs.map((ref) => rAttr(ref, 'id')).filter((id): id is string => !!id)
    const footerIds = footerRefs.map((ref) => rAttr(ref, 'id')).filter((id): id is string => !!id)
    const headerInfoList = headerIds.length
      ? await Promise.all(headerIds.map((id) => readPart(rels.get(id), headerCache)))
      : [await readPart(inheritedHeader ? rels.get(inheritedHeader) : undefined, headerCache)]
    const footerInfoList = footerIds.length
      ? await Promise.all(footerIds.map((id) => readPart(rels.get(id), footerCache)))
      : [await readPart(inheritedFooter ? rels.get(inheritedFooter) : undefined, footerCache)]
    if (headerIds.length) inheritedHeader = headerIds[0]
    if (footerIds.length) inheritedFooter = footerIds[0]
    if (headerRefs.length) inheritedHasHeader = true
    if (footerRefs.length) inheritedHasFooter = true
    sections.push({
      index: i,
      widthTwips: pgSz ? Number(wAttr(pgSz, 'w')) : undefined,
      heightTwips: pgSz ? Number(wAttr(pgSz, 'h')) : undefined,
      orientation: pgSz ? wAttr(pgSz, 'orient') ?? undefined : undefined,
      margins,
      startParaIndex: 0,
      hasPageNumberField: footerInfoList.some((info) => info.hasPage),
      hasFooterReference: inheritedHasFooter,
      hasHeaderReference: inheritedHasHeader,
      footerText: footerInfoList.map((info) => info.text).filter(Boolean).join(' / '),
      headerText: headerInfoList.map((info) => info.text).filter(Boolean).join(' / '),
    })
  }

  const paragraphs: ParaInfo[] = []
  const pEls = Array.from(body.getElementsByTagNameNS(W_NS, 'p'))
  let sectionCursor = 0
  for (let index = 0; index < pEls.length; index++) {
    const p = pEls[index]
    const inTable = !!closestLocal(p, 'tbl')
    const inTextBox = !!closestLocal(p, 'txbxContent')
    const pPr = kid(p, 'pPr')
    const pStyle = wAttr(kid(pPr, 'pStyle'), 'val') ?? defaultParaStyleId
    const styleFmt = styleFormat(styles, pStyle, theme)
    const directPara = readParaFormat(pPr)
    const paraFormat = mergePara(docDefaultsPara, styleFmt.para, directPara)
    const baseRun = mergeRun(docDefaultsRun, styleFmt.run)

    const runs: RunInfo[] = []
    let hasDrawing = false
    let hasPageBreak = false
    const fieldInstructions: string[] = collectFieldInstructions(p)
    for (const r of Array.from(p.getElementsByTagNameNS(W_NS, 'r'))) {
      if (closestLocal(r, 'txbxContent') && !inTextBox) {
        // Paragraph inside a text box: keep text but mark separately below.
      }
      const rPr = kid(r, 'rPr')
      const rStyle = wAttr(kid(rPr, 'rStyle'), 'val') ?? undefined
      const charStyle = rStyle ? styleFormat(styles, rStyle, theme) : undefined
      const fmt = mergeRun(baseRun, charStyle?.run, readRunFormat(rPr, theme))
      const text = textOfRun(r)
      const drawing = runHasDrawing(r)
      if (drawing) hasDrawing = true
      for (const br of Array.from(r.getElementsByTagNameNS(W_NS, 'br'))) {
        if ((wAttr(br, 'type') ?? '') === 'page') hasPageBreak = true
      }
      runs.push({
        text,
        format: fmt,
        hasDrawing: drawing,
        fieldInstruction: fieldInstructions.find((x) => x && text === '') ? undefined : undefined,
      })
    }
    const text = runs.map((r) => r.text).join('')
    const trimmed = text.trim()
    const styleName = pStyle ? styles.get(pStyle)?.name ?? pStyle : 'Normal'
    const styledHeading =
      (paraFormat.outlineLevel !== undefined && paraFormat.outlineLevel >= 0 && paraFormat.outlineLevel <= 8) ||
      HEADING_NAME_RE.test(styleName) ||
      HEADING_ID_RE.test(pStyle ?? '')
    const headingLevel = inTable || inTextBox ? null : headingLevelOf(trimmed, styleName, pStyle ?? '', paraFormat)
    const caption = inTable || inTextBox ? null : captionInfo(text)
    const hasTocField = fieldInstructions.some((x) => /\bTOC\b/.test(x))
    const lastSect = parasHasSectPr(pPr)
    const sectionIndex = sectionCursor
    if (lastSect) sectionCursor++
    const charCount = trimmed.replace(/\s/g, '').length

    paragraphs.push({
      index,
      text,
      styleId: pStyle ?? 'Normal',
      styleName,
      format: paraFormat,
      runs,
      baseRun,
      headingLevel: headingLevel && headingLevel <= 3 ? headingLevel : null,
      isHeadingStyle: styledHeading,
      isCaption: !!caption,
      captionKind: caption?.kind ?? null,
      captionNumber: caption?.number ?? null,
      inTable,
      inTextBox,
      hasDrawing,
      hasPageBreak,
      hasTocField,
      hasSectionBreak: !!lastSect,
      prevIsTable: false,
      nextIsTable: false,
      sectionIndex,
      category: 'body',
      zone: 'front',
      empty: trimmed.length === 0,
      charCount,
    })
  }

  // Table adjacency: used to check whether a table caption is above its table.
  const pIndexByElement = new Map<Element, number>()
  pEls.forEach((el, i) => pIndexByElement.set(el, i))
  const bodyChildren = Array.from(body.children)
  let lastBlock: Element | null = null
  for (const block of bodyChildren) {
    if (block.localName === 'p') {
      const idx = pIndexByElement.get(block)
      if (idx !== undefined) {
        paragraphs[idx].prevIsTable = lastBlock?.localName === 'tbl'
        if (!paragraphs[idx].empty || paragraphs[idx].hasDrawing) lastBlock = block
      }
    } else if (block.localName === 'tbl') {
      lastBlock = block
    }
  }
  lastBlock = null
  for (let i = bodyChildren.length - 1; i >= 0; i--) {
    const block = bodyChildren[i]
    if (block.localName === 'p') {
      const idx = pIndexByElement.get(block)
      if (idx !== undefined) {
        paragraphs[idx].nextIsTable = lastBlock?.localName === 'tbl'
        if (!paragraphs[idx].empty || paragraphs[idx].hasDrawing) lastBlock = block
      }
    } else if (block.localName === 'tbl') {
      lastBlock = block
    }
  }

  // Section start paragraph indexes
  for (const sec of sections) {
    const idx = paragraphs.findIndex((p) => p.sectionIndex === sec.index)
    sec.startParaIndex = idx >= 0 ? idx : 0
  }

  // Classify zones and categories in one pass.
  let zone: Zone = 'front'
  const hasAnyTocField = paragraphs.some((p) => p.hasTocField)
  for (const p of paragraphs) {
    if (p.empty) {
      p.category = 'empty'
      p.zone = zone
      continue
    }
    const marker = markerOf(p.text)
    if (marker) {
      p.category = marker
      zone = zoneAfterMarker(marker, zone)
      p.zone = zone
      if (p.headingLevel === null && ['abstract-title', 'abstract-en-title', 'toc-title', 'reference-title', 'ack-title', 'appendix-title'].includes(marker)) {
        p.headingLevel = 1
        p.isHeadingStyle = true
      }
      continue
    }
    p.zone = zone
    if (zone === 'front') {
      p.category = 'cover'
    } else if (zone === 'toc') {
      p.category = p.styleName.match(TOC_STYLE_RE) || isTocEntry(p.text, p.styleName) ? 'toc' : 'body'
    } else if (zone === 'references') {
      p.category = 'reference'
    } else if (zone === 'ack') {
      p.category = 'ack'
    } else if (zone === 'appendix') {
      p.category = 'appendix'
    } else if (zone === 'abstract' || zone === 'abstractEn') {
      p.category = marker === 'keywords' ? 'keywords' : zone === 'abstract' ? 'abstract' : 'abstract-en'
      if (p.headingLevel && p.headingLevel <= 3) p.category = 'heading'
    }
    if (p.category === 'body' && p.headingLevel) p.category = 'heading'
    if (p.category === 'body' && p.isCaption) p.category = 'caption'
  }
  // headings inside abstract zones (e.g. subheadings) are still headings
  for (const p of paragraphs) {
    if (p.headingLevel && ['abstract', 'abstractEn', 'body'].includes(p.category)) p.category = 'heading'
  }

  const bodyParas = paragraphs.filter(
    (p) =>
      p.category === 'body' &&
      !p.inTable &&
      !p.inTextBox &&
      p.charCount >= 10,
  )
  const bodyFontCount = new Map<string, number>()
  const bodySizeCount = new Map<string, number>()
  const spacingCount = new Map<string, number>()
  for (const p of bodyParas) {
    const f = majorityFont(p)
    if (f) bodyFontCount.set(f, (bodyFontCount.get(f) ?? 0) + 1)
    const s = majoritySize(p)
    if (s) bodySizeCount.set(formatSize(s), (bodySizeCount.get(formatSize(s)) ?? 0) + 1)
    const sp = spacingLabel(p.format)
    spacingCount.set(sp, (spacingCount.get(sp) ?? 0) + 1)
  }
  const toDistinct = (m: Map<string, number>): DistinctValue[] =>
    Array.from(m.entries())
      .map(([value, count]) => ({ value, count }))
      .sort((a, b) => b.count - a.count)

  const headingCounts: Record<number, number> = {}
  for (const p of paragraphs) {
    if (p.headingLevel) headingCounts[p.headingLevel] = (headingCounts[p.headingLevel] ?? 0) + 1
  }

  const stats: DocStats = {
    paragraphCount: paragraphs.length,
    bodyParagraphCount: bodyParas.length,
    bodyCharCount: bodyParas.reduce((a, p) => a + p.charCount, 0),
    distinctBodyFonts: toDistinct(bodyFontCount),
    distinctBodySizes: toDistinct(bodySizeCount),
    distinctBodySpacings: toDistinct(spacingCount),
    figureCaptionCount: paragraphs.filter((p) => p.captionKind === 'figure').length,
    tableCaptionCount: paragraphs.filter((p) => p.captionKind === 'table').length,
    referenceCount: paragraphs.filter((p) => p.category === 'reference').length,
    headingCounts,
  }

  return {
    fileName,
    fileSize: buffer.byteLength,
    paragraphs,
    sections,
    hasTocField: hasAnyTocField,
    styles,
    core: parseCore(coreXml),
    stats,
    warnings,
    buffer,
  }
}

function closestLocal(el: Element, localName: string): Element | null {
  let cur: Element | null = el.parentElement
  while (cur) {
    if (cur.localName === localName) return cur
    cur = cur.parentElement
  }
  return null
}

function parasHasSectPr(pPr: Element | null): boolean {
  return !!kid(pPr, 'sectPr')
}

export function majorityFont(p: ParaInfo): string | undefined {
  const counts = new Map<string, number>()
  for (const r of p.runs) {
    const chars = r.text.replace(/\s/g, '').length
    if (!chars) continue
    const font = r.format.eastAsiaFont || p.baseRun.eastAsiaFont || r.format.asciiFont || p.baseRun.asciiFont
    if (font) counts.set(font, (counts.get(font) ?? 0) + chars)
  }
  if (!counts.size) return p.baseRun.eastAsiaFont || p.baseRun.asciiFont
  return Array.from(counts.entries()).sort((a, b) => b[1] - a[1])[0][0]
}

export function majoritySize(p: ParaInfo): number | undefined {
  const counts = new Map<number, number>()
  for (const r of p.runs) {
    const chars = r.text.replace(/\s/g, '').length
    if (!chars) continue
    const size = r.format.sizePt ?? p.baseRun.sizePt
    if (size) counts.set(size, (counts.get(size) ?? 0) + chars)
  }
  if (!counts.size) return p.baseRun.sizePt
  return Array.from(counts.entries()).sort((a, b) => b[1] - a[1])[0][0]
}

export function majorityAsciiFont(p: ParaInfo): string | undefined {
  const counts = new Map<string, number>()
  for (const r of p.runs) {
    const chars = r.text.replace(/[^\x00-\x7F]/g, '').length
    if (!chars) continue
    const font = r.format.asciiFont || p.baseRun.asciiFont
    if (font) counts.set(font, (counts.get(font) ?? 0) + chars)
  }
  if (!counts.size) return p.baseRun.asciiFont
  return Array.from(counts.entries()).sort((a, b) => b[1] - a[1])[0][0]
}

export function boldRatio(p: ParaInfo): number {
  let boldChars = 0
  let total = 0
  for (const r of p.runs) {
    const chars = r.text.replace(/\s/g, '').length
    if (!chars) continue
    total += chars
    if (r.format.bold ?? p.baseRun.bold) boldChars += chars
  }
  return total ? boldChars / total : 0
}

export function fontMix(p: ParaInfo): string[] {
  const set = new Set<string>()
  for (const r of p.runs) {
    if (!r.text.replace(/\s/g, '').length) continue
    const f = r.format.eastAsiaFont || p.baseRun.eastAsiaFont
    if (f) set.add(normalizeFontName(f))
  }
  return Array.from(set)
}

export function sizeMix(p: ParaInfo): number[] {
  const set = new Set<number>()
  for (const r of p.runs) {
    if (!r.text.replace(/\s/g, '').length) continue
    const s = r.format.sizePt ?? p.baseRun.sizePt
    if (s) set.add(s)
  }
  return Array.from(set)
}

export function formatSize(pt?: number): string {
  if (!pt) return '未设置'
  const cn: Record<number, string> = {
    42: '初号',
    36: '小初',
    26: '一号',
    24: '小一',
    22: '二号',
    18: '小二',
    16: '三号',
    15: '小三',
    14: '四号',
    12: '小四',
    10.5: '五号',
    9: '小五',
  }
  const label = cn[Math.round(pt * 10) / 10]
  return label ? `${label} · ${formatNumber(pt)}pt` : `${formatNumber(pt)}pt`
}

export function formatNumber(n: number): string {
  return Number.isInteger(n) ? String(n) : String(Math.round(n * 10) / 10)
}

export function lineSpacingText(fmt: ParaFormat): string {
  if (fmt.lineTwips === undefined) return '未设置（单倍）'
  if (fmt.lineRule === 'exact') return `固定值 ${formatNumber(fmt.lineTwips / 20)} 磅`
  if (fmt.lineRule === 'atLeast') return `最小值 ${formatNumber(fmt.lineTwips / 20)} 磅`
  const multiple = fmt.lineTwips / 240
  if (Math.abs(multiple - 1) < 0.01) return '单倍行距'
  return `${formatNumber(multiple)} 倍行距`
}

export function alignmentText(a?: string): string {
  switch (a) {
    case 'both':
    case 'distribute':
      return '两端对齐'
    case 'center':
      return '居中'
    case 'right':
      return '右对齐'
    case 'left':
      return '左对齐'
    default:
      return '未设置（左对齐）'
  }
}

export function indentText(fmt: ParaFormat): string {
  if (fmt.firstLineChars !== undefined) return `首行缩进 ${formatNumber(fmt.firstLineChars)} 字符`
  if (fmt.firstLineTwips !== undefined) return `首行缩进 ${formatNumber(fmt.firstLineTwips / 20)} 磅`
  return '无首行缩进'
}

export function spacingLabel(fmt: ParaFormat): string {
  return lineSpacingText(fmt)
}

export function sectionLabel(sec: SectionInfo): string {
  const w = sec.widthTwips ? `${formatNumber(twipsToCm(sec.widthTwips))}×${formatNumber(twipsToCm(sec.heightTwips ?? 0))} cm` : '纸张未设置'
  return `第 ${sec.index + 1} 节 · ${w}`
}

export function marginLabel(sec: SectionInfo): string {
  const m = sec.margins
  const one = (v?: number) => (v === undefined ? '—' : `${formatNumber(twipsToCm(v))} cm`)
  return `上 ${one(m.top)} / 下 ${one(m.bottom)} / 左 ${one(m.left)} / 右 ${one(m.right)}`
}
