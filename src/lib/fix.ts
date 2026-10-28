import JSZip from 'jszip'
import { W_NS, type ParsedDoc } from './docx'
import { cmToTwips, type RuleProfile } from './profiles'
import { ensurePageNumbers } from './fixPageNumbers'

const PPR_ORDER = [
  'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
  'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
  'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
  'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap',
  'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId',
  'cnfStyle', 'rPr', 'sectPr', 'pPrChange',
]

const SECTPR_ORDER = [
  'headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar',
  'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign',
  'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid',
  'printerSettings', 'sectPrChange',
]

const RPR_ORDER = [
  'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike',
  'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden',
  'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect',
  'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout',
  'specVanish', 'oMath',
]

const kid = (el: Element | null | undefined, localName: string): Element | null =>
  el ? Array.from(el.children).find((c) => c.localName === localName) ?? null : null

const orderedInsert = (parent: Element, child: Element, order: string[]) => {
  const idx = order.indexOf(child.localName)
  const next = Array.from(parent.children).find((c) => {
    const i = order.indexOf(c.localName)
    return i > idx && idx >= 0
  })
  if (next) parent.insertBefore(child, next)
  else parent.appendChild(child)
}

const setChild = (parent: Element, name: string, attrs: Record<string, string>, order: string[]): Element => {
  let el = kid(parent, name)
  if (!el) {
    el = parent.ownerDocument.createElementNS(W_NS, `w:${name}`)
    orderedInsert(parent, el, order)
  }
  for (const [k, v] of Object.entries(attrs)) el.setAttributeNS(W_NS, `w:${k}`, v)
  return el
}

const ensurePPr = (p: Element): Element => {
  let pPr = kid(p, 'pPr')
  if (!pPr) {
    pPr = p.ownerDocument.createElementNS(W_NS, 'w:pPr')
    p.insertBefore(pPr, p.firstChild)
  }
  return pPr
}

const ensureRPr = (r: Element): Element => {
  let rPr = kid(r, 'rPr')
  if (!rPr) {
    rPr = r.ownerDocument.createElementNS(W_NS, 'w:rPr')
    r.insertBefore(rPr, r.firstChild)
  }
  return rPr
}

const setRunFont = (r: Element, opts: { eastAsia?: string; ascii?: string; sizePt?: number; bold?: boolean }) => {
  const rPr = ensureRPr(r)
  if (opts.eastAsia || opts.ascii) {
    const rFonts = setChild(rPr, 'rFonts', {}, RPR_ORDER)
    if (opts.eastAsia) {
      rFonts.setAttributeNS(W_NS, 'w:eastAsia', opts.eastAsia)
      rFonts.removeAttributeNS(W_NS, 'eastAsiaTheme')
    }
    if (opts.ascii) {
      rFonts.setAttributeNS(W_NS, 'w:ascii', opts.ascii)
      rFonts.setAttributeNS(W_NS, 'w:hAnsi', opts.ascii)
      rFonts.setAttributeNS(W_NS, 'w:cs', opts.ascii)
      rFonts.removeAttributeNS(W_NS, 'asciiTheme')
      rFonts.removeAttributeNS(W_NS, 'hAnsiTheme')
      rFonts.removeAttributeNS(W_NS, 'cstheme')
    }
  }
  if (opts.sizePt !== undefined) {
    const hp = String(Math.round(opts.sizePt * 2))
    setChild(rPr, 'sz', { val: hp }, RPR_ORDER)
    setChild(rPr, 'szCs', { val: hp }, RPR_ORDER)
  }
  if (opts.bold !== undefined) {
    setChild(rPr, 'b', { val: opts.bold ? '1' : '0' }, RPR_ORDER)
    setChild(rPr, 'bCs', { val: opts.bold ? '1' : '0' }, RPR_ORDER)
  }
}

const setParaSpacing = (pPr: Element, lineTwips: number, rule: 'exact' | 'auto', zeroBeforeAfter: boolean) => {
  const attrs: Record<string, string> = {
    line: String(lineTwips),
    lineRule: rule,
  }
  if (zeroBeforeAfter) {
    attrs.before = '0'
    attrs.after = '0'
  }
  setChild(pPr, 'spacing', attrs, PPR_ORDER)
}

const setParaIndent = (pPr: Element, chars: number) => {
  const ind = setChild(pPr, 'ind', { firstLineChars: String(Math.round(chars * 100)) }, PPR_ORDER)
  ind.removeAttributeNS(W_NS, 'firstLine')
  ind.removeAttributeNS(W_NS, 'hanging')
  ind.removeAttributeNS(W_NS, 'hangingChars')
  ind.setAttributeNS(W_NS, 'w:firstLineChars', String(Math.round(chars * 100)))
}

const setParaAlign = (pPr: Element, value: string) => {
  setChild(pPr, 'jc', { val: value }, PPR_ORDER)
}

const setPageBreakBefore = (pPr: Element) => {
  const pb = setChild(pPr, 'pageBreakBefore', {}, PPR_ORDER)
  pb.removeAttributeNS(W_NS, 'val')
}

const trimLeadingSpaces = (p: Element) => {
  let trimming = true
  const texts = Array.from(p.getElementsByTagNameNS(W_NS, 't'))
  for (const t of texts) {
    if (!trimming) break
    const old = t.textContent ?? ''
    const trimmed = old.replace(/^[\u3000 ]+/, '')
    if (trimmed !== old) {
      t.textContent = trimmed
      t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve')
    }
    if (trimmed.length) trimming = false
  }
}

const allRuns = (p: Element): Element[] => Array.from(p.getElementsByTagNameNS(W_NS, 'r'))

const isDrawingRun = (r: Element): boolean =>
  !!(r.getElementsByTagNameNS(W_NS, 'drawing').length || r.getElementsByTagNameNS(W_NS, 'pict').length)

function buildTocFieldParagraph(owner: Document): Element {
  const w = (name: string) => owner.createElementNS(W_NS, `w:${name}`)
  const p = w('p')
  const pPr = w('pPr')
  const jc = w('jc')
  jc.setAttributeNS(W_NS, 'w:val', 'left')
  pPr.appendChild(jc)
  p.appendChild(pPr)

  const beginRun = w('r')
  const begin = w('fldChar')
  begin.setAttributeNS(W_NS, 'w:fldCharType', 'begin')
  begin.setAttributeNS(W_NS, 'w:dirty', 'true')
  beginRun.appendChild(begin)
  p.appendChild(beginRun)

  const instrRun = w('r')
  const instr = w('instrText')
  instr.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve')
  instr.textContent = ' TOC \\o "1-3" \\h \\z \\u '
  instrRun.appendChild(instr)
  p.appendChild(instrRun)

  const sepRun = w('r')
  const sep = w('fldChar')
  sep.setAttributeNS(W_NS, 'w:fldCharType', 'separate')
  sepRun.appendChild(sep)
  p.appendChild(sepRun)

  const resultRun = w('r')
  const resultPr = w('rPr')
  const fonts = w('rFonts')
  fonts.setAttributeNS(W_NS, 'w:ascii', 'Times New Roman')
  fonts.setAttributeNS(W_NS, 'w:hAnsi', 'Times New Roman')
  fonts.setAttributeNS(W_NS, 'w:eastAsia', '宋体')
  resultPr.appendChild(fonts)
  const sz = w('sz')
  sz.setAttributeNS(W_NS, 'w:val', '24')
  resultPr.appendChild(sz)
  resultRun.appendChild(resultPr)
  const resultText = w('t')
  resultText.textContent = '右键更新目录'
  resultRun.appendChild(resultText)
  p.appendChild(resultRun)

  const endRun = w('r')
  const end = w('fldChar')
  end.setAttributeNS(W_NS, 'w:fldCharType', 'end')
  endRun.appendChild(end)
  p.appendChild(endRun)

  return p
}

function rebuildToc(pEls: Element[], doc: ParsedDoc): number {
  const title = doc.paragraphs.find((p) => p.category === 'toc-title')
  if (!title) return 0
  const entries = doc.paragraphs.filter(
    (p) => p.category === 'toc' && !p.empty && p.index > title.index,
  )
  if (!entries.length) return 0
  const anchor = pEls[entries[0].index]
  if (!anchor?.parentNode) return 0
  const field = buildTocFieldParagraph(anchor.ownerDocument)
  anchor.parentNode.insertBefore(field, anchor)
  for (const entry of entries) {
    const el = pEls[entry.index]
    if (el?.parentNode) el.parentNode.removeChild(el)
  }
  return entries.length
}

export interface FixResult {
  buffer: ArrayBuffer
  applied: string[]
  changes: Array<{ ruleId: string; count: number }>
}

export async function applyFixes(
  source: ArrayBuffer,
  doc: ParsedDoc,
  profile: RuleProfile,
  selectedRules: string[],
): Promise<FixResult> {
  const zip = await JSZip.loadAsync(source)
  const xml = await zip.file('word/document.xml')!.async('string')
  const parsed = new DOMParser().parseFromString(xml, 'application/xml')
  const body = parsed.getElementsByTagNameNS(W_NS, 'body')[0]
  if (!body) throw new Error('无法读取文档正文。')

  const pEls = Array.from(body.getElementsByTagNameNS(W_NS, 'p'))
  const selected = new Set(selectedRules)
  const changes = new Map<string, number>()
  const touch = (ruleId: string, n = 1) => changes.set(ruleId, (changes.get(ruleId) ?? 0) + n)

  // 页面设置
  const sectPrs = Array.from(parsed.getElementsByTagNameNS(W_NS, 'sectPr'))
  if (selected.has('page.size')) {
    for (const sect of sectPrs) {
      const pgSz = setChild(sect, 'pgSz', {
        w: String(profile.page.widthTwips),
        h: String(profile.page.heightTwips),
      }, SECTPR_ORDER)
      pgSz.removeAttributeNS(W_NS, 'orient')
      touch('page.size')
    }
  }
  if (selected.has('page.margins')) {
    for (const sect of sectPrs) {
      const m = profile.page.marginsCm
      setChild(sect, 'pgMar', {
        top: String(cmToTwips(m.top)),
        right: String(cmToTwips(m.right)),
        bottom: String(cmToTwips(m.bottom)),
        left: String(cmToTwips(m.left)),
      }, SECTPR_ORDER)
      touch('page.margins')
    }
  }
  if (selected.has('page.pageNumber')) {
    try {
      const pageFix = await ensurePageNumbers(zip, parsed, doc)
      if (pageFix.changed > 0) touch('page.pageNumber', pageFix.changed)
    } catch (e) {
      console.warn('页码自动修复失败：', e)
    }
  }

  if (selected.has('structure.tocField')) {
    const removed = rebuildToc(pEls, doc)
    if (removed > 0) touch('structure.tocField', removed)
  }

  const paraByIndex = new Map(doc.paragraphs.map((p) => [p.index, p]))
  const bodyCategories = new Set(['body', 'abstract', 'abstract-en'])
  const firstBodyH1 = doc.paragraphs.find(
    (p) => p.category === 'heading' && p.headingLevel === 1 && p.zone === 'body',
  )?.index

  pEls.forEach((p, index) => {
    const info = paraByIndex.get(index)
    if (!info || info.empty) return
    const pPr = ensurePPr(p)
    const runs = allRuns(p).filter((r) => !isDrawingRun(r))

    if (bodyCategories.has(info.category) && !info.inTable && !info.inTextBox && info.charCount >= 5) {
      if (selected.has('body.font') && info.category !== 'abstract-en') {
        for (const r of runs) setRunFont(r, { eastAsia: profile.body.eastAsiaFont })
        touch('body.font')
      }
      if (selected.has('body.ascii')) {
        for (const r of runs) setRunFont(r, { ascii: profile.body.asciiFont })
        touch('body.ascii')
      }
      if (selected.has('body.size')) {
        for (const r of runs) setRunFont(r, { sizePt: profile.body.sizePt })
        touch('body.size')
      }
      if (selected.has('body.spacing')) {
        setParaSpacing(pPr, profile.body.line.twips, profile.body.line.rule, false)
        touch('body.spacing')
      }
      if (selected.has('body.indent') && info.category === 'body' && info.charCount >= 15) {
        setParaIndent(pPr, profile.body.firstLineChars)
        touch('body.indent')
      }
      if (selected.has('body.align')) {
        setParaAlign(pPr, profile.body.alignment)
        touch('body.align')
      }
      if (selected.has('body.beforeAfter')) {
        const spacing = kid(pPr, 'spacing')
        if (spacing) {
          spacing.setAttributeNS(W_NS, 'w:before', '0')
          spacing.setAttributeNS(W_NS, 'w:after', '0')
        } else {
          setParaSpacing(pPr, profile.body.line.twips, profile.body.line.rule, true)
        }
        touch('body.beforeAfter')
      }
      if (selected.has('body.bold')) {
        for (const r of runs) setRunFont(r, { bold: false })
        touch('body.bold')
      }
      if (selected.has('body.leadingSpace') && /^(\u3000|\s{2,})/.test(info.text)) {
        trimLeadingSpaces(p)
        touch('body.leadingSpace')
      }
    }

    if (info.category === 'heading' && info.headingLevel) {
      const level = info.headingLevel as 1 | 2 | 3
      const target = profile.heading[level]
      if (target && selected.has('heading.font')) {
        for (const r of runs) setRunFont(r, { eastAsia: target.font })
        touch('heading.font')
      }
      if (target && selected.has('heading.size')) {
        for (const r of runs) setRunFont(r, { sizePt: target.sizePt })
        touch('heading.size')
      }
      if (level === 1 && selected.has('heading.pageBreak') && index !== firstBodyH1) {
        setPageBreakBefore(pPr)
        touch('heading.pageBreak')
      }
    }

    if (info.category === 'caption' && selected.has('caption.format')) {
      for (const r of runs) setRunFont(r, { eastAsia: profile.caption.font, sizePt: profile.caption.sizePt })
      setParaAlign(pPr, 'center')
      touch('caption.format')
    }

    if (info.category === 'reference-title') {
      if (selected.has('reference.title')) {
        for (const r of runs) setRunFont(r, { eastAsia: profile.reference.titleFont, sizePt: profile.reference.titleSizePt })
        setParaAlign(pPr, 'left')
        touch('reference.title')
      }
      if (selected.has('reference.pageBreak')) {
        setPageBreakBefore(pPr)
        touch('reference.pageBreak')
      }
    }

    if (info.category === 'reference' && selected.has('reference.font')) {
      for (const r of runs) {
        setRunFont(r, {
          eastAsia: profile.reference.font,
          ascii: profile.reference.asciiFont,
          sizePt: profile.reference.sizePt,
        })
      }
      touch('reference.font')
    }

    if (info.category === 'ack-title' && selected.has('structure.ackPageBreak')) {
      setPageBreakBefore(pPr)
      touch('structure.ackPageBreak')
    }

    if (info.category === 'appendix-title' && selected.has('structure.appendixPageBreak')) {
      setPageBreakBefore(pPr)
      touch('structure.appendixPageBreak')
    }

    const specialTargets: Array<[string, { font: string; sizePt: number; alignment: 'center'; bold?: boolean } | undefined]> = [
      ['special.abstractTitle', info.category === 'abstract-title' ? profile.special.abstractTitle : undefined],
      ['special.enAbstractTitle', info.category === 'abstract-en-title' ? profile.special.enAbstractTitle : undefined],
      ['special.tocTitle', info.category === 'toc-title' ? profile.special.tocTitle : undefined],
      ['special.ackTitle', info.category === 'ack-title' ? profile.special.ackTitle : undefined],
      ['special.appendixTitle', info.category === 'appendix-title' ? profile.special.appendixTitle : undefined],
    ]
    for (const [ruleId, target] of specialTargets) {
      if (!target || !selected.has(ruleId)) continue
      for (const r of runs) {
        setRunFont(r, {
          eastAsia: target.font,
          ascii: target.font,
          sizePt: target.sizePt,
          bold: target.bold,
        })
      }
      setParaAlign(pPr, target.alignment)
      touch(ruleId)
    }
  })

  const out = new XMLSerializer().serializeToString(parsed)
  zip.file('word/document.xml', out)
  const buffer = await zip.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  })

  return {
    buffer,
    applied: Array.from(changes.keys()),
    changes: Array.from(changes.entries()).map(([ruleId, count]) => ({ ruleId, count })),
  }
}

export function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = fileName
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 4000)
}

export function fixedFileName(original: string): string {
  const base = original.replace(/\.docx$/i, '')
  const d = new Date()
  const stamp = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`
  return `${base}_已修复_${stamp}.docx`
}
