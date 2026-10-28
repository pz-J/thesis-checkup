import JSZip from 'jszip'
import { W_NS, type ParsedDoc } from './docx'

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
const XML_NS = 'http://www.w3.org/XML/1998/namespace'

const FOOTER_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
const FOOTER_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer'
const RELS_PATH = 'word/_rels/document.xml.rels'
const CT_PATH = '[Content_Types].xml'

const SECTPR_ORDER = [
  'headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar',
  'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign',
  'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid',
  'printerSettings', 'sectPrChange',
]

const parseXml = (xml: string) => new DOMParser().parseFromString(xml, 'application/xml')
const serialize = (node: Document | Element) => new XMLSerializer().serializeToString(node)
const withDeclaration = (xml: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml.replace(/^<\?xml[^>]*\?>\s*/i, '')}`

function hasPageField(xml: string): boolean {
  const doc = parseXml(xml)
  const instr = Array.from(doc.getElementsByTagNameNS(W_NS, 'instrText'))
  if (instr.some((el) => /\bPAGE\b/.test(el.textContent ?? ''))) return true
  const simple = Array.from(doc.getElementsByTagNameNS(W_NS, 'fldSimple'))
  return simple.some((el) => /\bPAGE\b/.test(el.getAttributeNS(W_NS, 'instr') ?? el.getAttribute('w:instr') ?? ''))
}

function buildPageFieldParagraph(doc: Document): Element {
  const w = (name: string) => doc.createElementNS(W_NS, `w:${name}`)

  const p = w('p')
  const pPr = w('pPr')
  const jc = w('jc')
  jc.setAttributeNS(W_NS, 'w:val', 'center')
  pPr.appendChild(jc)
  p.appendChild(pPr)

  const fldBegin = w('r')
  const beginChar = w('fldChar')
  beginChar.setAttributeNS(W_NS, 'w:fldCharType', 'begin')
  fldBegin.appendChild(beginChar)
  p.appendChild(fldBegin)

  const instrRun = w('r')
  const instr = w('instrText')
  instr.setAttributeNS(XML_NS, 'xml:space', 'preserve')
  instr.textContent = ' PAGE '
  instrRun.appendChild(instr)
  p.appendChild(instrRun)

  const sepRun = w('r')
  const sepChar = w('fldChar')
  sepChar.setAttributeNS(W_NS, 'w:fldCharType', 'separate')
  sepRun.appendChild(sepChar)
  p.appendChild(sepRun)

  const textRun = w('r')
  const rPr = w('rPr')
  const rFonts = w('rFonts')
  rFonts.setAttributeNS(W_NS, 'w:ascii', 'Times New Roman')
  rFonts.setAttributeNS(W_NS, 'w:hAnsi', 'Times New Roman')
  rFonts.setAttributeNS(W_NS, 'w:eastAsia', '宋体')
  rPr.appendChild(rFonts)
  const sz = w('sz')
  sz.setAttributeNS(W_NS, 'w:val', '21')
  rPr.appendChild(sz)
  const szCs = w('szCs')
  szCs.setAttributeNS(W_NS, 'w:val', '21')
  rPr.appendChild(szCs)
  textRun.appendChild(rPr)
  const t = w('t')
  t.textContent = '1'
  textRun.appendChild(t)
  p.appendChild(textRun)

  const endRun = w('r')
  const endChar = w('fldChar')
  endChar.setAttributeNS(W_NS, 'w:fldCharType', 'end')
  endRun.appendChild(endChar)
  p.appendChild(endRun)

  return p
}

function footerXmlWithPageField(): string {
  const doc = new DOMParser().parseFromString(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="${W_NS}" xmlns:r="${R_NS}"></w:ftr>`,
    'application/xml',
  )
  const root = doc.documentElement
  root.appendChild(buildPageFieldParagraph(doc))
  return withDeclaration(serialize(root))
}

function appendPageFieldToFooter(xml: string): string {
  const doc = parseXml(xml)
  const root = doc.documentElement
  root.appendChild(buildPageFieldParagraph(doc))
  return withDeclaration(serialize(root))
}

interface FooterPart {
  rid: string
  path: string
}

async function readFooterParts(zip: JSZip): Promise<FooterPart[]> {
  const file = zip.file(RELS_PATH)
  if (!file) return []
  const doc = parseXml(await file.async('string'))
  const parts: FooterPart[] = []
  for (const rel of Array.from(doc.getElementsByTagNameNS(REL_NS, 'Relationship'))) {
    if (!(rel.getAttribute('Type') ?? '').endsWith('/footer')) continue
    const id = rel.getAttribute('Id')
    const target = rel.getAttribute('Target')
    if (!id || !target) continue
    const path = target.startsWith('/') ? target.slice(1) : `word/${target.replace(/^\.\//, '')}`
    if (zip.file(path)) parts.push({ rid: id, path })
  }
  return parts
}

function addContentTypeOverride(zip: JSZip, partPath: string): Promise<void> | void {
  const file = zip.file(CT_PATH)
  if (!file) return
  return file.async('string').then((xml) => {
    const doc = parseXml(xml)
    const root = doc.documentElement
    const partName = `/${partPath}`
    const exists = Array.from(doc.getElementsByTagNameNS(CT_NS, 'Override')).some(
      (el) => el.getAttribute('PartName') === partName,
    )
    if (exists) return
    const override = doc.createElementNS(CT_NS, 'Override')
    override.setAttribute('PartName', partName)
    override.setAttribute('ContentType', FOOTER_CONTENT_TYPE)
    root.appendChild(override)
    zip.file(CT_PATH, withDeclaration(serialize(doc)))
  })
}

async function addRelationship(zip: JSZip, target: string): Promise<string> {
  const file = zip.file(RELS_PATH)
  if (!file) throw new Error('文档缺少关系表，无法插入页码。')
  const xml = await file.async('string')
  const doc = parseXml(xml)
  const root = doc.documentElement
  const ids = new Set(
    Array.from(doc.getElementsByTagNameNS(REL_NS, 'Relationship')).map((el) => el.getAttribute('Id') ?? ''),
  )
  let n = 1
  while (ids.has(`rId${n}`)) n++
  const rid = `rId${n}`
  const rel = doc.createElementNS(REL_NS, 'Relationship')
  rel.setAttribute('Id', rid)
  rel.setAttribute('Type', FOOTER_REL_TYPE)
  rel.setAttribute('Target', target)
  root.appendChild(rel)
  zip.file(RELS_PATH, withDeclaration(serialize(doc)))
  return rid
}

function nextFooterPath(zip: JSZip): string {
  let n = 1
  while (zip.file(`word/footer-checkup${n === 1 ? '' : n}.xml`)) n++
  return `word/footer-checkup${n === 1 ? '' : n}.xml`
}

function addFooterReference(sect: Element, rid: string, doc: Document) {
  const ref = doc.createElementNS(W_NS, 'w:footerReference')
  ref.setAttributeNS(W_NS, 'w:type', 'default')
  ref.setAttributeNS(R_NS, 'r:id', rid)
  const idx = SECTPR_ORDER.indexOf('footerReference')
  const next = Array.from(sect.children).find((child) => {
    const i = SECTPR_ORDER.indexOf(child.localName)
    return i > idx && idx >= 0
  })
  if (next) sect.insertBefore(ref, next)
  else sect.appendChild(ref)
}

export interface PageNumberFixResult {
  changed: number
  createdPart?: string
}

/**
 * 给缺少自动页码的节补上 PAGE 域。
 * - 已有页脚但没有 PAGE 域：直接在该页脚追加一段居中页码。
 * - 没有任何页脚：创建一个共享页脚部件，并把它引用到这些节。
 */
export async function ensurePageNumbers(
  zip: JSZip,
  documentDom: Document,
  doc: ParsedDoc,
): Promise<PageNumberFixResult> {
  const sectPrs = Array.from(documentDom.getElementsByTagNameNS(W_NS, 'sectPr'))
  const footerParts = await readFooterParts(zip)
  const edited = new Set<string>()
  let reusable: FooterPart | undefined
  let createdRid: string | undefined
  let createdPath: string | undefined
  let changed = 0

  for (let i = 0; i < sectPrs.length; i++) {
    const sect = sectPrs[i]
    const section = doc.sections[i]
    if (!section || section.hasPageNumberField) continue

    const refs = Array.from(sect.children).filter((c) => c.localName === 'footerReference')
    const ids = refs
      .map((ref) => ref.getAttributeNS(R_NS, 'id') ?? ref.getAttribute('r:id') ?? '')
      .filter(Boolean)

    let fixedThisSection = false

    for (const id of ids) {
      const part = footerParts.find((p) => p.rid === id)
      if (!part) continue
      if (edited.has(part.path)) {
        fixedThisSection = true
        continue
      }
      const file = zip.file(part.path)
      if (!file) continue
      const xml = await file.async('string')
      if (hasPageField(xml)) {
        fixedThisSection = true
        continue
      }
      zip.file(part.path, appendPageFieldToFooter(xml))
      edited.add(part.path)
      fixedThisSection = true
    }

    if (fixedThisSection) {
      changed++
      continue
    }

    // 没有可用页脚：优先复用文档里已经带 PAGE 域的页脚，否则新建一个。
    if (!reusable) {
      for (const part of footerParts) {
        const file = zip.file(part.path)
        if (!file) continue
        const xml = await file.async('string')
        if (hasPageField(xml)) {
          reusable = part
          break
        }
      }
    }

    if (reusable) {
      addFooterReference(sect, reusable.rid, documentDom)
      changed++
      continue
    }

    if (!createdRid) {
      const path = nextFooterPath(zip)
      zip.file(path, footerXmlWithPageField())
      const target = path.replace(/^word\//, '')
      createdRid = await addRelationship(zip, target)
      await addContentTypeOverride(zip, path)
      createdPath = path
    }
    addFooterReference(sect, createdRid, documentDom)
    changed++
  }

  return { changed, createdPart: createdPath }
}
