// @vitest-environment jsdom
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/lib/docx'
import { applyFixes } from '../src/lib/fix'
import { TJU_PROFILE, parseProfile } from '../src/lib/profiles'
import { evaluate } from '../src/lib/rules'

function sampleBuffer(): ArrayBuffer {
  const b = readFileSync(resolve(process.cwd(), 'public/sample-thesis.docx'))
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer
}

function cleanBuffer(): ArrayBuffer {
  const b = readFileSync(resolve(process.cwd(), 'public/sample-clean.docx'))
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer
}

function fixtureBuffer(name: string): ArrayBuffer {
  const b = readFileSync(resolve(process.cwd(), 'tests/fixtures', name))
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer
}

describe('sample thesis', () => {
  it('parses structure and finds the seeded issues', async () => {
    const doc = await parseDocx(sampleBuffer(), 'sample-thesis.docx')
    console.log('paragraphs', doc.paragraphs.length, 'sections', doc.sections.length)
    console.log('stats', doc.stats)
    console.log(
      'zones',
      doc.paragraphs
        .filter((p) => !p.empty)
        .map((p) => `${p.index}:${p.category}:${p.headingLevel ?? '-'}:${p.text.slice(0, 18)}`)
        .join('\n'),
    )
    const report = evaluate(doc, TJU_PROFILE)
    console.log('score', report.score, report.grade, report.counts)
    console.log(
      'findings',
      report.findings.map((f) => ({
        id: f.ruleId,
        sev: f.severity,
        n: f.occurrences.length,
        actual: f.occurrences[0]?.actual,
        text: f.occurrences[0]?.text?.slice(0, 30),
      })),
    )
    expect(doc.paragraphs.length).toBeGreaterThan(30)
    expect(doc.stats.bodyParagraphCount).toBeGreaterThan(5)
    expect(report.findings.some((f) => f.ruleId === 'page.margins')).toBe(true)
    expect(report.findings.some((f) => f.ruleId === 'body.font')).toBe(true)
    expect(report.findings.some((f) => f.ruleId === 'body.size')).toBe(true)
    expect(report.findings.some((f) => f.ruleId === 'body.spacing')).toBe(true)
    expect(report.findings.some((f) => f.ruleId === 'body.indent')).toBe(true)
    expect(report.findings.some((f) => f.ruleId === 'reference.numbering')).toBe(true)
    expect(report.findings.some((f) => f.ruleId === 'reference.citationMapping')).toBe(true)
    expect(report.findings.some((f) => f.ruleId === 'structure.tocField')).toBe(true)
    expect(report.findings.some((f) => f.ruleId === 'page.pageNumber')).toBe(true)
  })

  it('repairs deterministic issues and improves the score', async () => {
    const buf = sampleBuffer()
    const doc = await parseDocx(buf, 'sample-thesis.docx')
    const before = evaluate(doc, TJU_PROFILE)
    const fixable = before.findings.filter((f) => f.fixable).map((f) => f.ruleId)
    const result = await applyFixes(buf, doc, TJU_PROFILE, fixable)
    console.log('applied', result.applied, result.changes)
    const afterDoc = await parseDocx(result.buffer, 'fixed.docx')
    const after = evaluate(afterDoc, TJU_PROFILE)
    console.log('score before/after', before.score, after.score, after.findings.map((f) => f.ruleId))
    expect(after.score).toBeGreaterThan(before.score)
    expect(after.score).toBeGreaterThan(94)
    expect(after.findings.some((f) => ['page.margins', 'body.font', 'body.size', 'body.spacing', 'body.indent'].includes(f.ruleId))).toBe(false)
    expect(after.findings.some((f) => f.ruleId === 'caption.format')).toBe(false)
    expect(after.findings.some((f) => f.ruleId === 'page.pageNumber')).toBe(false)
    expect(after.findings.some((f) => f.ruleId === 'structure.tocField')).toBe(false)
    expect(afterDoc.hasTocField).toBe(true)
    expect(afterDoc.sections.every((s) => s.hasPageNumberField)).toBe(true)

    const zip = await JSZip.loadAsync(result.buffer)
    const footerPath = Object.keys(zip.files).find((p) => /^word\/footer[^/]*\.xml$/i.test(p))
    expect(footerPath).toBeTruthy()
    const footerXml = await zip.file(footerPath!)!.async('string')
    expect(footerXml).toContain('PAGE')
    const contentTypes = await zip.file('[Content_Types].xml')!.async('string')
    expect(contentTypes).toContain('footer+xml')
  })

  it('does not produce false positives on a clean thesis', async () => {
    const doc = await parseDocx(cleanBuffer(), 'sample-clean.docx')
    const report = evaluate(doc, TJU_PROFILE)
    console.log(
      'clean score',
      report.score,
      report.findings.map((f) => ({ id: f.ruleId, actual: f.occurrences[0]?.actual, text: f.occurrences[0]?.text?.slice(0, 30) })),
    )
    expect(report.score).toBeGreaterThanOrEqual(98)
    expect(report.findings).toHaveLength(0)
  })

  it('normalizes a custom rule pack and falls back on bad values', () => {
    const { profile, error } = parseProfile({
      name: '某校本科论文',
      shortName: '某校',
      page: { marginsCm: { left: 3.0, right: 2.0 } },
      body: { sizePt: 14, eastAsiaFont: '仿宋' },
      heading: { 1: { sizePt: 18 } },
      structure: { requireEnglishAbstract: false },
    })
    expect(error).toBeUndefined()
    expect(profile).toBeTruthy()
    expect(profile!.page.marginsCm.left).toBe(3)
    expect(profile!.page.marginsCm.right).toBe(2)
    expect(profile!.page.marginsCm.top).toBe(TJU_PROFILE.page.marginsCm.top)
    expect(profile!.body.sizePt).toBe(14)
    expect(profile!.body.eastAsiaFont).toBe('仿宋')
    expect(profile!.heading[1].sizePt).toBe(18)
    expect(profile!.structure.requireEnglishAbstract).toBe(false)
    expect(profile!.id).not.toBe(TJU_PROFILE.id)

    const roundTrip = parseProfile(profile).profile
    expect(roundTrip!.id).toBe(profile!.id)
  })

  it('rejects a rule pack without a name', () => {
    const { profile, error } = parseProfile({ body: { sizePt: 12 } })
    expect(profile).toBeUndefined()
    expect(error).toContain('name')
  })

  it('appends a page number to an existing footer without losing its text', async () => {
    const buffer = fixtureBuffer('existing-footer.docx')
    const doc = await parseDocx(buffer, 'existing-footer.docx')
    const before = evaluate(doc, TJU_PROFILE)
    expect(before.findings.some((f) => f.ruleId === 'page.pageNumber')).toBe(true)
    expect(doc.sections[0].footerText).toContain('天津大学')

    const fixed = await applyFixes(buffer, doc, TJU_PROFILE, ['page.pageNumber'])
    const afterDoc = await parseDocx(fixed.buffer, 'existing-footer-fixed.docx')
    expect(afterDoc.sections[0].hasPageNumberField).toBe(true)
    expect(afterDoc.sections[0].footerText).toContain('天津大学')
  })

  it('adds page numbers to a section that lacks them, while keeping the existing one', async () => {
    const buffer = fixtureBuffer('multi-section.docx')
    const doc = await parseDocx(buffer, 'multi-section.docx')
    expect(doc.sections.length).toBeGreaterThanOrEqual(2)
    expect(doc.sections[0].hasPageNumberField).toBe(true)
    expect(doc.sections[1].hasPageNumberField).toBe(false)

    const fixed = await applyFixes(buffer, doc, TJU_PROFILE, ['page.pageNumber'])
    const afterDoc = await parseDocx(fixed.buffer, 'multi-section-fixed.docx')
    expect(afterDoc.sections.every((s) => s.hasPageNumberField)).toBe(true)
    expect(afterDoc.sections[1].footerText).toContain('第二节页脚')

    const again = await applyFixes(fixed.buffer, afterDoc, TJU_PROFILE, ['page.pageNumber'])
    expect(again.applied).not.toContain('page.pageNumber')
  })

  it('recognizes an existing TOC field instead of asking for one', async () => {
    const doc = await parseDocx(fixtureBuffer('toc-field.docx'), 'toc-field.docx')
    const report = evaluate(doc, TJU_PROFILE)
    expect(doc.hasTocField).toBe(true)
    expect(report.findings.some((f) => f.ruleId === 'structure.tocField')).toBe(false)
  })
})
