import type { Severity } from './docx'
import type { Finding, Report } from './rules'

export interface ParaIssue {
  severity: Severity
  titles: string[]
  ruleIds: string[]
}

const RANK: Record<Severity, number> = { error: 0, warn: 1, info: 2 }

export const SEVERITY_LABEL: Record<Severity, string> = {
  error: '必改',
  warn: '建议',
  info: '提示',
}

export function worst(a: Severity, b: Severity): Severity {
  return RANK[a] <= RANK[b] ? a : b
}

export function buildParaIssueMap(report: Report): Map<number, ParaIssue> {
  const map = new Map<number, ParaIssue>()
  for (const f of report.findings) {
    for (const o of f.occurrences) {
      if (o.paragraphIndex === undefined) continue
      const cur = map.get(o.paragraphIndex)
      if (!cur) {
        map.set(o.paragraphIndex, { severity: f.severity, titles: [f.title], ruleIds: [f.ruleId] })
      } else {
        cur.severity = worst(cur.severity, f.severity)
        if (!cur.titles.includes(f.title)) cur.titles.push(f.title)
        if (!cur.ruleIds.includes(f.ruleId)) cur.ruleIds.push(f.ruleId)
      }
    }
  }
  return map
}

export function findingsForParagraph(report: Report, index: number): Finding[] {
  return report.findings.filter((f) => f.occurrences.some((o) => o.paragraphIndex === index))
}

export function uniqueActuals(f: Finding): string[] {
  const set = new Set<string>()
  for (const o of f.occurrences) if (o.actual) set.add(o.actual)
  return Array.from(set)
}
