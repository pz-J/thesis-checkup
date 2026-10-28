import { useMemo, useState } from 'react'
import { formatSize } from '../lib/docx'
import { GENERIC_PROFILE, TJU_PROFILE, parseProfile, type RuleProfile } from '../lib/profiles'
import { CloseIcon, DownloadIcon } from './Icons'

interface Props {
  profile: RuleProfile
  custom: boolean
  onClose: () => void
  onImport: (profile: RuleProfile) => void
  onDelete?: () => void
}

type LineMode = 'exact' | 'auto'

interface FormState {
  name: string
  shortName: string
  top: number
  right: number
  bottom: number
  left: number
  bodyFont: string
  bodySize: number
  lineMode: LineMode
  lineValue: number
  indent: number
  h1Font: string
  h1Size: number
  h2Font: string
  h2Size: number
  h3Font: string
  h3Size: number
  captionFont: string
  captionSize: number
  refTitleFont: string
  refTitleSize: number
  refFont: string
  refSize: number
  reqChinese: boolean
  reqEnglish: boolean
  reqKeywords: boolean
  reqToc: boolean
  reqRefs: boolean
  reqAck: boolean
  ackTerms: string
}

const FONTS = ['宋体', '仿宋', '楷体', '黑体', '微软雅黑', 'Times New Roman']
const SIZES = [9, 10.5, 12, 14, 15, 16, 18, 22, 24, 26, 36]

const sizeOptions = (extra: number[]) =>
  Array.from(new Set([...SIZES, ...extra.filter((n) => Number.isFinite(n))])).sort((a, b) => a - b)

const fontOptions = (extra: string[]) =>
  Array.from(new Set([...FONTS, ...extra.filter(Boolean)]))

function initialForm(profile: RuleProfile, custom: boolean): FormState {
  const lineValue =
    profile.body.line.rule === 'exact' ? profile.body.line.twips / 20 : profile.body.line.twips / 240
  return {
    name: custom ? profile.name : '我的学校格式',
    shortName: custom ? profile.shortName : '我的学校',
    top: profile.page.marginsCm.top,
    right: profile.page.marginsCm.right,
    bottom: profile.page.marginsCm.bottom,
    left: profile.page.marginsCm.left,
    bodyFont: profile.body.eastAsiaFont,
    bodySize: profile.body.sizePt,
    lineMode: profile.body.line.rule === 'auto' ? 'auto' : 'exact',
    lineValue: Math.round(lineValue * 10) / 10,
    indent: profile.body.firstLineChars,
    h1Font: profile.heading[1].font,
    h1Size: profile.heading[1].sizePt,
    h2Font: profile.heading[2].font,
    h2Size: profile.heading[2].sizePt,
    h3Font: profile.heading[3].font,
    h3Size: profile.heading[3].sizePt,
    captionFont: profile.caption.font,
    captionSize: profile.caption.sizePt,
    refTitleFont: profile.reference.titleFont,
    refTitleSize: profile.reference.titleSizePt,
    refFont: profile.reference.font,
    refSize: profile.reference.sizePt,
    reqChinese: profile.structure.requireChineseAbstract,
    reqEnglish: profile.structure.requireEnglishAbstract,
    reqKeywords: profile.structure.requireKeywords,
    reqToc: profile.structure.requireToc,
    reqRefs: profile.structure.requireReferences,
    reqAck: profile.structure.requireAcknowledgement,
    ackTerms: profile.structure.acknowledgementTerms.join('、'),
  }
}

function buildProfile(form: FormState, base: RuleProfile, custom: boolean): RuleProfile {
  const id = custom ? base.id : `custom-${Date.now().toString(36)}`
  return {
    ...base,
    id,
    name: form.name.trim(),
    shortName: form.shortName.trim() || '我的学校',
    source: '按表单设置',
    sourceUrl: undefined,
    note: '由「设置我的学校格式」表单创建。请以学院最新模板为准。',
    page: {
      ...base.page,
      marginsCm: { top: form.top, right: form.right, bottom: form.bottom, left: form.left },
    },
    body: {
      ...base.body,
      eastAsiaFont: form.bodyFont,
      sizePt: form.bodySize,
      line: {
        rule: form.lineMode,
        twips: Math.round(form.lineMode === 'exact' ? form.lineValue * 20 : form.lineValue * 240),
      },
      firstLineChars: form.indent,
    },
    heading: {
      1: { font: form.h1Font, sizePt: form.h1Size },
      2: { font: form.h2Font, sizePt: form.h2Size },
      3: { font: form.h3Font, sizePt: form.h3Size },
    },
    caption: { ...base.caption, font: form.captionFont, sizePt: form.captionSize },
    reference: {
      ...base.reference,
      titleFont: form.refTitleFont,
      titleSizePt: form.refTitleSize,
      font: form.refFont,
      sizePt: form.refSize,
    },
    structure: {
      ...base.structure,
      requireChineseAbstract: form.reqChinese,
      requireEnglishAbstract: form.reqEnglish,
      requireKeywords: form.reqKeywords,
      requireToc: form.reqToc,
      requireReferences: form.reqRefs,
      requireAcknowledgement: form.reqAck,
      acknowledgementTerms: form.ackTerms
        .split(/[、,，;；]/)
        .map((x) => x.trim())
        .filter(Boolean),
    },
  }
}

function validate(form: FormState): string | null {
  if (!form.name.trim()) return '先给这个规则包起个名字，比如「XX 学院本科论文」。'
  const margins = [form.top, form.right, form.bottom, form.left]
  if (margins.some((n) => !Number.isFinite(n) || n < 0.4 || n > 8)) return '页边距请在 0.4–8 cm 之间。'
  if (!Number.isFinite(form.bodySize) || form.bodySize < 5 || form.bodySize > 72) return '正文字号请在 5–72 pt 之间。'
  if (form.lineMode === 'exact' && (!Number.isFinite(form.lineValue) || form.lineValue < 5 || form.lineValue > 100)) {
    return '固定行距请在 5–100 磅之间。'
  }
  if (form.lineMode === 'auto' && (!Number.isFinite(form.lineValue) || form.lineValue < 0.8 || form.lineValue > 3)) {
    return '倍数行距请在 0.8–3 倍之间。'
  }
  if (form.indent < 0 || form.indent > 8) return '首行缩进请在 0–8 字符之间。'
  const sizes = [form.h1Size, form.h2Size, form.h3Size, form.captionSize, form.refTitleSize, form.refSize]
  if (sizes.some((n) => !Number.isFinite(n) || n < 5 || n > 72)) return '标题、题注和参考文献字号请在 5–72 pt 之间。'
  return null
}

export function RuleEditor({ profile, custom, onClose, onImport, onDelete }: Props) {
  const [mode, setMode] = useState<'form' | 'json'>('form')
  const [form, setForm] = useState<FormState>(() => initialForm(profile, custom))
  const [error, setError] = useState<string | null>(null)
  const [jsonText, setJsonText] = useState(() => JSON.stringify(profile, null, 2))

  const allFonts = useMemo(
    () =>
      fontOptions([
        profile.body.eastAsiaFont,
        profile.heading[1].font,
        profile.heading[2].font,
        profile.heading[3].font,
        profile.caption.font,
        profile.reference.titleFont,
        profile.reference.font,
      ]),
    [profile],
  )
  const allSizes = useMemo(
    () =>
      sizeOptions([
        profile.body.sizePt,
        profile.heading[1].sizePt,
        profile.heading[2].sizePt,
        profile.heading[3].sizePt,
        profile.caption.sizePt,
        profile.reference.titleSizePt,
        profile.reference.sizePt,
      ]),
    [profile],
  )

  const set = <K extends keyof FormState,>(key: K, value: FormState[K]) =>
    setForm((prev) => ({ ...prev, [key]: value }))

  const applyPreset = (preset: RuleProfile) => {
    setForm((prev) => ({ ...initialForm(preset, true), name: prev.name, shortName: prev.shortName }))
    setError(null)
  }

  const applyForm = () => {
    const validation = validate(form)
    if (validation) {
      setError(validation)
      return
    }
    const candidate = buildProfile(form, profile, custom)
    const { profile: normalized, error: parseError } = parseProfile(candidate)
    if (!normalized) {
      setError(parseError ?? '规则包格式不正确。')
      return
    }
    setError(null)
    onImport(normalized)
  }

  const applyJson = () => {
    let raw: unknown
    try {
      raw = JSON.parse(jsonText)
    } catch {
      setError('JSON 解析失败，请检查逗号、引号与括号。')
      return
    }
    const { profile: next, error: parseError } = parseProfile(raw)
    if (!next) {
      setError(parseError ?? '规则包格式不正确。')
      return
    }
    setError(null)
    onImport(next)
  }

  const currentJson = () => JSON.stringify(buildProfile(form, profile, custom), null, 2)

  const download = () => {
    const blob = new Blob([mode === 'json' ? jsonText : currentJson()], {
      type: 'application/json;charset=utf-8',
    })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `${(form.shortName || '我的学校').trim()}_规则包.json`
    document.body.appendChild(a)
    a.click()
    a.remove()
    setTimeout(() => URL.revokeObjectURL(url), 3000)
  }

  const switchMode = (next: 'form' | 'json') => {
    if (next === 'json') setJsonText(currentJson())
    setError(null)
    setMode(next)
  }

  const textField = (
    label: string,
    value: string,
    onChange: (v: string) => void,
    hint?: string,
    placeholder?: string,
  ) => (
    <label className="field">
      <span>{label}</span>
      <input value={value} placeholder={placeholder} onChange={(e) => onChange(e.target.value)} />
      {hint && <small>{hint}</small>}
    </label>
  )

  const numberField = (
    label: string,
    value: number,
    onChange: (v: number) => void,
    step = 0.1,
    hint?: string,
    suffix?: string,
  ) => (
    <label className="field">
      <span>{label}</span>
      <div className="number-input">
        <input
          type="number"
          step={step}
          value={Number.isFinite(value) ? value : ''}
          onChange={(e) => onChange(Number(e.target.value))}
        />
        {suffix && <em>{suffix}</em>}
      </div>
      {hint && <small>{hint}</small>}
    </label>
  )

  const fontSelect = (label: string, value: string, onChange: (v: string) => void) => (
    <label className="field">
      <span>{label}</span>
      <select value={value} onChange={(e) => onChange(e.target.value)}>
        {allFonts.map((font) => (
          <option key={font} value={font}>
            {font}
          </option>
        ))}
      </select>
    </label>
  )

  const sizeSelect = (label: string, value: number, onChange: (v: number) => void) => (
    <label className="field">
      <span>{label}</span>
      <select value={value} onChange={(e) => onChange(Number(e.target.value))}>
        {allSizes.map((size) => (
          <option key={size} value={size}>
            {formatSize(size)}
          </option>
        ))}
      </select>
    </label>
  )

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div
        className="modal editor-modal"
        role="dialog"
        aria-modal="true"
        aria-label="设置我的学校格式"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="modal-head">
          <div>
            <span className="modal-kicker">SET UP YOUR SCHOOL</span>
            <h2>设置我的学校格式</h2>
            <p>不用懂 JSON。按学院模板填几个数；不知道具体数值也没关系，先用默认值，体检后按报告提示再改。</p>
          </div>
          <button className="icon-button large" onClick={onClose} aria-label="关闭">
            <CloseIcon width={18} height={18} />
          </button>
        </div>

        <div className="editor-tabs" role="tablist">
          <button
            className={`editor-tab ${mode === 'form' ? 'active' : ''}`}
            onClick={() => switchMode('form')}
            role="tab"
            aria-selected={mode === 'form'}
          >
            按表单设置
          </button>
          <button
            className={`editor-tab ${mode === 'json' ? 'active' : ''}`}
            onClick={() => switchMode('json')}
            role="tab"
            aria-selected={mode === 'json'}
          >
            JSON 高级模式
          </button>
          <span className="editor-tab-note">普通用户用左边就够了</span>
        </div>

        {mode === 'form' ? (
          <div className="form-body">
            <div className="editor-help">
              <strong>不知道学校要求？</strong>
              <span>
                去教务处或学院官网找「毕业设计（论文）模板」，找不到就问导师或班长要一份 Word 模板。
                拿到模板后，照着里面「页面设置」和样式里的数字填即可。
              </span>
            </div>
            <div className="preset-row">
              <span>快速填入：</span>
              <button className="chip-button small" onClick={() => applyPreset(TJU_PROFILE)}>
                天大本科默认
              </button>
              <button className="chip-button small" onClick={() => applyPreset(GENERIC_PROFILE)}>
                通用中文论文
              </button>
              <button className="text-button small" onClick={() => setForm(initialForm(profile, custom))}>
                恢复当前规则
              </button>
            </div>

            <section className="form-section">
              <div className="form-section-head">
                <h3>1. 规则包名称</h3>
                <p>会出现在顶部下拉框里，方便你以后切换。</p>
              </div>
              <div className="field-grid two">
                {textField('名称', form.name, (v) => set('name', v), '例如：XX 学院本科毕业论文', 'XX 学院本科毕业论文')}
                {textField('简称', form.shortName, (v) => set('shortName', v), '最多 8 个字，显示在顶部', 'XX 学院')}
              </div>
            </section>

            <section className="form-section">
              <div className="form-section-head">
                <h3>2. 页面与正文</h3>
                <p>不知道具体数值？先用学校模板里的页面设置，或者保持默认。</p>
              </div>
              <div className="margin-grid">
                {numberField('上边距', form.top, (v) => set('top', v), 0.1, undefined, 'cm')}
                {numberField('下边距', form.bottom, (v) => set('bottom', v), 0.1, undefined, 'cm')}
                {numberField('左边距', form.left, (v) => set('left', v), 0.1, undefined, 'cm')}
                {numberField('右边距', form.right, (v) => set('right', v), 0.1, undefined, 'cm')}
              </div>
              <div className="inline-tools">
                <button className="text-button small" onClick={() => setForm((f) => ({ ...f, top: 2.5, right: 2.5, bottom: 2.5, left: 2.5 }))}>
                  四边都填 2.5 cm
                </button>
                <span>纸张固定按 A4 处理；Word 里可从「布局 → 页边距」查看。如学校用 B5 或 Letter，请用 JSON 高级模式。</span>
              </div>
              <div className="field-grid three">
                {fontSelect('正文中文字体', form.bodyFont, (v) => set('bodyFont', v))}
                {sizeSelect('正文字号', form.bodySize, (v) => set('bodySize', v))}
                <label className="field">
                  <span>首行缩进</span>
                  <select value={form.indent} onChange={(e) => set('indent', Number(e.target.value))}>
                    {[0, 1, 2, 3, 4].map((n) => (
                      <option key={n} value={n}>
                        {n} 字符
                      </option>
                    ))}
                  </select>
                </label>
              </div>
              <div className="field-grid three">
                <label className="field">
                  <span>行距类型</span>
                  <select
                    value={form.lineMode}
                    onChange={(e) => {
                      const next = e.target.value as LineMode
                      setForm((prev) => ({
                        ...prev,
                        lineMode: next,
                        lineValue: next === 'exact' ? 20 : 1.5,
                      }))
                    }}
                  >
                    <option value="exact">固定值（磅）</option>
                    <option value="auto">倍数行距</option>
                  </select>
                </label>
                {numberField(
                  form.lineMode === 'exact' ? '固定行距' : '行距倍数',
                  form.lineValue,
                  (v) => set('lineValue', v),
                  form.lineMode === 'exact' ? 0.5 : 0.1,
                  form.lineMode === 'exact' ? 'Word 里选「固定值」时填的数字' : '例如 1.5 倍',
                  form.lineMode === 'exact' ? '磅' : '倍',
                )}
                <div className="field field-static">
                  <span>对齐方式</span>
                  <div className="static-value">两端对齐（默认）</div>
                  <small>中文论文通常使用两端对齐</small>
                </div>
              </div>
            </section>

            <section className="form-section">
              <div className="form-section-head">
                <h3>3. 标题、题注与参考文献</h3>
                <p>只填字号和字体即可，位置与编号由体检规则处理。</p>
              </div>
              <div className="heading-grid">
                <div className="heading-row">
                  <span className="heading-row-label">一级标题</span>
                  {fontSelect('字体', form.h1Font, (v) => set('h1Font', v))}
                  {sizeSelect('字号', form.h1Size, (v) => set('h1Size', v))}
                </div>
                <div className="heading-row">
                  <span className="heading-row-label">二级标题</span>
                  {fontSelect('字体', form.h2Font, (v) => set('h2Font', v))}
                  {sizeSelect('字号', form.h2Size, (v) => set('h2Size', v))}
                </div>
                <div className="heading-row">
                  <span className="heading-row-label">三级标题</span>
                  {fontSelect('字体', form.h3Font, (v) => set('h3Font', v))}
                  {sizeSelect('字号', form.h3Size, (v) => set('h3Size', v))}
                </div>
              </div>
              <div className="field-grid four">
                {fontSelect('图表题注字体', form.captionFont, (v) => set('captionFont', v))}
                {sizeSelect('图表题注字号', form.captionSize, (v) => set('captionSize', v))}
                {fontSelect('参考文献标题字体', form.refTitleFont, (v) => set('refTitleFont', v))}
                {sizeSelect('参考文献标题字号', form.refTitleSize, (v) => set('refTitleSize', v))}
              </div>
              <div className="field-grid two narrow">
                {fontSelect('参考文献条目字体', form.refFont, (v) => set('refFont', v))}
                {sizeSelect('参考文献条目字号', form.refSize, (v) => set('refSize', v))}
              </div>
            </section>

            <section className="form-section">
              <div className="form-section-head">
                <h3>4. 必须包含的部分</h3>
                <p>不确定就全勾选；如果学院不要求英文摘要，可以取消。</p>
              </div>
              <div className="check-grid">
                <label className="check-field">
                  <input type="checkbox" checked={form.reqChinese} onChange={(e) => set('reqChinese', e.target.checked)} />
                  中文摘要
                </label>
                <label className="check-field">
                  <input type="checkbox" checked={form.reqEnglish} onChange={(e) => set('reqEnglish', e.target.checked)} />
                  英文摘要
                </label>
                <label className="check-field">
                  <input type="checkbox" checked={form.reqKeywords} onChange={(e) => set('reqKeywords', e.target.checked)} />
                  关键词
                </label>
                <label className="check-field">
                  <input type="checkbox" checked={form.reqToc} onChange={(e) => set('reqToc', e.target.checked)} />
                  目录
                </label>
                <label className="check-field">
                  <input type="checkbox" checked={form.reqRefs} onChange={(e) => set('reqRefs', e.target.checked)} />
                  参考文献
                </label>
                <label className="check-field">
                  <input type="checkbox" checked={form.reqAck} onChange={(e) => set('reqAck', e.target.checked)} />
                  致谢 / 谢辞
                </label>
              </div>
              {textField(
                '致谢标题的写法',
                form.ackTerms,
                (v) => set('ackTerms', v),
                '用顿号或逗号分隔，例如：致谢、谢辞',
                '致谢、谢辞',
              )}
            </section>

            {error && <div className="editor-error">{error}</div>}

            <div className="editor-note">
              表单只覆盖最常用的规则；纸张固定 A4，正文两端对齐。需要 B5、Letter、特殊摘要标题字号等细节时，切到 JSON 高级模式。
              规则包保存在本机，不会上传。
            </div>
          </div>
        ) : (
          <div className="form-body">
            <div className="editor-toolbar">
              <span>高级模式：直接编辑完整规则包 JSON</span>
              <div className="editor-toolbar-actions">
                <button className="text-button small" onClick={() => setJsonText(currentJson())}>
                  用表单值生成 JSON
                </button>
              </div>
            </div>
            <textarea
              className="rule-editor"
              value={jsonText}
              onChange={(e) => {
                setJsonText(e.target.value)
                setError(null)
              }}
              spellCheck={false}
              aria-label="规则包 JSON"
            />
            {error && <div className="editor-error">{error}</div>}
            <div className="editor-note">
              数值越界会回退到默认值；字体名直接写入文档。不同学院、不同年份的细则可能不同，请以最新模板为准。
            </div>
          </div>
        )}

        <div className="modal-actions spread">
          <div>
            {custom && onDelete && (
              <button className="chip-button danger" onClick={onDelete}>
                删除这个规则包
              </button>
            )}
          </div>
          <div className="modal-actions">
            <button className="chip-button" onClick={onClose}>
              取消
            </button>
            {mode === 'form' ? (
              <button className="primary-button" onClick={applyForm}>
                保存并应用
              </button>
            ) : (
              <>
                <button className="text-button small" onClick={download}>
                  下载 JSON
                </button>
                <button className="primary-button" onClick={applyJson}>
                  导入并应用
                </button>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}
