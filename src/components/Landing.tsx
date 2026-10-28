import { useRef, useState } from 'react'
import type { RuleProfile } from '../lib/profiles'
import { AlertIcon, ArrowRightIcon, FileIcon, ScanIcon, ShieldIcon, UploadIcon } from './Icons'

interface Props {
  profile: RuleProfile
  ruleCount: number
  busy: string | null
  error: string | null
  onFile: (file: File) => void
  onSample: (kind: 'issues' | 'clean') => void
  onEditRules: () => void
}

export function Landing({ profile, ruleCount, busy, error, onFile, onSample, onEditRules }: Props) {
  const inputRef = useRef<HTMLInputElement>(null)
  const [dragging, setDragging] = useState(false)

  const pick = () => inputRef.current?.click()

  return (
    <main className="landing">
      <section className="hero-grid">
        <div className="hero-copy">
          <div className="eyebrow">BEFORE YOU SUBMIT · 提交前自查</div>
          <h1>
            交稿之前，
            <br />
            先<em>体检</em>一遍。
          </h1>
          <p className="hero-sub">
            把 Word 论文拖进来，几十秒看清格式问题。页面、正文、标题、图表题注、参考文献，逐项比对学校规则，并给出可以定位到段落的修改建议。
          </p>
          <div className="hero-stats">
            <div>
              <strong>{ruleCount}</strong>
              <span>项格式规则</span>
            </div>
            <div>
              <strong>0</strong>
              <span>字节上传</span>
            </div>
            <div>
              <strong>.docx</strong>
              <span>直接在浏览器解析</span>
            </div>
          </div>
        </div>

        <div className="drop-wrap">
          <div
            className={`drop-zone ${dragging ? 'is-dragging' : ''} ${busy ? 'is-busy' : ''}`}
            onDragOver={(e) => {
              e.preventDefault()
              setDragging(true)
            }}
            onDragLeave={() => setDragging(false)}
            onDrop={(e) => {
              e.preventDefault()
              setDragging(false)
              const f = e.dataTransfer.files?.[0]
              if (f) onFile(f)
            }}
          >
            <div className="drop-corner tl" />
            <div className="drop-corner tr" />
            <div className="drop-corner bl" />
            <div className="drop-corner br" />
            {busy ? (
              <div className="scan-state">
                <div className="scan-line" />
                <ScanIcon width={26} height={26} />
                <p>{busy}</p>
                <span>论文没有离开这台电脑</span>
              </div>
            ) : (
              <>
                <div className="drop-icon">
                  <FileIcon width={26} height={26} />
                </div>
                <h2>拖入你的论文</h2>
                <p>或选择一个 .docx 文件</p>
                <button className="primary-button" onClick={pick}>
                  <UploadIcon width={16} height={16} />
                  选择文件
                </button>
                <button className="text-button" onClick={() => onSample('issues')}>
                  没有文件？加载示例论文
                  <ArrowRightIcon width={14} height={14} />
                </button>
                <button className="text-button muted" onClick={() => onSample('clean')}>
                  看一份满分样例
                </button>
              </>
            )}
            <input
              ref={inputRef}
              type="file"
              accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              hidden
              onChange={(e) => {
                const f = e.target.files?.[0]
                if (f) onFile(f)
                e.currentTarget.value = ''
              }}
            />
          </div>

          <div className="profile-card">
            <div className="profile-card-head">
              <span>当前规则</span>
              <strong>{profile.shortName}</strong>
            </div>
            <p>
              A4 · {profile.page.marginsCm.top} cm 页边距 · {profile.body.eastAsiaFont}
              {profile.body.sizePt}pt ·{' '}
              {profile.body.line.rule === 'exact'
                ? `固定 ${profile.body.line.twips / 20} 磅行距`
                : `${profile.body.line.twips / 240} 倍行距`}
            </p>
            <div className="profile-card-foot">
              <span className="profile-note">以学院最新模板为准</span>
              <button className="text-button small" onClick={onEditRules}>
                不是这个规则？设置我的学校
              </button>
            </div>
          </div>

          <div className="drop-foot">
            <ShieldIcon width={14} height={14} />
            支持 .docx；旧版 .doc 请先在 Word 中另存为 .docx
          </div>
        </div>
      </section>

      {error && (
        <div className="error-banner" role="alert">
          <AlertIcon width={16} height={16} />
          <div>
            <strong>没能读取这份文件</strong>
            <span>{error}</span>
          </div>
        </div>
      )}

      <section className="landing-notes">
        <div className="note-block">
          <div className="note-kicker">它检查什么</div>
          <ul className="check-list">
            <li>纸张、页边距、页码字段</li>
            <li>正文字体、字号、行距、缩进</li>
            <li>标题层级、样式与章前分页</li>
            <li>图表题注与参考文献著录</li>
          </ul>
        </div>
        <div className="note-block">
          <div className="note-kicker">它不做什么</div>
          <ul className="plain-list">
            <li>不查重，不判断内容原创性</li>
            <li>不检测所谓「AI 率」</li>
            <li>不上传论文，不依赖账号</li>
            <li>不替代学院正式审查</li>
          </ul>
        </div>
        <div className="note-block">
          <div className="note-kicker">常见的退修</div>
          <ul className="plain-list">
            <li>页边距差了 5 毫米</li>
            <li>正文不知何时变成了五号</li>
            <li>图题没有居中、放在图上方</li>
            <li>参考文献缺 [J] 或编号断档</li>
          </ul>
        </div>
      </section>
    </main>
  )
}
