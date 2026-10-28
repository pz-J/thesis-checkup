import type { RuleProfile } from '../lib/profiles'
import { ShieldIcon } from './Icons'

interface Props {
  profiles: RuleProfile[]
  profileId: string
  onProfileChange: (id: string) => void
  onAddProfile: () => void
  onRules: () => void
}

export function Header({ profiles, profileId, onProfileChange, onAddProfile, onRules }: Props) {
  return (
    <header className="site-header">
      <button className="brand" onClick={() => window.scrollTo({ top: 0, behavior: 'smooth' })}>
        <svg className="brand-mark" viewBox="0 0 24 24" aria-hidden="true">
          <rect x="3.2" y="3.2" width="17.6" height="17.6" rx="2.4" fill="none" stroke="currentColor" strokeWidth="1.5" />
          <path d="M7.6 7v10" stroke="var(--accent)" strokeWidth="1.8" />
          <path d="M11 7.2h5.4M11 11.8h4.2M11 16.4h3" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
        </svg>
        <span className="brand-text">
          论文体检
          <small>THESIS CHECKUP</small>
        </span>
      </button>

      <div className="header-right">
        <label className="profile-select">
          <span>规则包</span>
          <select
            value={profileId}
            onChange={(e) => {
              if (e.target.value === '__new__') {
                onAddProfile()
                return
              }
              onProfileChange(e.target.value)
            }}
            aria-label="选择学校规则包"
          >
            {profiles.map((p) => (
              <option key={p.id} value={p.id}>
                {p.shortName}
              </option>
            ))}
            <option value="__new__">＋ 设置我的学校</option>
          </select>
        </label>
        <button className="chip-button subtle" onClick={onRules}>
          规则说明
        </button>
        <span className="privacy-pill" title="所有解析与修复都在浏览器内完成">
          <ShieldIcon width={14} height={14} />
          本地运行 · 文件不上传
        </span>
      </div>
    </header>
  )
}
