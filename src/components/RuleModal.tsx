import type { RuleProfile } from '../lib/profiles'
import { CloseIcon } from './Icons'

export function RuleModal({ profile, onClose, onEdit }: { profile: RuleProfile; onClose: () => void; onEdit: () => void }) {
  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal rule-modal" role="dialog" aria-modal="true" aria-label="规则包说明" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <div>
            <span className="modal-kicker">RULE PACK</span>
            <h2>{profile.name}</h2>
            <p>{profile.note}</p>
          </div>
          <button className="icon-button large" onClick={onClose} aria-label="关闭">
            <CloseIcon width={18} height={18} />
          </button>
        </div>

        <table className="rule-table">
          <tbody>
            <tr>
              <th>页面</th>
              <td>
                A4 · 页边距 上 {profile.page.marginsCm.top} / 下 {profile.page.marginsCm.bottom} / 左{' '}
                {profile.page.marginsCm.left} / 右 {profile.page.marginsCm.right} cm
              </td>
            </tr>
            <tr>
              <th>正文</th>
              <td>
                {profile.body.eastAsiaFont} {profile.body.sizePt}pt ·{' '}
                {profile.body.line.rule === 'exact'
                  ? `固定 ${profile.body.line.twips / 20} 磅行距`
                  : `${profile.body.line.twips / 240} 倍行距`}{' '}
                · 首行缩进 {profile.body.firstLineChars} 字符 · 两端对齐
              </td>
            </tr>
            <tr>
              <th>一级标题</th>
              <td>
                {profile.heading[1].font} {profile.heading[1].sizePt}pt
              </td>
            </tr>
            <tr>
              <th>二级标题</th>
              <td>
                {profile.heading[2].font} {profile.heading[2].sizePt}pt
              </td>
            </tr>
            <tr>
              <th>三级标题</th>
              <td>
                {profile.heading[3].font} {profile.heading[3].sizePt}pt
              </td>
            </tr>
            <tr>
              <th>图表题注</th>
              <td>
                {profile.caption.font} {profile.caption.sizePt}pt · 居中；图题在图下方，表题在表上方
              </td>
            </tr>
            <tr>
              <th>参考文献</th>
              <td>
                标题 {profile.reference.titleFont} {profile.reference.titleSizePt}pt 顶格；条目{' '}
                {profile.reference.font} {profile.reference.sizePt}pt（外文 Times New Roman）
              </td>
            </tr>
          </tbody>
        </table>

        <div className="rule-source">
          <span>规则来源</span>
          <p>{profile.source}</p>
          {profile.sourceUrl && (
            <a href={profile.sourceUrl} target="_blank" rel="noreferrer">
              查看试行稿存档（第三方，仅供参考）
            </a>
          )}
        </div>

        <div className="rule-disclaimer">
          规则包是公开规范的工程化表达，不是学院官方审查结论。不同年份、学院或学位类型可能另有细则，请以教务 / 研究生院最新模板为准。
        </div>

        <div className="modal-actions spread">
          <button className="chip-button" onClick={onClose}>
            关闭
          </button>
          <button className="primary-button" onClick={onEdit}>
            设置我的学校格式
          </button>
        </div>
      </div>
    </div>
  )
}
