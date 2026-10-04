import { setLang } from './i18n'

// 界面默认语言跟随浏览器，而 jsdom 报 en-US；单元测试固定为中文，与 e2e 的 locale: 'zh-CN' 对齐。
// 语言自动检测本身由 e2e「首次访问按浏览器语言选择德语」用例覆盖。
setLang('zh')
