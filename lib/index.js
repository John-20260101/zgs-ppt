/**
 * zgs-ppt — DSH 技能提供者：把本包自带的 SKILL.md 挂到 `ctx.skills`。
 *
 * 为什么要有这一层：插件市场（以及 `dsh plugin add`）只认「package.json 里声明了
 * `dsh.bundle`、根目录带 `cordis.patch.yml`」的包；而技能正文只是一份 SKILL.md，
 * 它要进入会话的技能目录，就需要一个 provider。本文件就是那层很薄的接线。
 *
 * 三条刻意的性质：
 *   · **不复制正文** —— 候选与正文都在调用时从 SKILL.md 现读，改正文不必动本文件，
 *     也不会出现「描述一份、正文一份、各自漂移」；
 *   · **零依赖** —— 只用 `node:` 内置模块；
 *   · **资源基准指向包根** —— 正文里写的 `references/page-skeletons.md`、
 *     `scripts/build_pptx.py` 因此能按相对路径解析。
 *
 * @module zgs-ppt
 */

import { watch } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'

/** 包根（lib/ 的上一级）。 */
const ROOT_URL = new URL('..', import.meta.url)
const SKILL_URL = new URL('SKILL.md', ROOT_URL)
const SKILL_PATH = fileURLToPath(SKILL_URL)

/** provider 名，在 `ctx.skills` 注册表里唯一。 */
const PROVIDER_NAME = 'zgs-ppt'

/** 编辑器保存常常连发多个事件，合并成一次目录失效。 */
const WATCH_DEBOUNCE_MS = 150

/**
 * bundled 档的秩（DSH 里 BUNDLED_SKILL_RANK = 600，低秩优先）。
 * 写常量而非 import，是为了保持零依赖：装了插件的机器若在
 * `~/.agents/skills/zgs-ppt` 另放了一份本地版（user-agents = 500），本地那份仍然优先。
 */
const RANK = 600

/**
 * 读 SKILL.md 的 YAML frontmatter。只认本技能真正用到的四个键，
 * 因此不需要 YAML 依赖：name / description / disable-model-invocation / user-invocable。
 *
 * @param {string} text - SKILL.md 全文。
 * @returns {{fm: Record<string, string>, body: string}}
 */
export function parseFrontmatter(text) {
  // 去 BOM、CRLF 归一：编码与 Windows 检出不该改变解析结果。
  const src = text.replace(/^\uFEFF/, '').replace(/\r\n/g, '\n')
  if (!src.startsWith('---')) return { fm: {}, body: text }
  const end = src.indexOf('\n---', 3)
  if (end === -1) return { fm: {}, body: text }
  const fm = {}
  for (const line of src.slice(3, end).split('\n')) {
    const m = line.match(/^([A-Za-z0-9_-]+):\s*(.*)$/)
    if (m) fm[m[1]] = m[2].trim().replace(/^["']|["']$/g, '')
  }
  return { fm, body: src.slice(end + 4).trimStart() }
}

/** 技能名必须是 kebab-case，否则注册表会拒收这个候选。 */
export function isSkillName(name) {
  return /^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(String(name))
}

/** frontmatter 的调用开关：省略即两处都开。 */
function invocationOf(fm) {
  return {
    modelInvocable: String(fm['disable-model-invocation'] ?? '').toLowerCase() !== 'true',
    userInvocable: String(fm['user-invocable'] ?? '').toLowerCase() !== 'false',
  }
}

/**
 * 列出候选：每次调用都重读 SKILL.md，所以改完立刻生效，不需要版本号或缓存失效。
 *
 * 读失败不吞：让 `list()` 拒绝比返回空数组诚实——空数组等于告诉模型「这个技能被删了」。
 *
 * @returns {Promise<object[]>}
 */
export async function list() {
  const { fm } = parseFrontmatter(await readFile(SKILL_URL, 'utf8'))
  if (!isSkillName(fm.name)) {
    throw new Error(`SKILL.md 的 frontmatter name 不是 kebab-case：${JSON.stringify(fm.name)}`)
  }
  if (!fm.description) throw new Error('SKILL.md 的 frontmatter 缺少 description')
  return [
    {
      name: fm.name,
      description: fm.description,
      invocation: invocationOf(fm),
      provider: PROVIDER_NAME,
      source: 'bundled',
      // 资源基准指向包根：正文里的 references/、scripts/ 相对路径因此可解析。
      resourceBase: { kind: 'directory', path: fileURLToPath(ROOT_URL) },
      rank: RANK,
      path: SKILL_PATH,
      locator: SKILL_URL,
    },
  ]
}

/**
 * 读正文：frontmatter 已剥离，返回的就是指令正文。
 *
 * @param {object} candidate - `list()` 给出的候选。
 * @returns {Promise<object>}
 */
export async function get(candidate) {
  const { fm, body } = parseFrontmatter(await readFile(SKILL_URL, 'utf8'))
  // 两次读之间文件被换过 —— 抛出，让注册表重新发现，而不是交出一份名不副实的定义。
  if (fm.name !== candidate.name) {
    throw new Error(
      `SKILL.md 的 name 已变为 ${JSON.stringify(fm.name)}，与候选 ${JSON.stringify(candidate.name)} 不符`,
    )
  }
  return {
    name: candidate.name,
    description: fm.description || candidate.description,
    invocation: candidate.invocation,
    provider: PROVIDER_NAME,
    source: 'bundled',
    resourceBase: candidate.resourceBase,
    path: SKILL_PATH,
    content: body,
  }
}

/** provider 实例（无内部状态：不缓存正文）。 */
export const provider = { name: PROVIDER_NAME, list, get }

/** Cordis 插件名（与 cordis.patch.yml 的 id 一致）。 */
export const name = PROVIDER_NAME
/** 依赖的服务：技能注册表。 */
export const inject = ['skills']

/**
 * 监听 SKILL.md，改动后让注册表重读目录。
 *
 * 监听目录而不是文件：编辑器「写临时文件再改名」的保存方式会让文件级 watcher 失效。
 * 只认 SKILL.md 一个文件名（同目录还有 references/、scripts/，它们改了与技能目录无关）。
 */
function watchSkillFile(control, ctx) {
  // 老宿主没有 invalidate：不装监听，退化成「改完重启生效」。
  if (typeof control?.invalidate !== 'function') return () => {}
  let timer
  let closed = false
  let watcher
  try {
    watcher = watch(fileURLToPath(ROOT_URL), { persistent: false }, (_event, filename) => {
      if (closed || (filename !== null && filename !== 'SKILL.md')) return
      if (timer !== undefined) clearTimeout(timer)
      timer = setTimeout(() => {
        timer = undefined
        if (!closed) control.invalidate()
      }, WATCH_DEBOUNCE_MS)
    })
  } catch (error) {
    ctx?.logger?.warn?.(`zgs-ppt: SKILL.md 监听未启动，改完需重启才刷新技能目录：${error.message}`)
    return () => {}
  }
  const stop = () => {
    if (closed) return
    closed = true
    if (timer !== undefined) clearTimeout(timer)
    watcher.close()
  }
  watcher.on('error', (error) => {
    if (closed) return
    ctx?.logger?.warn?.(`zgs-ppt: SKILL.md 监听已停用：${error.message}`)
    stop()
  })
  control.signal?.addEventListener?.('abort', stop, { once: true })
  return stop
}

/**
 * 挂载：把本 provider 注册到 `ctx.skills`。
 *
 * @param {{skills: {registerProvider: Function}, effect?: Function, logger?: object}} ctx
 */
export function apply(ctx) {
  ctx.skills.registerProvider((control) => {
    const stop = watchSkillFile(control, ctx)
    if (typeof ctx.effect === 'function') {
      ctx.effect(function* () {
        yield () => stop()
      }, 'zgs-ppt skill watcher')
    }
    return provider
  })
}
