#!/usr/bin/env node
/**
 * smoke_plugin.mjs —— 插件层自检（不装 DSH 也能跑）。
 *
 * 校验：manifest 形状（dsh.bundle → cordis.patch.yml）、patch 与 package 的 name 一致、
 * frontmatter 可解析、list()/get() 返回的定义完整、正文能读到「17 种页型」这类关键约定。
 */
import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { list, get, name, inject, isSkillName } from '../lib/index.js'

const root = new URL('..', import.meta.url)
const pkgPath = fileURLToPath(new URL('package.json', root))
const patchPath = fileURLToPath(new URL('cordis.patch.yml', root))

const fails = []
const ok = (cond, msg) => { if (!cond) fails.push(msg) }

const pkg = JSON.parse(await readFile(pkgPath, 'utf8'))
ok(pkg.name === 'zgs-ppt', `package.json name 应为 zgs-ppt，实际 ${pkg.name}`)
ok(pkg.dsh?.bundle?.patch === './cordis.patch.yml', 'package.json 缺 dsh.bundle.patch')
ok(pkg.main === 'lib/index.js', 'package.json main 应指向 lib/index.js')

const patch = await readFile(patchPath, 'utf8').catch(() => '')
ok(patch.includes(`name: '${pkg.name}'`) || patch.includes(`name: "${pkg.name}"`), 'cordis.patch.yml 的 name 与 package.json 不一致')

ok(name === 'zgs-ppt', `插件导出 name 应为 zgs-ppt，实际 ${name}`)
ok(Array.isArray(inject) && inject.includes('skills'), '插件必须 inject skills')

const candidates = await list()
ok(candidates.length === 1, `list() 应返回 1 条候选，实际 ${candidates.length}`)
const c = candidates[0]
ok(isSkillName(c.name), `技能名不是 kebab-case：${c.name}`)
ok(c.name === pkg.name, '技能名与包名应一致')
ok(typeof c.description === 'string' && c.description.length > 40, '技能描述过短或缺失')
ok(c.resourceBase?.kind === 'directory', 'resourceBase 应为目录')

const def = await get(c)
ok(def.content.length > 2000, '技能正文过短，疑似只读到 frontmatter')
for (const needle of ['17 种页型', '12', 'awarded-cases', 'verify_pptx']) {
  ok(def.content.includes(needle), `技能正文缺少关键约定：${needle}`)
}

if (fails.length) {
  for (const f of fails) console.error('✗', f)
  process.exit(1)
}
console.log(`✓ 插件自检通过：${c.name} · 正文 ${def.content.length} 字 · 资源基准 ${c.resourceBase.path}`)
