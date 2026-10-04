/** 离线自测:node --test(依赖已构建的 lib/) */
import assert from 'node:assert/strict'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { test } from 'node:test'

import { resolveConfig } from '../lib/config.js'
import { parseCsv, toCsv } from '../lib/lib/csv.js'
import { editWorkbook, readWorkbook, writeWorkbook } from '../lib/lib/xlsx.js'
import { parseMarkdown, readDocx, writeDocx } from '../lib/lib/docx.js'
import { buildOfficeTools } from '../lib/tools.js'

function tmpDir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'office-suite-'))
}

test('csv roundtrip', () => {
  const rows = [['姓名', '备注'], ['张三', '含,逗号'], ['李四', '含"引号"']]
  const text = toCsv(rows)
  assert.deepEqual(parseCsv(text), rows)
})

test('markdown parser', () => {
  const blocks = parseMarkdown('# 标题\n\n- 甲\n- 乙\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n\n正文')
  assert.equal(blocks[0].kind, 'h1')
  assert.equal(blocks[1].kind, 'bullet')
  assert.equal(blocks[2].kind, 'bullet')
  assert.equal(blocks[3].kind, 'table')
  assert.deepEqual(blocks[3].rows, [['A', 'B'], ['1', '2']])
  assert.equal(blocks[4].kind, 'p')
})

test('xlsx write/read/edit', async () => {
  const dir = tmpDir()
  const file = path.join(dir, 't.xlsx')
  await writeWorkbook(file, [
    { name: '明细', rows: [['产品', '数量', '金额'], ['键盘', 10, '=B2*199'], ['鼠标', 20, '=B3*89']], widths: [14, 8, 10] },
  ])
  const read = await readWorkbook(file, undefined, 0, 100)
  assert.deepEqual(read.sheets, ['明细'])
  assert.match(read.text, /键盘/)
  assert.match(read.text, /鼠标/)

  const edited = await editWorkbook(file, [
    { op: 'set_value', sheet: '明细', cell: 'E1', value: '合计' },
    { op: 'set_style', sheet: '明细', cell: 'E1', style: { bold: true, fill: 'DDEBF7' } },
    { op: 'set_style', sheet: '明细', range: 'A1:C1', style: { bold: true } },
    { op: 'add_sheet', name: '汇总', rows: [['项目', '值']] },
  ])
  assert.equal(edited.applied, 4)
  const read2 = await readWorkbook(file, '汇总', 0, 10)
  assert.match(read2.text, /项目/)
  fs.rmSync(dir, { recursive: true, force: true })
})

test('docx write/read', async () => {
  const dir = tmpDir()
  const file = path.join(dir, 't.docx')
  const md = '# 测试报告\n\n这是正文。\n\n- 要点一\n- 要点二\n\n| 项目 | 值 |\n| --- | --- |\n| A | 1 |'
  const result = await writeDocx(file, md, '测试')
  assert.equal(result.paragraphs, 4)
  assert.equal(result.tables, 1)
  const text = await readDocx(file)
  assert.match(text, /测试报告/)
  assert.match(text, /要点一/)
  assert.match(text, /正文/)
  fs.rmSync(dir, { recursive: true, force: true })
})

test('tools: registration + end-to-end via tools.ts', async () => {
  const registered = []
  const ctx = { tools: { register: (def) => registered.push(def) } }
  const config = resolveConfig({})
  for (const def of buildOfficeTools(config)) ctx.tools.register(def)
  assert.deepEqual(registered.map((r) => r.name), [
    'office_read', 'office_write_docx', 'office_write_xlsx', 'office_edit_xlsx', 'office_csv_write', 'office_health',
  ])

  const dir = tmpDir()
  const exec = { agent: { session: { header: { cwd: dir } } } }
  const byName = new Map(buildOfficeTools(config).map((d) => [d.name, d]))

  const written = await byName.get('office_write_xlsx').execute(
    { path: 'out.xlsx', sheets: [{ name: 'S', rows: [['k', 'v'], ['a', 1]] }] },
    exec,
  )
  assert.equal(written.written[0].name, 'S')

  const read = await byName.get('office_read').execute({ path: 'out.xlsx' }, exec)
  assert.match(read.text, /a/)
  // render 是模型可见通道:必须带上正文,否则调用方看不到内容(防回归)
  const readRendered = byName.get('office_read').output.render({}, read).map((b) => b.text).join('\n')
  assert.match(readRendered, /已读取/)
  assert.match(readRendered, /a/)

  await byName.get('office_csv_write').execute({ path: 'out.csv', rows: [['a', 'b'], ['1,2', '3']] }, exec)
  const csvRead = await byName.get('office_read').execute({ path: 'out.csv' }, exec)
  assert.match(csvRead.text, /1,2/)

  await byName.get('office_write_docx').execute({ path: 'out.docx', markdown: '# T\n\nbody' }, exec)
  const docRead = await byName.get('office_read').execute({ path: 'out.docx' }, exec)
  assert.match(docRead.text, /body/)
  const docRendered = byName.get('office_read').output.render({}, docRead).map((b) => b.text).join('\n')
  assert.match(docRendered, /body/)

  // 错误路径:不存在的文件、空参数、无法识别的格式、空 ops
  await assert.rejects(() => byName.get('office_read').execute({ path: 'nope.docx' }, exec), /文件不存在/)
  await assert.rejects(() => byName.get('office_read').execute({}, exec), /path不能为空/)
  await byName.get('office_write_docx').execute({ path: 'note.md', markdown: 'plain' }, exec)
  await assert.rejects(() => byName.get('office_read').execute({ path: 'note.md' }, exec), /无法识别的格式/)
  await assert.rejects(() => byName.get('office_edit_xlsx').execute({ path: 'out.xlsx', ops: [] }, exec), /ops必须是非空数组/)

  const health = await byName.get('office_health').execute({}, exec)
  assert.equal(health.ok, true)
  fs.rmSync(dir, { recursive: true, force: true })
})
