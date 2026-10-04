/** 8 个宿主工具定义:office_read / write_docx / write_xlsx / edit_xlsx / csv_write / docx_info / fix_docx / health */
import fs from 'node:fs';
import path from 'node:path';
import { ensureParent, fileExists, resolvePath, truncate } from './fs.js';
import { detectDelimiter, parseCsv, toCsv } from './lib/csv.js';
import { readDocx, writeDocx } from './lib/docx.js';
import { fixDocx, inspectDocx } from './lib/docx-xml.js';
import { editWorkbook, readWorkbook, writeWorkbook } from './lib/xlsx.js';
const PLUGIN_VERSION = '0.2.0';
const OBJ_SCHEMA = { type: 'object', additionalProperties: true };
function block(text) {
    return [{ type: 'text', text }];
}
/** office_read 的模型可见渲染:必须带上正文,否则调用方看不到文档内容 */
function renderRead(v) {
    const r = (v ?? {});
    const head = `已读取 ${r.path}(${r.format}${r.sheets?.length ? `, sheets: ${r.sheets.join('/')}` : ''}${typeof r.totalRows === 'number' ? `, 共 ${r.totalRows} 行` : ''})`;
    const warn = r.truncated ? '\n⚠️ 超长已截断,可用 offset/limit 翻页或缩小范围' : '';
    return `${head}${warn}\n\n${r.text ?? ''}`;
}
function str(args, key) {
    return typeof args[key] === 'string' ? args[key].trim() : '';
}
function numOr(args, key, fallback) {
    const v = args[key];
    return typeof v === 'number' && Number.isFinite(v) ? v : fallback;
}
function requireString(args, key, label) {
    const v = str(args, key);
    if (!v)
        throw new Error(`${label}不能为空`);
    return v;
}
function requireArray(args, key, label) {
    const v = args[key];
    if (!Array.isArray(v) || v.length === 0)
        throw new Error(`${label}必须是非空数组`);
    return v;
}
export function buildOfficeTools(config) {
    return [
        {
            name: 'office_read',
            description: '读取 Office 办公文档内容并转为纯文本:Word(.docx)、Excel(.xlsx)、CSV/TSV/TXT。' +
                'Excel 默认读第一个工作表,可用 sheet 指定、offset/limit 翻页;内置 read 工具读不了二进制,读取办公文档一律用本工具。',
            parameters: {
                type: 'object',
                properties: {
                    path: { type: 'string', description: '文件路径(相对会话工作目录或绝对路径)' },
                    format: { type: 'string', enum: ['auto', 'docx', 'xlsx', 'csv'], description: '默认 auto 按扩展名识别' },
                    sheet: { type: 'string', description: 'xlsx 工作表名,缺省第一个' },
                    offset: { type: 'number', description: 'xlsx/csv 跳过的行数(含表头行,从 0 开始),默认 0' },
                    limit: { type: 'number', description: 'xlsx/csv 最多返回行数,默认按插件配置 maxSheetRows' },
                },
                required: ['path'],
                additionalProperties: false,
            },
            output: { schema: OBJ_SCHEMA, render: (_a, v) => block(renderRead(v)) },
            async execute(rawArgs, exec) {
                const args = (rawArgs ?? {});
                const file = resolvePath(requireString(args, 'path', 'path'), exec);
                if (!fileExists(file))
                    throw new Error(`文件不存在: ${file}`);
                const ext = path.extname(file).toLowerCase();
                if (ext === '.xls') {
                    throw new Error('旧版 .xls 二进制格式暂不支持,请先用 Excel/WPS 另存为 .xlsx(本插件读 .xlsx)');
                }
                const format = str(args, 'format') || 'auto';
                const kind = format === 'auto'
                    ? ext === '.docx'
                        ? 'docx'
                        : ext === '.xlsx'
                            ? 'xlsx'
                            : ext === '.csv' || ext === '.tsv' || ext === '.txt'
                                ? 'csv'
                                : ''
                    : format;
                const offset = Math.max(0, Math.floor(numOr(args, 'offset', 0)));
                const limit = Math.max(1, Math.floor(numOr(args, 'limit', config.maxSheetRows)));
                if (kind === 'docx') {
                    const raw = await readDocx(file);
                    const out = truncate(raw, config.maxReadChars);
                    return { path: file, format: 'docx', text: out.text, truncated: out.truncated };
                }
                if (kind === 'xlsx') {
                    const sheet = str(args, 'sheet') || undefined;
                    const result = await readWorkbook(file, sheet, offset, limit);
                    const out = truncate(result.text, config.maxReadChars);
                    return {
                        path: file,
                        format: 'xlsx',
                        sheets: result.sheets,
                        totalRows: result.totalRows,
                        text: out.text,
                        truncated: out.truncated,
                    };
                }
                if (kind === 'csv') {
                    const raw = fs.readFileSync(file, 'utf8');
                    const delimiter = ext === '.tsv' ? '\t' : detectDelimiter(raw);
                    const all = parseCsv(raw, delimiter);
                    const page = all.slice(offset, offset + limit);
                    const text = page.map(r => r.join('\t')).join('\n');
                    const out = truncate(text, config.maxReadChars);
                    return {
                        path: file,
                        format: 'csv',
                        totalRows: all.length,
                        text: out.text,
                        truncated: out.truncated || all.length > offset + limit,
                    };
                }
                throw new Error(`无法识别的格式: ${file}(支持 .docx/.xlsx/.csv/.tsv/.txt,或用 format 指定)`);
            },
        },
        {
            name: 'office_write_docx',
            description: '生成 Word 文档(.docx):传入 Markdown 文本,支持 # 标题、- 无序列表、1. 有序列表、| 管道表格、普通段落。生成/汇报/简历/合同类文档用本工具。',
            parameters: {
                type: 'object',
                properties: {
                    path: { type: 'string', description: '输出 .docx 路径(相对会话工作目录或绝对路径)' },
                    markdown: { type: 'string', description: 'Markdown 内容(# 标题 / - 列表 / 1. 有序 / | 表格 / 段落)' },
                    title: { type: 'string', description: '文档标题(元数据),可选' },
                },
                required: ['path', 'markdown'],
                additionalProperties: false,
            },
            output: {
                schema: OBJ_SCHEMA,
                render: (_a, v) => {
                    const r = v;
                    return block(`已生成 ${r.path}(${r.paragraphs} 段落, ${r.tables} 表格)`);
                },
            },
            async execute(rawArgs, exec) {
                const args = (rawArgs ?? {});
                const file = resolvePath(requireString(args, 'path', 'path'), exec);
                const markdown = requireString(args, 'markdown', 'markdown');
                ensureParent(file);
                const result = await writeDocx(file, markdown, str(args, 'title') || undefined);
                return result;
            },
        },
        {
            name: 'office_write_xlsx',
            description: '生成 Excel 工作簿(.xlsx):多工作表、表头样式、公式、列宽。数据报表/台账/汇总表用本工具;rows 为二维数组(首行表头)或 data 为对象数组。',
            parameters: {
                type: 'object',
                properties: {
                    path: { type: 'string', description: '输出 .xlsx 路径' },
                    sheets: {
                        type: 'array',
                        description: '工作表数组;每个元素 { name?, rows?, data?, header?, widths? }',
                        items: {
                            type: 'object',
                            properties: {
                                name: { type: 'string' },
                                rows: { type: 'array', items: { type: 'array' } },
                                data: { type: 'array', items: { type: 'object' } },
                                header: { type: 'boolean' },
                                widths: { type: 'array', items: { type: 'number' } },
                            },
                        },
                    },
                },
                required: ['path', 'sheets'],
                additionalProperties: false,
            },
            output: {
                schema: OBJ_SCHEMA,
                render: (_a, v) => {
                    const r = v;
                    const detail = r.written.map(w => `${w.name} ${w.rows} 行`).join(', ');
                    return block(`已生成 ${r.path}(${detail})`);
                },
            },
            async execute(rawArgs, exec) {
                const args = (rawArgs ?? {});
                const file = resolvePath(requireString(args, 'path', 'path'), exec);
                const sheets = requireArray(args, 'sheets', 'sheets');
                ensureParent(file);
                return await writeWorkbook(file, sheets);
            },
        },
        {
            name: 'office_edit_xlsx',
            description: '编辑已有 Excel(.xlsx):ops 操作数组按序执行。支持 set_value(写单元格)、set_style(加粗/填充/对齐/数字格式,cell 或 A1:B2 range)、add_row、add_sheet、delete_sheet。改已有表格用本工具,不要整表重建。注意:value 以 = 开头会被 Excel 当作公式(需要公式时正是此行为)。',
            parameters: {
                type: 'object',
                properties: {
                    path: { type: 'string', description: '目标 .xlsx 路径' },
                    ops: {
                        type: 'array',
                        description: '操作数组',
                        items: {
                            type: 'object',
                            properties: {
                                op: { type: 'string', enum: ['set_value', 'set_style', 'add_row', 'add_sheet', 'delete_sheet'] },
                                sheet: { type: 'string' },
                                cell: { type: 'string', description: '单元格,如 A1' },
                                range: { type: 'string', description: '区域,如 A1:C3(set_style)' },
                                value: { description: 'set_value 的值' },
                                style: {
                                    type: 'object',
                                    properties: {
                                        bold: { type: 'boolean' },
                                        fill: { type: 'string', description: '填充色 ARGB 或 RRGGBB' },
                                        color: { type: 'string' },
                                        align: { type: 'string', enum: ['left', 'center', 'right'] },
                                        format: { type: 'string', description: '数字格式' },
                                    },
                                },
                                row: { type: 'array', description: 'add_row 的一行' },
                                name: { type: 'string', description: 'add_sheet/delete_sheet 的工作表名' },
                                rows: { type: 'array', items: { type: 'array' }, description: 'add_sheet 的初始行' },
                            },
                            required: ['op'],
                        },
                    },
                },
                required: ['path', 'ops'],
                additionalProperties: false,
            },
            output: { schema: OBJ_SCHEMA, render: (_a, v) => block(`已应用 ${v.applied} 个操作`) },
            async execute(rawArgs, exec) {
                const args = (rawArgs ?? {});
                const file = resolvePath(requireString(args, 'path', 'path'), exec);
                const ops = requireArray(args, 'ops', 'ops');
                if (!fileExists(file))
                    throw new Error(`文件不存在: ${file}(edit 只改已有文件;新建请用 office_write_xlsx)`);
                return await editWorkbook(file, ops);
            },
        },
        {
            name: 'office_csv_write',
            description: '写入 CSV 文件:二维数组 rows,自动 RFC 4180 转义(含逗号/引号/换行的字段加引号),CRLF 行尾,Excel 可直接打开。',
            parameters: {
                type: 'object',
                properties: {
                    path: { type: 'string', description: '输出 .csv 路径' },
                    rows: { type: 'array', items: { type: 'array' }, description: '二维数组,首行建议表头' },
                    delimiter: { type: 'string', description: '分隔符,默认 ,(可用 \\t)' },
                },
                required: ['path', 'rows'],
                additionalProperties: false,
            },
            output: {
                schema: OBJ_SCHEMA,
                render: (_a, v) => {
                    const r = v;
                    return block(`已写入 ${r.path}(${r.rows} 行)`);
                },
            },
            async execute(rawArgs, exec) {
                const args = (rawArgs ?? {});
                const file = resolvePath(requireString(args, 'path', 'path'), exec);
                const rows = requireArray(args, 'rows', 'rows');
                const delimiter = str(args, 'delimiter') || config.csvDelimiter;
                ensureParent(file);
                fs.writeFileSync(file, toCsv(rows, delimiter), 'utf8');
                return { path: file, rows: rows.length };
            },
        },
        {
            name: 'office_health',
            description: 'dsh-office-plugin 离线自检:确认插件已加载、依赖引擎(exceljs/mammoth/docx/jszip)可用、当前配置。不访问网络。',
            parameters: { type: 'object', properties: {}, additionalProperties: false },
            output: { schema: OBJ_SCHEMA, render: renderHealth },
            async execute() {
                return {
                    ok: true,
                    plugin: 'dsh-office-plugin',
                    version: PLUGIN_VERSION,
                    engines: { exceljs: 'ok', mammoth: 'ok', docx: 'ok', jszip: 'ok' },
                    config,
                };
            },
        },
        {
            name: 'office_docx_info',
            description: '获取 Word(.docx)排版参数(只读诊断):段落总数、空段落数、行距模式分布(exact 固定/auto 倍数及磅值)、段前段后、字号分布、中文字体分布,可返回逐段明细。诊断“行距太大/字体未安装/空白行多”用本工具,不动文件。',
            parameters: {
                type: 'object',
                properties: {
                    path: { type: 'string', description: '文件路径(相对会话工作目录或绝对路径)' },
                    withParagraphs: { type: 'boolean', description: '是否返回逐段明细,默认 true' },
                    limit: { type: 'number', description: '逐段明细条数上限,默认 50' },
                },
                required: ['path'],
                additionalProperties: false,
            },
            output: { schema: OBJ_SCHEMA, render: (_a, v) => block(renderDocxInfo(v)) },
            async execute(rawArgs, exec) {
                const args = (rawArgs ?? {});
                const file = resolvePath(requireString(args, 'path', 'path'), exec);
                if (!fileExists(file))
                    throw new Error(`文件不存在: ${file}`);
                const info = await inspectDocx(file, Math.max(1, Math.floor(numOr(args, 'limit', 50))));
                if (args.withParagraphs === false)
                    info.items = [];
                return info;
            },
        },
        {
            name: 'office_fix_docx',
            description: '修复 Word(.docx)行距:把所有段落的行距统一改为单倍(single)/倍数(multiple)/固定磅值(exact),可顺带删除纯文本空段落(表格内、含图片、含分节符的段落除外)。只改写 spacing,字体字号等其他格式原样保留;outputPath 另存为新文件;dryRun 只统计不写盘。修复“行距过大/空行太多”用本工具。',
            parameters: {
                type: 'object',
                properties: {
                    path: { type: 'string', description: '目标 .docx 路径' },
                    mode: { type: 'string', enum: ['single', 'multiple', 'exact'], description: 'single 单倍 / multiple 倍数 / exact 固定磅值' },
                    multiple: { type: 'number', description: 'mode=multiple 时的倍数,如 1.15' },
                    linePt: { type: 'number', description: 'mode=exact 时的磅值,如 14' },
                    removeEmptyParagraphs: { type: 'boolean', description: '删除纯文本空段落,默认 false' },
                    outputPath: { type: 'string', description: '另存路径,默认原地覆盖' },
                    dryRun: { type: 'boolean', description: '只统计改动不写盘,默认 false' },
                },
                required: ['path'],
                additionalProperties: false,
            },
            output: { schema: OBJ_SCHEMA, render: (_a, v) => block(renderFixDocx(v)) },
            async execute(rawArgs, exec) {
                const args = (rawArgs ?? {});
                const file = resolvePath(requireString(args, 'path', 'path'), exec);
                if (!fileExists(file))
                    throw new Error(`文件不存在: ${file}`);
                const mode = (str(args, 'mode') || 'single');
                if (mode === 'multiple' && typeof args.multiple !== 'number')
                    throw new Error('mode=multiple 需要 multiple(倍数,如 1.15)');
                if (mode === 'exact' && typeof args.linePt !== 'number')
                    throw new Error('mode=exact 需要 linePt(磅值,如 14)');
                const outPath = str(args, 'outputPath');
                return await fixDocx(file, {
                    mode,
                    multiple: typeof args.multiple === 'number' ? args.multiple : undefined,
                    linePt: typeof args.linePt === 'number' ? args.linePt : undefined,
                    removeEmptyParagraphs: args.removeEmptyParagraphs === true,
                    dryRun: args.dryRun === true,
                    outputPath: outPath ? resolvePath(outPath, exec) : undefined,
                });
            },
        },
    ];
}
function renderHealth(_args, value) {
    const v = value;
    return block(`dsh-office-plugin v${v.version}: exceljs / mammoth / docx / jszip 就绪`);
}
function renderDocxInfo(v) {
    const r = (v ?? {});
    const lines = [`docx 排版参数:${r.path}`, `段落 ${r.paragraphs} 个(空段落 ${r.emptyParagraphs},含图片/对象 ${r.withObjects})`];
    if (r.lineSpacing?.length) {
        lines.push('行距分布:');
        for (const s of r.lineSpacing) {
            lines.push(`  ${s.lineRule}${s.linePt !== null ? ` ${s.linePt}pt` : ''} × ${s.count} 段`);
        }
    }
    if (r.fonts && Object.keys(r.fonts).length > 0) {
        lines.push('中文字体:' + Object.entries(r.fonts).map(([k, n]) => `${k} × ${n}`).join(', '));
    }
    if (r.sizes && Object.keys(r.sizes).length > 0) {
        lines.push('字号分布:' + Object.entries(r.sizes).map(([k, n]) => `${k}pt × ${n}`).join(', '));
    }
    if (r.items?.length) {
        lines.push(`逐段明细(前 ${r.items.length} 条):`);
        for (const it of r.items) {
            const gap = [];
            if (it.beforePt)
                gap.push(`段前${it.beforePt}pt`);
            if (it.afterPt)
                gap.push(`段后${it.afterPt}pt`);
            const line = `${it.lineRule}${it.linePt !== null ? ` ${it.linePt}pt` : ' 默认'}`;
            lines.push(`  #${it.index} [${line}${gap.length ? `, ${gap.join('/')}` : ''}] ${it.sizePt ? `${it.sizePt}pt` : '-'} ${it.eastAsiaFont ?? '-'} | ${it.text || '(空行)'}`);
        }
    }
    return lines.join('\n');
}
function dist(bucket) {
    if (!bucket || Object.keys(bucket).length === 0)
        return '(无)';
    return Object.entries(bucket).map(([k, n]) => `${k} × ${n}`).join(', ');
}
function renderFixDocx(v) {
    const r = (v ?? {});
    const lines = [
        `${r.dryRun ? '[dryRun 未写盘] ' : ''}已修复 ${r.path}`,
        `行距统一为 ${r.rule};改写 ${r.changed} 段,删除空段 ${r.removedEmpty ?? 0} 个`,
        `修复前:${dist(r.before)}`,
        `修复后:${dist(r.after)}`,
    ];
    return lines.join('\n');
}
