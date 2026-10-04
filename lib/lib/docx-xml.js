/** docx XML 级检查与行距修复:只动 <w:spacing>,其他格式原样保留 */
import fs from 'node:fs';
import jszip from 'jszip';
const PARA_RE = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
const OBJECT_RE = /<w:(?:drawing|pict|object)\b/;
const SECTPR_RE = /<w:sectPr\b/;
function parseAttrs(tag) {
    const attrs = {};
    const re = /w:(\w+)="([^"]*)"/g;
    let m;
    while ((m = re.exec(tag)) !== null)
        attrs[m[1]] = m[2];
    return attrs;
}
function unescapeText(s) {
    return s
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&amp;/g, '&');
}
function paragraphText(p) {
    return unescapeText(p.replace(/<[^>]+>/g, '')).replace(/\s+/g, ' ').trim();
}
async function readDocumentXml(file) {
    const zip = await jszip.loadAsync(fs.readFileSync(file));
    const entry = zip.file('word/document.xml');
    if (!entry)
        throw new Error('不是有效的 .docx(缺少 word/document.xml)');
    return await entry.async('string');
}
/** 计算新 spacing 的 line/lineRule 属性 */
function spacingAttrsOf(mode, multiple, linePt) {
    if (mode === 'single')
        return { line: '240', lineRule: 'auto' };
    if (mode === 'multiple') {
        const m = typeof multiple === 'number' && multiple > 0 ? multiple : 1;
        return { line: String(Math.round(240 * m)), lineRule: 'auto' };
    }
    const pt = typeof linePt === 'number' && linePt > 0 ? linePt : 12;
    return { line: String(Math.round(pt * 20)), lineRule: 'exact' };
}
/** 在 pPr 内插入 spacing(遵守 CT_PPr 子元素顺序:spacing 在 ind/jc 之前) */
function insertSpacingIntoPpr(p, spacingXml) {
    const pprOpen = /<w:pPr>/.exec(p);
    if (!pprOpen || pprOpen.index === undefined) {
        const openTag = /^<w:p\b[^>]*>/.exec(p);
        if (!openTag)
            return p;
        return openTag[0] + `<w:pPr>${spacingXml}</w:pPr>` + p.slice(openTag[0].length);
    }
    const end = p.indexOf('</w:pPr>', pprOpen.index);
    if (end === -1)
        return p;
    const pprInner = p.slice(pprOpen.index, end + '</w:pPr>'.length);
    const anchor = /<w:(?:ind|jc|textAlignment|outlineLvl|rPr)\b/.exec(pprInner);
    const nextPpr = anchor && anchor.index !== undefined
        ? pprInner.slice(0, anchor.index) + spacingXml + pprInner.slice(anchor.index)
        : pprInner.replace(/<\/w:pPr>$/, spacingXml + '</w:pPr>');
    return p.slice(0, pprOpen.index) + nextPpr + p.slice(pprOpen.index + pprInner.length);
}
/** 替换已有 <w:spacing/> 的 line/lineRule,保留 before/after 等其他属性 */
function replaceSpacing(spacingTag, line, lineRule) {
    const parts = Object.entries({ ...parseAttrs(spacingTag), line, lineRule }).map(([k, v]) => `w:${k}="${v}"`);
    return `<w:spacing ${parts.join(' ')}/>`;
}
/** 表格区域 <w:tbl>…</w:tbl> 的字符区间,删空段落时需跳过 */
function tableSpans(xml) {
    const spans = [];
    const re = /<w:tbl>[\s\S]*?<\/w:tbl>/g;
    let m;
    while ((m = re.exec(xml)) !== null)
        spans.push([m.index, m.index + m[0].length]);
    return spans;
}
function ruleKeyOf(p) {
    const ppr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(p)?.[0] ?? '';
    const sp = /<w:spacing\b[^>]*\/>/.exec(ppr);
    const a = sp ? parseAttrs(sp[0]) : {};
    return `${a.lineRule ?? '未设置'}${a.line ? `/${Number(a.line) / 20}pt` : ''}`;
}
/** 读取 docx 排版参数(不动文件) */
export async function inspectDocx(file, itemLimit = 50) {
    const xml = await readDocumentXml(file);
    const items = [];
    let index = 0;
    for (const p of xml.match(PARA_RE) ?? []) {
        const ppr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(p)?.[0] ?? '';
        const spacing = /<w:spacing\b[^>]*\/>/.exec(ppr);
        const attrs = spacing ? parseAttrs(spacing[0]) : {};
        const sz = /<w:sz w:val="(\d+)"/.exec(p);
        const font = /w:eastAsia="([^"]+)"/.exec(p);
        const text = paragraphText(p);
        items.push({
            index: index++,
            text: text.slice(0, 40),
            isEmpty: text === '',
            hasObject: OBJECT_RE.test(p),
            lineRule: attrs.lineRule ?? '',
            linePt: attrs.line ? Number(attrs.line) / 20 : null,
            beforePt: attrs.before ? Number(attrs.before) / 20 : null,
            afterPt: attrs.after ? Number(attrs.after) / 20 : null,
            sizePt: sz ? Number(sz[1]) / 2 : null,
            eastAsiaFont: font ? font[1] : null,
        });
    }
    const lineSpacing = new Map();
    const fonts = {};
    const sizes = {};
    for (const it of items) {
        const key = `${it.lineRule || '(未设置)'}|${it.linePt ?? ''}`;
        const cur = lineSpacing.get(key) ?? { lineRule: it.lineRule || '(未设置)', linePt: it.linePt, count: 0 };
        cur.count++;
        lineSpacing.set(key, cur);
        if (it.eastAsiaFont)
            fonts[it.eastAsiaFont] = (fonts[it.eastAsiaFont] ?? 0) + 1;
        if (it.sizePt)
            sizes[String(it.sizePt)] = (sizes[String(it.sizePt)] ?? 0) + 1;
    }
    return {
        path: file,
        paragraphs: items.length,
        emptyParagraphs: items.filter(i => i.isEmpty).length,
        withObjects: items.filter(i => i.hasObject).length,
        lineSpacing: [...lineSpacing.values()].sort((a, b) => b.count - a.count),
        fonts,
        sizes,
        items: items.slice(0, Math.max(1, itemLimit)),
    };
}
/** 修复行距(可选删空段落),返回变更统计;只动 spacing,其他格式不动 */
export async function fixDocx(file, options) {
    const zip = await jszip.loadAsync(fs.readFileSync(file));
    const entry = zip.file('word/document.xml');
    if (!entry)
        throw new Error('不是有效的 .docx(缺少 word/document.xml)');
    const xml = await entry.async('string');
    const { line, lineRule } = spacingAttrsOf(options.mode, options.multiple, options.linePt);
    const ruleName = options.mode === 'single' ? '单倍' : options.mode === 'multiple' ? `${options.multiple ?? 1} 倍` : `${options.linePt ?? 12}pt 固定`;
    const rule = `line=${line} lineRule=${lineRule}(${ruleName})`;
    const before = {};
    const spans = tableSpans(xml);
    const inTable = (offset) => spans.some(([s, e]) => offset >= s && offset < e);
    let changed = 0;
    let removedEmpty = 0;
    const spacingXml = `<w:spacing w:line="${line}" w:lineRule="${lineRule}"/>`;
    const out = xml.replace(PARA_RE, (p, offset) => {
        const key = ruleKeyOf(p);
        before[key] = (before[key] ?? 0) + 1;
        if (options.removeEmptyParagraphs) {
            const removable = paragraphText(p) === '' && !OBJECT_RE.test(p) && !SECTPR_RE.test(p) && !inTable(offset);
            if (removable) {
                removedEmpty++;
                return '';
            }
        }
        const existing = /<w:spacing\b[^>]*\/>/.exec(p);
        if (existing) {
            const replaced = replaceSpacing(existing[0], line, lineRule);
            if (replaced !== existing[0])
                changed++;
            return p.split(existing[0]).join(replaced);
        }
        changed++;
        return insertSpacingIntoPpr(p, spacingXml);
    });
    const after = {};
    for (const p of out.match(PARA_RE) ?? []) {
        const key = ruleKeyOf(p);
        after[key] = (after[key] ?? 0) + 1;
    }
    const target = options.outputPath ?? file;
    if (!options.dryRun) {
        // 必须经 zip 重新打包:直接写 XML 字符串会得到非 docx 文件
        zip.file('word/document.xml', out);
        const buf = await zip.generateAsync({ type: 'nodebuffer' });
        fs.writeFileSync(target, buf);
    }
    return { path: target, dryRun: options.dryRun === true, changed, removedEmpty, rule, before, after };
}
