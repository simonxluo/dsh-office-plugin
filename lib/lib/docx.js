/** Word(.docx)读写:读取用 mammoth,生成用 docx(均支持 Markdown 子集) */
import fs from 'node:fs';
import mammoth from 'mammoth';
import { AlignmentType, Document, HeadingLevel, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType, } from 'docx';
/** 读取 docx 纯文本(批注/修订/文本框不展开) */
export async function readDocx(file) {
    const { value } = await mammoth.extractRawText({ path: file });
    return value;
}
/** Markdown 子集:# 标题 / -  bullet / 1. 有序 / | 表格 / 普通段落 */
export function parseMarkdown(md) {
    const blocks = [];
    const lines = md.replace(/\r\n/g, '\n').split('\n');
    let i = 0;
    while (i < lines.length) {
        const line = lines[i];
        if (/^\s*\|.*\|\s*$/.test(line) && i + 1 < lines.length && /^\s*\|[\s:|-]+\|\s*$/.test(lines[i + 1])) {
            const cellsOf = (l) => l.trim().replace(/^\|/, '').replace(/\|$/, '').split('|').map(c => c.trim());
            const rows = [cellsOf(line)];
            i += 2;
            while (i < lines.length && /^\s*\|.*\|\s*$/.test(lines[i])) {
                rows.push(cellsOf(lines[i]));
                i++;
            }
            blocks.push({ kind: 'table', rows });
            continue;
        }
        const h = /^(#{1,3})\s+(.*)$/.exec(line);
        if (h) {
            blocks.push({ kind: `h${h[1].length}`, text: h[2].trim() });
            i++;
            continue;
        }
        const b = /^\s*[-*+]\s+(.*)$/.exec(line);
        if (b) {
            blocks.push({ kind: 'bullet', text: b[1].trim() });
            i++;
            continue;
        }
        const n = /^\s*\d+[.)]\s+(.*)$/.exec(line);
        if (n) {
            blocks.push({ kind: 'numbered', text: n[1].trim() });
            i++;
            continue;
        }
        if (line.trim() === '') {
            i++;
            continue;
        }
        blocks.push({ kind: 'p', text: line.trim() });
        i++;
    }
    return blocks;
}
/** 由 Markdown 子集生成 docx(覆盖同路径文件) */
export async function writeDocx(file, markdown, title) {
    const blocks = parseMarkdown(markdown);
    const children = [];
    let paragraphs = 0;
    let tables = 0;
    for (const b of blocks) {
        if (b.kind === 'table' && b.rows) {
            const rows = b.rows.map((cells, idx) => new TableRow({
                tableHeader: idx === 0,
                children: cells.map(c => new TableCell({ children: [new Paragraph(c)] })),
            }));
            children.push(new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }));
            tables++;
            continue;
        }
        const heading = b.kind === 'h1'
            ? HeadingLevel.HEADING_1
            : b.kind === 'h2'
                ? HeadingLevel.HEADING_2
                : b.kind === 'h3'
                    ? HeadingLevel.HEADING_3
                    : undefined;
        children.push(new Paragraph({
            heading,
            bullet: b.kind === 'bullet' ? { level: 0 } : undefined,
            numbering: b.kind === 'numbered' ? { reference: 'office-numbered', level: 0 } : undefined,
            children: [new TextRun(b.text ?? '')],
        }));
        paragraphs++;
    }
    const doc = new Document({
        creator: 'dsh-office-suite',
        title: title && title.trim() ? title.trim() : undefined,
        numbering: {
            config: [
                {
                    reference: 'office-numbered',
                    levels: [{ level: 0, format: 'decimal', text: '%1.', alignment: AlignmentType.START }],
                },
            ],
        },
        sections: [{ children }],
    });
    const buf = await Packer.toBuffer(doc);
    fs.writeFileSync(file, buf);
    return { path: file, paragraphs, tables };
}
