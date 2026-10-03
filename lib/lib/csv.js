/** CSV/TSV 解析与序列化(RFC 4180,零依赖) */
/** 解析 CSV 文本为二维数组。引号内可含分隔符/换行/双写引号 */
export function parseCsv(text, delimiter = ',') {
    const rows = [];
    let row = [];
    let field = '';
    let inQuotes = false;
    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (inQuotes) {
            if (ch === '"') {
                if (text[i + 1] === '"') {
                    field += '"';
                    i++;
                }
                else {
                    inQuotes = false;
                }
            }
            else {
                field += ch;
            }
            continue;
        }
        if (ch === '"') {
            inQuotes = true;
            continue;
        }
        if (ch === delimiter) {
            row.push(field);
            field = '';
            continue;
        }
        if (ch === '\r')
            continue;
        if (ch === '\n') {
            row.push(field);
            rows.push(row);
            row = [];
            field = '';
            continue;
        }
        field += ch;
    }
    if (field.length > 0 || row.length > 0) {
        row.push(field);
        rows.push(row);
    }
    return rows.filter(r => r.length > 1 || r[0] !== '');
}
/** 序列化为 CSV(CRLF 行尾;含分隔符/引号/换行/首尾空白的字段加引号) */
export function toCsv(rows, delimiter = ',') {
    const lines = rows.map(row => row
        .map(cell => {
        const s = cell === null || cell === undefined ? '' : String(cell);
        const needQuote = s.includes(delimiter) || s.includes('"') || s.includes('\n') || s.includes('\r') || /^\s|\s$/.test(s);
        return needQuote ? '"' + s.replace(/"/g, '""') + '"' : s;
    })
        .join(delimiter));
    return lines.join('\r\n') + '\r\n';
}
/** 从首行猜分隔符,支持逗号/制表符/分号/竖线 */
export function detectDelimiter(sample) {
    const firstLine = sample.split(/\r?\n/, 1)[0] ?? '';
    let best = ',';
    let bestCount = 0;
    for (const d of [',', '\t', ';', '|']) {
        const count = firstLine.split(d).length - 1;
        if (count > bestCount) {
            bestCount = count;
            best = d;
        }
    }
    return best;
}
