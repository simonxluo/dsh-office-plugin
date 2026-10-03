/** CSV/TSV 解析与序列化(RFC 4180,零依赖) */
/** 解析 CSV 文本为二维数组。引号内可含分隔符/换行/双写引号 */
export declare function parseCsv(text: string, delimiter?: string): string[][];
/** 序列化为 CSV(CRLF 行尾;含分隔符/引号/换行/首尾空白的字段加引号) */
export declare function toCsv(rows: unknown[][], delimiter?: string): string;
/** 从首行猜分隔符,支持逗号/制表符/分号/竖线 */
export declare function detectDelimiter(sample: string): string;
