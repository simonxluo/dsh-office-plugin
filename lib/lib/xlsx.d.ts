export interface SheetInput {
    /** 工作表名,缺省 Sheet1/Sheet2… */
    name?: string;
    /** 二维数组(首行默认表头) */
    rows?: unknown[][];
    /** 对象数组(键做表头),与 rows 二选一 */
    data?: Record<string, unknown>[];
    /** 首行是否加粗+底纹表头样式,默认 true */
    header?: boolean;
    /** 列宽(字符数) */
    widths?: number[];
}
export interface CellStyle {
    bold?: boolean;
    /** 填充色,ARGB 十六进制(如 FF DDEBF7 或 DDEBF7) */
    fill?: string;
    /** 字体色,同上 */
    color?: string;
    align?: 'left' | 'center' | 'right';
    /** 数字格式,如 '#,##0.00'、'yyyy-mm-dd' */
    format?: string;
}
export type EditOp = {
    op: 'set_value';
    sheet?: string;
    cell: string;
    value?: unknown;
} | {
    op: 'set_style';
    sheet?: string;
    cell?: string;
    range?: string;
    style: CellStyle;
} | {
    op: 'add_row';
    sheet?: string;
    row?: unknown[];
} | {
    op: 'add_sheet';
    name: string;
    rows?: unknown[][];
} | {
    op: 'delete_sheet';
    name: string;
};
export interface ReadResult {
    sheets: string[];
    /** TSV 文本:[Sheet: 名] 分段 */
    text: string;
    totalRows: number;
}
/** 读取工作表(默认第一个),offset 从 0 开始计数据行 */
export declare function readWorkbook(file: string, sheetName: string | undefined, offset: number, limit: number): Promise<ReadResult>;
/** 新建工作簿并写入(覆盖同路径文件) */
export declare function writeWorkbook(file: string, sheets: SheetInput[]): Promise<{
    path: string;
    written: Array<{
        name: string;
        rows: number;
    }>;
}>;
/** 就地编辑已有工作簿 */
export declare function editWorkbook(file: string, ops: EditOp[]): Promise<{
    path: string;
    applied: number;
}>;
