export interface ParagraphInfo {
    index: number;
    /** 段落文本预览(前 40 字) */
    text: string;
    isEmpty: boolean;
    /** 含图片/图形/书签等对象,不可当空段落删除 */
    hasObject: boolean;
    lineRule: string;
    linePt: number | null;
    beforePt: number | null;
    afterPt: number | null;
    sizePt: number | null;
    eastAsiaFont: string | null;
}
export interface DocxInfo {
    path: string;
    paragraphs: number;
    emptyParagraphs: number;
    withObjects: number;
    /** 行距模式分布,如 [{ lineRule: 'exact', linePt: 17, count: 15 }] */
    lineSpacing: Array<{
        lineRule: string;
        linePt: number | null;
        count: number;
    }>;
    fonts: Record<string, number>;
    sizes: Record<string, number>;
    items: ParagraphInfo[];
}
export type FixMode = 'single' | 'multiple' | 'exact';
export interface FixOptions {
    mode: FixMode;
    /** mode=multiple 时生效(倍数) */
    multiple?: number;
    /** mode=exact 时生效(磅) */
    linePt?: number;
    /** 删除纯文本空段落(表格内/含对象/含 sectPr 的除外) */
    removeEmptyParagraphs?: boolean;
    /** 只统计不写盘 */
    dryRun?: boolean;
    /** 另存路径,默认原地覆盖 */
    outputPath?: string;
}
export interface FixResult {
    path: string;
    dryRun: boolean;
    changed: number;
    removedEmpty: number;
    /** 生效的行距描述,如 'line=240 lineRule=auto(单倍)' */
    rule: string;
    before: Record<string, number>;
    after: Record<string, number>;
}
/** 读取 docx 排版参数(不动文件) */
export declare function inspectDocx(file: string, itemLimit?: number): Promise<DocxInfo>;
/** 修复行距(可选删空段落),返回变更统计;只动 spacing,其他格式不动 */
export declare function fixDocx(file: string, options: FixOptions): Promise<FixResult>;
