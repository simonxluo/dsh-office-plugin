/** 读取 docx 纯文本(批注/修订/文本框不展开) */
export declare function readDocx(file: string): Promise<string>;
type MdBlockKind = 'h1' | 'h2' | 'h3' | 'p' | 'bullet' | 'numbered' | 'table';
interface MdBlock {
    kind: MdBlockKind;
    text?: string;
    rows?: string[][];
}
/** Markdown 子集:# 标题 / -  bullet / 1. 有序 / | 表格 / 普通段落 */
export declare function parseMarkdown(md: string): MdBlock[];
/** 由 Markdown 子集生成 docx(覆盖同路径文件) */
export declare function writeDocx(file: string, markdown: string, title?: string): Promise<{
    path: string;
    paragraphs: number;
    tables: number;
}>;
export {};
