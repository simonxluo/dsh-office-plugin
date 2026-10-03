/** 会话工作目录:与 dsh-email 同源,exec.agent.session.header.cwd */
export declare function cwdOf(exec: unknown): string;
/** 相对路径按会话工作目录解析;绝对路径原样使用 */
export declare function resolvePath(input: string, exec: unknown): string;
export declare function ensureParent(file: string): void;
export declare function fileExists(file: string): boolean;
export declare function truncate(text: string, maxChars: number): {
    text: string;
    truncated: boolean;
};
