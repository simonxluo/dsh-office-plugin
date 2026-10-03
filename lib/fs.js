/** 文件系统助手:路径解析(相对会话工作目录)、截断 */
import fs from 'node:fs';
import path from 'node:path';
/** 会话工作目录:与 dsh-email 同源,exec.agent.session.header.cwd */
export function cwdOf(exec) {
    const cwd = exec
        ?.agent?.session?.header?.cwd;
    return typeof cwd === 'string' && cwd.length > 0 ? cwd : process.cwd();
}
/** 相对路径按会话工作目录解析;绝对路径原样使用 */
export function resolvePath(input, exec) {
    const p = String(input);
    return path.isAbsolute(p) ? p : path.resolve(cwdOf(exec), p);
}
export function ensureParent(file) {
    fs.mkdirSync(path.dirname(file), { recursive: true });
}
export function fileExists(file) {
    try {
        return fs.statSync(file).isFile();
    }
    catch {
        return false;
    }
}
export function truncate(text, maxChars) {
    if (text.length <= maxChars)
        return { text, truncated: false };
    return { text: text.slice(0, maxChars), truncated: true };
}
