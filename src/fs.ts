/** 文件系统助手:路径解析(相对会话工作目录)、截断 */
import fs from 'node:fs'
import path from 'node:path'

/** 会话工作目录:与 dsh-email 同源,exec.agent.session.header.cwd */
export function cwdOf(exec: unknown): string {
  const cwd = (exec as { agent?: { session?: { header?: { cwd?: string } } } } | undefined)
    ?.agent?.session?.header?.cwd
  return typeof cwd === 'string' && cwd.length > 0 ? cwd : process.cwd()
}

/** 相对路径按会话工作目录解析;绝对路径原样使用 */
export function resolvePath(input: string, exec: unknown): string {
  const p = String(input)
  return path.isAbsolute(p) ? p : path.resolve(cwdOf(exec), p)
}

export function ensureParent(file: string): void {
  fs.mkdirSync(path.dirname(file), { recursive: true })
}

export function fileExists(file: string): boolean {
  try {
    return fs.statSync(file).isFile()
  } catch {
    return false
  }
}

export function truncate(text: string, maxChars: number): { text: string; truncated: boolean } {
  if (text.length <= maxChars) return { text, truncated: false }
  return { text: text.slice(0, maxChars), truncated: true }
}
