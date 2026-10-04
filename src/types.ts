/** dsh-office-plugin 共享类型 */

/** 工具输出的模型可见渲染块(最小文本块) */
export interface TextBlock {
  type: 'text'
  text: string
}

/** 工具执行上下文(exec)中我们关心的部分:会话工作目录 */
export interface OfficeExec {
  agent?: {
    session?: {
      header?: {
        cwd?: string
      }
    }
  }
}

export interface OfficeToolDefinition {
  name: string
  description: string
  parameters: Record<string, unknown>
  output: {
    schema: Record<string, unknown>
    render(args: unknown, value: unknown): TextBlock[]
  }
  execute(args: unknown, exec?: unknown): Promise<unknown>
}
