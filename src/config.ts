/** 插件配置:schemastery 声明 + 运行期默认值 */
import z from '@deepseek-ai/schemastery'

export interface OfficeConfig {
  /** office_read 单次返回的字符上限 */
  maxReadChars: number
  /** xlsx/csv 单次读取的行数上限(配合 offset 翻页) */
  maxSheetRows: number
  /** csv 默认分隔符(逗号/制表符/分号) */
  csvDelimiter: string
}

export const Config: z = z.object({
  maxReadChars: z.number().min(1000).max(500000).volatile(),
  maxSheetRows: z.number().min(10).max(100000).volatile(),
  csvDelimiter: z.string().volatile(),
})

function clampNumber(value: unknown, fallback: number, min: number, max: number): number {
  const n = typeof value === 'number' && Number.isFinite(value) ? Math.floor(value) : fallback
  return Math.min(max, Math.max(min, n))
}

/** 从 bundle 行 config(可能为空对象 / 宿主包装值)解析出生效配置 */
export function resolveConfig(config: Partial<OfficeConfig> = {}): OfficeConfig {
  return {
    maxReadChars: clampNumber(config?.maxReadChars, 60000, 1000, 500000),
    maxSheetRows: clampNumber(config?.maxSheetRows, 500, 10, 100000),
    csvDelimiter:
      typeof config?.csvDelimiter === 'string' && config.csvDelimiter.length > 0
        ? config.csvDelimiter
        : ',',
  }
}
