/** 插件配置:schemastery 声明 + 运行期默认值 */
import z from '@deepseek-ai/schemastery';
export const Config = z.object({
    maxReadChars: z.number().min(1000).max(500000).volatile(),
    maxSheetRows: z.number().min(10).max(100000).volatile(),
    csvDelimiter: z.string().volatile(),
});
function clampNumber(value, fallback, min, max) {
    const n = typeof value === 'number' && Number.isFinite(value) ? Math.floor(value) : fallback;
    return Math.min(max, Math.max(min, n));
}
/** 从 bundle 行 config(可能为空对象 / 宿主包装值)解析出生效配置 */
export function resolveConfig(config = {}) {
    return {
        maxReadChars: clampNumber(config?.maxReadChars, 60000, 1000, 500000),
        maxSheetRows: clampNumber(config?.maxSheetRows, 500, 10, 100000),
        csvDelimiter: typeof config?.csvDelimiter === 'string' && config.csvDelimiter.length > 0
            ? config.csvDelimiter
            : ',',
    };
}
