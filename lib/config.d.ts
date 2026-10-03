/** 插件配置:schemastery 声明 + 运行期默认值 */
import z from '@deepseek-ai/schemastery';
export interface OfficeConfig {
    /** office_read 单次返回的字符上限 */
    maxReadChars: number;
    /** xlsx/csv 单次读取的行数上限(配合 offset 翻页) */
    maxSheetRows: number;
    /** csv 默认分隔符(逗号/制表符/分号) */
    csvDelimiter: string;
}
export declare const Config: z;
/** 从 bundle 行 config(可能为空对象 / 宿主包装值)解析出生效配置 */
export declare function resolveConfig(config?: Partial<OfficeConfig>): OfficeConfig;
