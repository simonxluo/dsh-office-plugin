/** dsh-office-plugin 入口:function plugin,Loader 元数据 + 工具注册 */
import { type OfficeConfig } from './config.js';
export declare const name = "tool-office";
export declare const inject: string[];
export declare const Config: import("@deepseek-ai/schemastery").default;
export type Config = OfficeConfig;
/** Cordis 激活:把 6 个 office_* 工具注册进宿主 tools registry */
export declare function apply(ctx: any, config?: Partial<OfficeConfig>): void;
