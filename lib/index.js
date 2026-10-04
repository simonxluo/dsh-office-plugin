/** dsh-office-plugin 入口:function plugin,Loader 元数据 + 工具注册 */
import { Config as ConfigSchema, resolveConfig } from './config.js';
import { buildOfficeTools } from './tools.js';
export const name = 'tool-office';
export const inject = ['tools'];
export const Config = ConfigSchema;
/** Cordis 激活:把 6 个 office_* 工具注册进宿主 tools registry */
export function apply(ctx, config = {}) {
    const resolved = resolveConfig(config ?? {});
    for (const definition of buildOfficeTools(resolved))
        ctx.tools.register(definition);
}
