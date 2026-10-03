# simon-office-toolkit

DeepSeek Harness 的 Office 工具包:让 agent 自助**读写 Word(.docx)/ Excel(.xlsx)/ CSV**。
纯 Node 实现,跨平台(macOS / Windows / Linux),不装 Office 也能用。

## 工具一览(6 个)

| 工具 | 作用 |
| --- | --- |
| `office_read` | 读取 .docx(mammoth 纯文本)/ .xlsx(exceljs,按工作表+翻页)/ .csv/.tsv/.txt;按扩展名自动识别,扫描件与加密封装会明确报错 |
| `office_write_docx` | Markdown 子集 → .docx:`#` 标题、`-` 无序、`1.` 有序、`\|` 管道表格、段落 |
| `office_write_xlsx` | 生成 .xlsx:多工作表、表头样式(加粗+底纹+冻结首行)、公式、列宽;`rows` 二维数组或 `data` 对象数组 |
| `office_edit_xlsx` | 就地编辑已有 .xlsx:`set_value` / `set_style`(加粗、填充、对齐、数字格式,cell 或 A1:B2)/ `add_row` / `add_sheet` / `delete_sheet`,按序执行 |
| `office_csv_write` | 写 .csv:RFC 4180 转义、CRLF 行尾,Excel 直接打开 |
| `office_health` | 离线自检:插件加载、三个引擎可用性、当前配置 |

路径规则:相对路径按**会话工作目录**解析,也接受绝对路径。写入自动创建父目录。
读取超长会截断并标记 `truncated`;xlsx/csv 用 `offset`/`limit` 翻页。

## 安装

```
dsh plugin --profile <profile> add github:simonxluo/simon-office-toolkit   # GitHub(已带构建产物)
dsh plugin --profile <profile> add link:/path/to/simon-office-toolkit      # 本地开发
```

装完重启 dsh 并**新建会话**(工具清单会话级加载)。

## 配置

在 profile 的 `cordis.patch.yml` 覆盖 `tool-office` 行:

```yaml
- id: tool-office
  config:
    maxReadChars: 60000   # office_read 单次返回字符上限(1000–500000)
    maxSheetRows: 500     # xlsx/csv 单次读取行数上限(10–100000)
    csvDelimiter: ','     # csv 默认分隔符
```

## 开发

```sh
npm install
npm run build     # tsc → lib/(产物提交进仓库,GitHub 安装免构建)
npm test          # 构建 + node --test 离线自测(5 组)
```

结构:`src/index.ts`(Loader 元数据+注册)、`src/tools.ts`(6 个工具契约)、
`src/lib/{csv,xlsx,docx}.ts`(格式引擎)、`src/config.ts`(schemastery 配置)。

## 安全说明

- 插件以 dsh 进程权限运行,读写经 `node:fs`;**相对路径锚定会话工作目录**;
  工具调用本身受会话权限/审批体系管辖
- 无网络访问、无遥测、无安装脚本(`preinstall`/`postinstall` 均无)
- 依赖:exceljs / mammoth / docx / @deepseek-ai/schemastery(均为活跃维护库)

## 许可

MIT
