# 57站多语言翻译 Code Agent 部署说明

本仓库用于在 Code Agent 平台上运行 57 站多语言翻译流水线。仓库只保存代码、规则和说明，不保存历史 Excel、JSONL、产出文件；运行时由脚本从文案系统 Headless API 拉取最新数据并生成本次产物。

## 目录约定

```text
01_原始数据/     运行时下载的文案系统 Excel
02_过程数据/     JSONL 断点文件、质检报告
03_产出/         翻译后 Excel
04_代码/         执行脚本
05_上传格式/     文案系统上传格式文件
```

其中除 `04_代码/` 外，其余目录默认不提交到 Git。

## 平台环境变量

可参考 `.env.example` 配置到 Code Agent 管理平台的环境变量或密钥管理中。

```bash
DIFY_API_KEY=你的 Dify workflow key
I18N_WORKDIR=仓库根目录，可不填，脚本会自动识别
DIFY_PROXY=http://127.0.0.1:30001/rest
MCP_PROXY=http://127.0.0.1:30001/rest
LCC_GW_NAME=language-configuration-center
MCP_PROXY_DC=central_p2_dc
ULP_COOKIE=ulp_token={{ulp_token}}
RUN_DATE=20260907
```

`DIFY_API_KEY` 必须放在平台密钥或环境变量里，不要写进代码。`ULP_COOKIE` 生产环境默认用 `ulp_token={{ulp_token}}`，由 mcp-proxy 注入真实登录态。

## 典型执行流程

推荐让 Code Agent 调用统一入口：

```bash
python3 04_代码/i18n_agent.py export-translate --langs da fi lv --empty-langs da fi lv --engine hybrid
```

如果需要拆开执行，也可以使用底层脚本。

1. 导出最新数据：

```bash
python3 04_代码/export_from_api.py --empty-langs da fi lv
```

2. 翻译：

```bash
python3 04_代码/translate_multi_site.py --langs da fi lv --engine hybrid --google-workers 2 --google-interval 0.8 --dify-workers 4
```

3. 生成上传格式：

```bash
python3 04_代码/build_upload_for_lang_20260612.py --lang lv --run-date 20260907
```

## Code Agent 提示词建议

```text
你是 57 站多语言翻译流水线执行 Agent。
任务目标：根据用户提供的语种、输入文件或导出条件，完成文案系统多语言补翻、质检和上传格式生成。

执行规则：
1. 先确认目标语种、任务模式、是否需要从文案系统导出、RUN_DATE。
2. 不修改原始输入文件，只在过程目录和产出目录写新文件。
3. 翻译必须使用 JSONL 断点文件，支持中断续跑。
4. 默认使用 hybrid 模式：Google 优先，Dify 兜底；如用户指定 dify，则只用 Dify。
5. 严格检查 EN 列不变、空翻译、占位符、HTML 标签、公式风险和英文残留。
6. 对 PH_MISMATCH、HTML 结构异常、语义风险项，输出明细并要求人工确认。
7. 最终返回产出文件路径、质检报告路径、失败项统计和待人工确认项。
8. 禁止硬编码凭据，所有密钥从环境变量读取。
```
