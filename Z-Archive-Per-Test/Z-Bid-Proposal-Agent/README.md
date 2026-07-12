记录 Bid Proposal Agent 智能体应用的使用方法。

## 目标

Bid Proposal Agent 用于辅助安全售前/FDE 在投标、POC、方案交付场景中快速生成方案草稿。

当前原型支持三个输入：

- 技术规范书：必传，用于解析项目范围、技术要求、产品能力要求和交付约束。
- 评分标准：必传，用于提取评分项、分值权重、应答重点和得分策略。支持 `.pdf`、`.docx`、`.xlsx`、`.csv`、`.txt`、`.md`；`.xls` 请另存为 `.xlsx` 或 `.csv` 后上传。
- 方案编写模板：可选，用于适配不同项目的章节结构、标题层级、表格样式和固定格式要求，减少生成后的格式调整。未上传时使用内置默认模板兜底。

同时支持生成后的标题调整配置：

- 全局标题操作：不调整、所有标题降一级、所有标题升一级。
- 最大标题级别：限制批量调整后不超过默认模板支持的标题层级。
- 指定章节标题：按章节标题名称匹配，将某个章节单独调整到指定标题级别。
- 冲突策略：指定章节规则优先于全局批量规则。

支持篇幅规划配置：

- 总体编写页数：定义生成文档的目标页数。
- 分配模式：按大纲自动分配、手动指定章节页数、自动分配后手动微调。
- 章节页数：用户可根据生成的大纲，按章节名称指定目标页数。
- 分配策略：手动指定章节页数优先，未指定章节由生成器结合剩余页数自动分配。

支持 Web 页面配置大模型认证：

- Anthropic-compatible API：填写 API Key、请求地址和模型名称。
- OpenAI-compatible API：填写 API Key、请求地址和模型名称。
- Codex CLI：默认使用当前登录态，也可填写指定模型。
- 配置保存到本地 `config/model-config.local.json`，不会提交到仓库。

支持更丰富的大纲与正文输出：

- 大纲可生成二级、三级、四级标题结构。
- Word 正文不再只堆叠长段落，会结合小标题、要点清单、检查表、关键控制点等形式组织内容。
- 正文段落保持首行缩进两个中文字符，标题和表格保持独立样式。

## 运行

启动本地后端服务：

```bash
python3 server.py
```

然后访问：

```text
http://127.0.0.1:8011
```

如果只想查看静态页面，也可以直接打开：

```bash
open src/index.html
```

静态方式只支持本地摘要预览，不会保存上传文件或生成任务输出。

## 模型配置

模型相关配置位于：

```text
config/model-config.json
```

系统会在每次模型调用前重新读取模型配置，因此用户在 Web 页面或本地文件修改 `model-config` 后，下一次章节编写会自动使用最新配置。

默认策略是优先使用配置文件中的模型信息；如果配置了 `env.ANTHROPIC_*`，会优先调用 Anthropic-compatible Messages API。未配置专用 API 时，使用本机 `Codex CLI` 当前登录态和当前默认模型，也就是跟随用户在 Codex 中已经填写/登录的模型信息。

本机私有配置可复制示例文件：

```bash
cp config/model-config.local.example.json config/model-config.local.json
```

`config/model-config.local.json` 已加入 `.gitignore`，可用于本机覆盖模型和 API 地址，不会提交到仓库。出于安全原因，Web 页面保存的 API Key 不再落盘，只保留在当前服务进程内存中；如需长期保存，请优先使用环境变量：

```bash
export OPENAI_API_KEY="你的 API Key"
```

配置优先级：

- 环境变量优先级最高，例如 `ANTHROPIC_MODEL`、`ANTHROPIC_AUTH_TOKEN`、`CODEX_MODEL`、`OPENAI_MODEL`、`OPENAI_BASE_URL`、`OPENAI_API_KEY`。
- 其次读取 `config/model-config.local.json`。
- 再读取 `config/model-config.json`。
- 如果配置文件包含 `env.ANTHROPIC_*`，默认 `providerPriority` 为 `anthropic -> codex -> openai -> local`。
- 否则默认 `providerPriority` 为 `codex -> openai -> local`。

当前支持这种用户登录/模型配置格式：

```json
{
  "env": {
    "ANTHROPIC_AUTH_TOKEN": "不要提交真实密钥",
    "ANTHROPIC_BASE_URL": "https://example.com/api/anthropic/v1",
    "ANTHROPIC_MODEL": "your-model-name"
  }
}
```

服务端会把 `env` 中的变量注入模型调用环境，并用于正文编写。

如果希望强制使用 OpenAI API，可在本机 local 配置中调整：

```json
{
  "writer": {
    "providerPriority": ["openai", "codex", "local"]
  },
  "openai": {
    "model": "gpt-4o-mini",
    "apiKeyEnv": "OPENAI_API_KEY"
  }
}
```

## 子智能体架构

当前 MVP 按 5 个子智能体组织：

1. 材料解析智能体：解析技术规范书和评分标准，抽取关键行和需求/评分信号。
2. 评分标准分析智能体：识别评分项、分值和优先级，形成应答策略。
3. 模板理解智能体：读取上传模板或默认模板，加载标题样式和模板约束。
4. 方案大纲智能体：以评分标准中的方案类评分项为主导生成章节大纲，技术规范书仅作为章节编写结合材料。
5. 文档导出智能体：导出 `proposal-draft.md` 和 `pipeline-result.json`。

大纲生成原则：

- 优先从评分表的技术部分、服务部分提取方案类评分项作为一级章节。
- 自动排除商务资质、人员配置、价格评分、技术规范应答情况等非方案编写章节。
- 每个一级章节会生成不超过 7 个子标题。
- 技术规范书用于补充每章正文写作要点，不再主导一级大纲名称。

流水线入口：

```text
agents/pipeline.py
```

API 创建任务后会输出：

```text
output/{jobId}/manifest.json
output/{jobId}/pipeline-result.json
output/{jobId}/proposal-draft.md
```

## 模板处理原则

当用户上传方案编写模板时，生成流程应优先遵循上传模板约束：

- 保留模板的章节顺序。
- 复用模板的标题层级。
- 识别固定表格和标准段落。
- 将技术规范书和评分标准中的内容填充到对应章节。
- 对缺失信息保留待补充占位，避免生成不确定内容。

如果未上传模板，则使用用户提供并已固化到项目内的 Word 默认模板：

```text
templates/default-proposal-template.docx
```

默认模板用于保证无模板场景下仍有稳定的章节结构和 Word 标题样式。后续可以根据公司标准方案、行业方案或客户常见招标格式持续迭代这个兜底模板。

模板样式已抽取为：

```text
templates/default-template-style.json
```

当前固化的标题样式覆盖 `Heading 1` 到 `Heading 8`，生成器应优先保留这些标题层级、字号、加粗、颜色、段前段后间距和行距配置。
