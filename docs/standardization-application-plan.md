# CWS HTML 标准化申请计划

本文档把“AI 可读可改的 CWS HTML workbook”整理成可执行的标准化申请路线。

## 申请目标

CWS HTML 的核心主张是：

- 文件本身是 HTML，可由浏览器直接打开。
- 文件中有明确的 CWS 标识，不靠文件名或产品文案猜测。
- workbook 的真实数据位于 `script#websheet-model` 的明文 JSON 中。
- sheet、cell、formula、style、image metadata 等结构保持明文；图片二进制 payload 可以作为 `assets.images` 单独存储，并可选使用 gzip-base64。
- AI 或程序只需要修改这一段 JSON，就能在不重写 runtime 的情况下编辑 workbook。
- `script#cws-ai-instructions` 和公开 AI guide 告诉大模型如何识别、读取、修改和保存。

## 当前第一版交付物

- `docs/cws-html-1.0-draft.md`：英文规格草案，适合作为外部讨论基础。
- `docs/cws-html-application-index.md`：申请材料索引和推荐审阅顺序。
- `docs/standardization-submission-package.md`：外部提交用总览材料。
- `docs/cws-html-explainer.md`：解释器文档，说明问题、方案、用例和成功标准。
- `docs/cws-html-use-cases-and-requirements.md`：用例和需求。
- `docs/cws-html-conformance-checklist.md`：合规检查表。
- `docs/cws-html-test-plan.md`：测试计划。
- `docs/cws-html-implementation-report.md`：当前实现覆盖报告。
- `docs/cws-html-round-trip-report-template.md`：真实导出和 AI 编辑 round-trip 证据模板。
- `docs/cws-html-w3c-cg-form-draft.md`：W3C Community Group 表单草稿。
- `docs/cws-html-incubation-roadmap.md`：孵化路线图。
- `docs/cws-html-open-issues.md`：开放议题清单。
- `docs/cws-html-governance-and-ipr.md`：治理、版本、贡献和许可/IPR 注意事项。
- `docs/cws-html-security-privacy.md`：安全和隐私考虑。
- `public/schema/cws-html-workbook-model-v1.schema.json`：workbook JSON Schema。
- `examples/cws-html/valid/`：合规示例。
- `examples/cws-html/invalid/`：非合规示例，用于说明边界。
- `scripts/validate-cws-html.mjs`：参考验证器。
- `scripts/diff-cws-model.mjs`：round-trip 模型差异检查工具。
- `docs/w3c-community-group-proposal.md`：W3C Community Group 提案起稿。

## 推荐路线

第一阶段：项目内事实规格

1. 在仓库中发布 `CWS HTML 1.0 Draft`。
2. 让 CWS 保存出的 HTML 默认符合该草案。
3. 用验证器检查示例和真实导出的 HTML。
4. 保持 AI guide、JSON Schema、示例文件和实现同步。

第二阶段：开放讨论和实现证明

1. 建立公开 issue/discussion，用于收集 AI 编辑、HTML workbook、spreadsheet interoperability 反馈。
2. 准备最少两个实现证明：
   - Cht WebSheet writer/reader。
   - 独立 validator 或 extractor。
3. 准备 round-trip 测试：
   - CWS 保存 HTML。
   - AI 修改 `script#websheet-model`。
   - CWS 重新打开。
   - 验证数据、公式、格式、图片、图形不被破坏。

第三阶段：W3C 孵化

优先申请或加入 W3C Community Group，因为 CWS HTML 是 HTML profile，和 Web 生态最接近。

需要准备：

- 规格草案 URL。
- 问题陈述：现有 HTML 导出格式不面向 AI 直接编辑 workbook 结构。
- 用例：AI 生成/修改設計書、浏览器内编辑、离线 HTML workbook、结构化 diff。
- 非目标：不替代 Excel/OOXML，不要求浏览器原生实现 spreadsheet。
- 测试套件和 reference validator。
- 至少一个可运行实现。

第四阶段：IANA / media type 判断

短期不建议注册新的媒体类型。CWS HTML 是 HTML 文件，默认继续使用 `text/html` 更容易被浏览器、邮件、服务器和文件系统接受。

如果以后有系统要求区分 CWS HTML，可考虑：

- 仍用 `text/html`，通过 `meta[name="cws:format"]` 和 `script#websheet-model` 识别。
- 或准备 vendor tree media type 注册说明，但注册前要评估浏览器兼容性和实际收益。

第五阶段：ISO/JIS

ISO/JIS 路线适合在已有行业采用后推进。进入该阶段前最好已经具备：

- 稳定规格。
- 多个实现或合作方。
- 明确市场必要性。
- 测试套件。
- 安全和隐私说明。
- 可证明的互操作案例。

## 申请材料清单

- 标准名称：CWS HTML Workbook Format 1.0。
- 范围：HTML profile for AI-readable and AI-editable workbook documents。
- 背景问题：普通 spreadsheet HTML export 通常是渲染结果，不是稳定可编辑 workbook model。
- 技术方案：HTML markers + `script#websheet-model` plain workbook JSON + optional encoded image assets + AI guide。
- 互操作策略：JSON Schema、validator、examples、round-trip tests。
- 兼容策略：未知字段必须保留，内部 workbook model version 与 CWS HTML schema version 分离。
- 安全策略：读取/编辑模型不需要执行脚本；AI 不应修改 runtime。
- 知识产权和许可：规格文档建议使用开放许可，参考实现继续遵循项目许可证。

## 下一步

1. 用真实 CWS 导出的 HTML 跑 `npm run spec:validate` 的同类检查。
2. 将公开网站 `/cws/schema/cws-html-workbook-model-v1.schema.json` 部署出来。
3. 准备 3 到 5 个真实业务样例，覆盖公式、格式、合并、图片、图形和多 sheet。
4. 用本仓库内容起草 W3C Community Group 页面或提案 issue。
5. 确认规格文本、schema、示例、validator、AI guide 的许可和公开仓库位置。
