# CWS Agent 五场景验收交接 · 2026-11-06（日本时间）

**独立验收通过 2/5**。仓储和领收书可复用现有证据。三张网页截图各生成了一份设计书，但都还缺最终线上验收。本轮仅整理交接，新增个人 API 调用 0 次。用户将由另一位 AI 接手，尚未指定下一轮付费重试次数；此前建议是每个剩余场景至多一轮，若再次失败，记录具体差异、先离线复现。

## 五个场景

以下相对证据路径均以测试 Mac 上的 `~/Downloads/CWS-five-repair-20261105/` 为根。任务 ID 指 Harness 持久化任务。A/B/C 的 236、238、240 是 CWS 界面会话，与 Harness 的会话标识不同。

| 场景 | 实际状态 | 证据与下一步 |
| --- | --- | --- |
| A · MacBook Air（CWS 会话 236） | 旧任务 `cws-ai-mw3h7j4j-afxku3m2` 为 `completed`，15/15，**独立验收失败**。导出的 14 张表中有 8 张表、17 个字段名偏离 Graphiti 原模板，如 `対象Sheet` 被改成 `対象シート`。 | `A/export.html`、`A/template-check.json`、`A/rejected-header-chrome.png`。新校验能在离线重放时拦住旧文件。保留旧回执和费用，另起新任务，用原图、原要求和空白工作簿完整生成。 |
| B · Wikipedia（CWS 会话 238） | `cws-ai-mw3hwfq0-pazomhzg` 为 `interrupted`，15/15。最后的线上证据仍有 1 处裁切：翻译图标 `文A`。十种语言标题和条目数的文字、颜色已有独立检查。 | `icon-fit-v68/` 保存原失败和修正后的浏览器度量。恢复本任务，核对图标完整可见、标题为蓝色、计数为灰色，再验证保存与重新导入。 |
| C · GitHub 定价（CWS 会话 240） | `cws-ai-mw3ir4af-r5ekeo3x` 为 `interrupted`，15/15。最后线上保留 14 个问题，涉及价格单位半角/全角斜杠和文字颜色。三张卡片已对齐，推荐标签独立，显示条件 1–24 各出现一次。 | `source-pixels/price-unit.png`、`edge-visual-replay/C-native.png`、`transcription-v69/`。原图是 `ユーザー/月`。纠正了 20 处单元格的离线重放仍有颜色问题待验，重点检查深色正文、蓝色链接和三处章节标题。恢复原任务。 |
| 仓储 · 四类产物 | `cws-ai-mw3osku3-8gr40ney` **通过**。9 张表覆盖流程图、时序图、I/F 设计书和 24 条测试用例。导出物通过 156/156 项检查。该场景检索后没有适用模板。 | `BusinessVerified/accepted.html`、`export-audit.json`、`acceptance.json`、`v70-live.json`（同目录）。以 BusinessVerified 为最终产物，Business、BusinessClean、BusinessFinal 是历史尝试。 |
| 日本领收书列表 | `cws-ai-mw3h1hyl-2p5bkc6h` **通过**。三张图中有四张凭证，合计 ¥9,378；两张 ¥739 的是不同凭证，均应计入。税额异常有标注。 | `Receipt/accepted.html`、`export-audit.json`、`acceptance.json`、`v70-live.json`（同目录）。通过真实保存和重新导入检查。 |

## 接手时先看这些文件

1. 工作目录是 `/Users/azami/ChtWebSheetAI`。先读 [AGENTS.md](../AGENTS.md) 和 [源码索引](source-modules.md)。宿主与 `cws-agent/` 是独立 Git 工作区，均有在途改动。
2. 证据根的 `cases.json` 给出原要求、请求哈希、工作簿 document ID、附件路径和 SHA-256。各目录的 `request.txt` 是可直接复用的提示词。本次交接已复核原图哈希。原图、已认证浏览器和任务数据库均在测试 Mac 上；另一台机器单独克隆代码，不能恢复这些线上任务。
3. 查看 `template-physical-v67/deployment.json`（Graphiti 表头及完成回执）、`runtime-scope-v70/deployment.json`（真实宿主运行时）、`transcription-v69/deployment.json` 与 `reversal-v71/deployment.json`（原图文字校正及防反复修改）。这些是本次部署的校验证据，未来版本应重新核对。仓储和领收书已在 v70 重新导入，共 10 张表，无文字裁切。
4. Chrome 的 `CWS-Fix-A.html`、`CWS-Fix-B.html`、`CWS-Fix-C.html` 分别在 A、B、C 目录。这些是根据原空白工作簿、原图标识和服务端检查点恢复的测试入口，不能当作最终产物。最终证据用 CWS 原生保存的 HTML。不要分享个人 API Key、浏览器凭据或完整私有日志。

## 线上复测顺序与完成条件

1. 在独立 Chrome 测试标签页打开 [线上 CWS](https://chtec.co.jp/cws/)，确认提供方为个人 OpenAI API、模型为 `gpt-6-luna`、推理为 `low`。串行跑剩余场景，同账号并发时服务可能返回 `409 codex_request_in_progress`。按 B、C、A 的顺序可先复用已有工作进度。
2. B/C 刷新后，先核对当前工作簿、原附件和 CWS 会话。宿主可能恢复其他标签页最后打开的文件，必要时用“打开”重新导入对应的 `CWS-Fix-*.html`。在项目根目录运行下面命令，看到观察器输出初始状态后，再在对应会话点击「作業を再開」。示例是 B，C 单独执行并将参数改为 C；每次使用唯一日志标签。

   ```sh
   cd /Users/azami/ChtWebSheetAI
   evidence="$HOME/Downloads/CWS-five-repair-20261105"
   "$evidence/begin-resume.sh" B "handoff-$(date +%Y%m%d-%H%M%S)"
   ```

   观察器只读，包含 `--wait-for-change --stop-on-interrupted`。等待稳定终态：视觉拒绝可能短暂显示 `failed`，随后才保存证据并变为 `interrupted`。出现相同失败，先用保存的样本重放，再决定下一步。

3. A 的旧完成回执已经写入，不能用作修复后的验证。为新尝试建立独立证据目录（例如 `AFresh/`），绑定新 task ID、实际 document ID、原提示词与原图 SHA。旧 `A/` 保持可追溯。现有 `observe.py` 会按请求哈希写回旧目录；新 A 须按新 task ID 独立观察。现有审计脚本固定读取 A/B/C 的目录，验证新 A 之前先将**场景 A**与**本次证据目录**分开传入，避免覆盖旧导出、任务快照、回执或计量。
4. 每份设计书都要验证 Graphiti 模板选中且引用可追溯、14 个章节和 129 个模板字段齐全，原生表的位置及表头精确匹配。用原图对照布局、文本、金额或计数、颜色和业务含义，查看**本次最终文档**实际渲染的全部工作表。通过后用 CWS「名前を付けて保存」保存，回读并重新导入，核对 document ID、单元格、原生表和图形一致且没有文字裁切。若出现部分失败、终止或丢失回执，确认界面保持未完成并保留检查点。
5. `audit/verify-design.py` 检查实际导出、模板字段和最终版本逐表视觉通过记录；`audit/design-reimport.mjs` 从该场景的 `export.html` 用线上宿主验证重新导入。两项均需成功，并将 `reimport-render.json` 与场景验收记录一起保留。`audit/wiki-caption-colors.php`、`audit/github-card-layout.php` 是补充专项检查。脚本处理的是已保存样本，不能代替本次页面与原图的对照。
6. 若需要再改代码，先离线复现。部署 CWS API 前执行 `npm --prefix cws-agent run build` 和 `node scripts/check-cws-api-build.mjs`，核对 `cws-agent/dist` 的实际变更上传、回读哈希与线上健康状态。涉及共享 Harness 重载时先协调，等到 `bridge.ready` 后才开始付费测试。

## 调用量与最终账单

当前 A、B、C 的 `meter.json` 合计有 **604 次有用量记录的调用**（A 172、B 187、C 245）。其中语义/视觉复检 418 次，局部修复 113 次，生成 42 次，校验修复 17 次，其他 14 次。604 是调用次数，不是 token 数或美元金额。尚未确定的请求保留为待对账。

五项全部验收后再发布费用。更新个人 OpenAI 项目中 CWS 所用 Key 的用量和成本导出，与各场景日志逐分钟核对输入、缓存命中、输出、缓存写入和实际美元费用，并区分实扣 API credits 与按费率估算的金额。现有 `audit/reconcile-usage.py` 按旧目录检查五份回执，需要先接入 A 的新尝试和实际重导入结果。分别列出成功生成和此前失败/重试的消耗，按请求去重。之前只有部分生成步骤的基线，不能据此断言完整任务已经节省多少费用。
