# Source module map

The application and CWS Agent use explicit ES module imports. Production still builds the existing web bundle, standalone runtime, and classic Agent IIFE; there is no runtime source concatenation or `eval` loader.

## Application

- `src/main.js`: browser entry only.
- `src/app-core.js`: shared live bindings and ordered initialization. It owns writable top-level bindings, formula registration, DOM handles, and the boot call.
- `src/modules/formula/`: formula calculations, plugin class factory, autocomplete, names and dialogs.
- `src/modules/agent/`: host operation contracts, normalization and application.
- `src/modules/workbook/`: model import, OOXML, sheet state, macros, saving, export and text measurement.
- `src/modules/drawing/`: shapes, charts, SVG, transforms and connectors.
- `src/modules/editing/`: selection, clipboard, keyboard, undo/history, autofill and formula evaluation.
- `src/modules/rendering/`: grid and canvas rendering.
- `src/modules/data/`: queries, connections, sorting, filters, validation, analysis and forecasts.
- `src/modules/ui/`: ribbon, annotations, protection, dialogs, views and help.
- `src/modules/standalone/`: embedded legacy standalone helper implementation.
- `src/styles.css`: ordered CSS imports. The cascade is preserved in `src/styles/`.

Feature modules contain function declarations, imports and exports only. Shared initialization stays in the core in its original order. Import cycles through live bindings are deliberate during this migration; adding module-level initialization to a feature file can introduce temporal-dead-zone failures and is rejected by a regression test. The formula plugin class is created at its original initialization point through a factory.

This is a source-organization change, not yet a complete elimination of shared application state. Future state isolation should be performed feature by feature with behavior tests.

## CWS Agent

- `cws-agent/src/agent.js`: bundle entry.
- `cws-agent/src/agent-core.js`: shared state and registration.
- `cws-agent/src/modules/`: panel/session UI, task history/recovery, chat/application, streaming/continuation, context/inventory, shape validation and operation contracts. `quota-status.js` computes remaining percentages and coalesces event-driven balance refreshes without inference or interval polling.
- `cws-agent/api/_agent_ops.php`: operation constants and feature-library loader.
- `cws-agent/api/agent-ops/`: response JSON, artifact contracts, intent validation, shape layout, normalization and value/address utilities.
- `cws-agent/api/ai/message.php`: HTTP authentication, request handling and response emission.
- `cws-agent/api/ai/message/`: content routing, planning, document patches, template contracts, semantic review, source/repair, quality checks, provider sessions and prompts.
- `cws-agent/src/modules/provider-settings.js` and `cws-agent/api/_personal_api.php` implement per-user CHT/OpenAI/compatible profiles. Only the Harness holds the private decryption key. Devlune routes personal inference through its separate authenticated personal-messages endpoint without reserving CHT quota. Harness runs template selection and named output contracts using the selected model.
- `api/ai/message/request-intent.php` handles the request-only semantic routing contract. A fenced task stores the structured route and an independent answer. Execution replays that answer through normal history and completion, and loads workbook evidence only for read/write paths. Host literal edits still use their deterministic route. The provider contract is `request_intent` in Harness and Devlune.
- `cws-agent/api/ai/message/cell-proof.php` verifies numeric cell citations and evaluates a single explicit arithmetic read. `formula-cell.php` binds a whole formula assignment to a one-step edit, its source values and the independent browser workbook diff. Both reuse fenced task resources. The second call checks the scoped candidate through the existing `existing_edit_review` read-only contract.
- `cws-agent/api/_quota.php`: stable inference request IDs and fail-closed quota interruptions. `api/ai/quota.php` exposes only the authenticated user’s CWS balance. The billing authority and admin configuration are in Devlune `auth-server/token-quota*.mjs`; `token-quota-periods.mjs` owns activation-based seven-day periods and the audited monthly-ledger migration.
- `cws-agent/api/_task_state.php`: fenced execution, persisted checkpoints, delivery intent (`cws.delivery_contract`) and completion gates. A confirmed workbook delivery cannot become an answer-only result; unfinished checkpoints and empty artifact results remain incomplete. The model's `_readOnlyAnswer` flag is never completion authority by itself.
- `cws-agent/api/ai/message/attachment-records.php`: independent source-region/record inventory for image-derived information lists, cached against the request, headers and image fingerprints. Source document numbers constrain deduplication; deterministic row/field and requested summary-formula checks run before semantic acceptance and again at the final delivery gate. Missing/unreadable states use task-defined display markers, never uncertain OCR fragments or empty form labels. Unverified inventory stays incomplete. `cws.plan_candidate` preserves a validated plan across downstream extraction failures without claiming execution progress. The Codex provider adds cached native OCR text/coordinates only to the source extraction phase; this is fallible evidence alongside the original images, not an authority that may override readable source content.
- `cws-agent/api/ai/message/attachment-materializer.php`: independently reviews extracted source fields, then compiles one information-list owner into native numeric cells, filters, readable widths and exact summary formulas. A separately planned summary cannot regenerate/overwrite that same table. Rejected source/review candidates remain durable diagnostic evidence; only reviewed facts may be materialized. Both OCR stages reuse the CWS session’s provider thread, and unchanged audited evidence is reused without another model call.
- `cws-agent/api/ai/message/attachment-field-audit.php`: builds a compact immutable fragment/quantity inventory; models select fragments and identify comparable quantities, while the host preserves source facts and computes explicit discrepancy notes. `attachment-consistency.php` caches these strict `source_consistency` decisions and checks column scope and explanatory arithmetic separately from OCR. It can select existing composite source fragments or repair a derived explanation, but cannot alter numeric facts, identifiers or uncertainty states. Compiled source plans bypass model re-planning and generic range expansion. Compatible unaccepted checkpoints can be recompiled while preserving progress and archiving the superseded visual evidence; final delivery reads durable source facts even if a later plan loses its metadata.

Source inventories and their independent field review use the `source_records` strict output contract. Keep the CWS `outputContract` envelope, Devlune's `auth-server/cws-codex.mjs` allowlist, and Harness `cht.provider.codex/cws/output-contracts.mjs` synchronized. Every field has a header, value and status; null is permitted for explicitly missing/unreadable evidence and normalized to the task's display marker. Tests verify PHP forwarding, Devlune forwarding and the actual Codex `turn/start.outputSchema` with low effort.

PHP feature libraries are loaded before request handling. Deploy the new libraries before replacing their loader endpoints. The Agent build copies those directories recursively and bundles the frontend as a classic script, preserving hosted and standalone integration.

## Tests and size limit

- `scripts/e2e-real-upload.mjs`: CLI flags and sequential orchestration.
- `scripts/e2e/helpers/`: fixtures and feature smoke-test helpers.
- `scripts/e2e/scenarios/`: ordered integration phases sharing explicit scenario results. Original statement order is retained.
- `cws-agent/test/helpers/source.mjs`: test-only assembly for existing VM/function contract tests; production does not use it.
- `cws-agent/test/helpers/endpoint-source.php`: test-only access to endpoint functions without running HTTP/authentication/provider side effects.

Run `npm run check:source-size`. Authored application, Agent, API, script and test source files must have fewer than 10,000 lines. The main build runs this check; the Agent build also checks its own source tree. Dependencies, generated `dist`/standalone runtime files and packaged archives are not authored source and are excluded.

`cws-agent/test/message-libraries.test.php` loads the production PHP feature entry and checks that every message library is included. Pure function tests alone cannot catch a deployed library omitted from the runtime loader.

For targeted work, search the function name under the relevant module directory and read only the matching section. Reading a file inside a local script does not itself send its contents to a model; text returned to the model becomes input context. Splitting improves navigation and limits unnecessary source reads, but does not by itself reduce CWS inference prompts, screenshot inputs or retry calls.

## Inference protocol boundary

CWS source-image observations are constrained by the Harness provider `visual_source` contract (`ChtCortexV2/plugins/cht.provider.codex/cws/output-contracts.mjs`). Keep its required fields synchronized with the JSON example and validation in `cws-agent/api/_visual_review.php`. Devlune `cws-codex.mjs` forwards the named contract; it does not define the observation fields. Update and test the provider schema before a coordinated plugin reload. Confirm the installed file matches the tested source and the bridge is ready before E2E.

## Codex session continuity

`cws-agent/api/_thread_reuse.php` binds every Codex inference phase to the authenticated `cws_ai_sessions.devlune_session_id`, including follow-up tasks, visual checks and source confirmation. Graphiti relevance classification receives that same provider session. Stage, purpose, step, prompt/schema, workbook, attachment and context-size changes never create a replacement thread. Other backends retain their existing lifecycle. Missing bindings and resume failures preserve incomplete checkpoints.

The Harness Codex provider leaves history compaction to Codex. Interruption or observed compaction invalidates only the incremental evidence dictionary and image references; a current snapshot is reintroduced into the same thread. A delta-based response spanning compaction is rejected as `codex_context_resync_required` rather than accepted with unverifiable references.

Stage instructions are appended as trusted developer items with `thread/inject_items` only when changed (or refreshed after compaction). `thread/resume.developerInstructions` alone does not update the persisted model-visible rules in the installed Codex build; do not rely on it for phase transitions.

CWS declarative requests reuse an independent Codex app-server process per authenticated owner/user/provider-session through `cht.provider.codex/cws/session-runtime.mjs`. The bounded pool defaults to four processes and a two-minute idle timeout (`cws_max_warm_sessions` / `cws_runtime_idle_ms` in the provider's own config). Failures, disconnects, capability upgrades and eviction close only the affected process; recovery resumes its persisted thread. Plugin shutdown closes all of its own leases. File-tool and unknown callers retain their one-request lifecycle. Actual skill file paths are disabled through process-local config overrides, never global desktop configuration. `CURRENT_USER_REQUEST` remains verbatim on every turn; its fixed authority explanation already lives in the declarative base protocol and is no longer repeated in each request.

Account-level unlimited quota is owned by Devlune `auth-server/token-quota.mjs` (`ai_quota_accounts.unlimited`), independently scoped per owner/app/subject. `cws-agent/src/modules/quota-status.js` displays infinity while retaining actual usage; reservations and uncertain-result guards remain active. External component-label collision repair lives in `cws-agent/api/agent-ops/shape-layout.php`: only the label’s own ancestor chain is an overlap exemption, and existing labels prefer nearby valid positions.

Existing shape edits share `src/modules/agent/existing-shape-model.js` between host and Agent. `updateShape` binds one sparse text/style/bounds update to the observed shape revision; native worksheet-pixel bounds retain IDs and layout bindings. `shapes-and-sheet-operations.js` applies native anchors and updates attached connector endpoints in the same undo record. PHP parity and parent containment live in `api/agent-ops/existing-shapes.php` and `_workbook_evidence.php`. `_task_state.php` persists accepted repaired ops and requires rendered acceptance for changed existing shapes even without new attachments; `ai/visual-review.php` uses the latest request/current content for that incremental gate. Incomplete rendered evidence cannot produce a completion receipt.

`_workbook_evidence.php` expands range writes on fully observed existing sheets into cell edits, including proposed writes into previously blank cells, so every changed value enters review and changed rows receive browser-measured, grow-only fitting. `ai/message/document-patches.php` merges targeted visual shape patches into the saved batch; omitted correct cell/text edits are preserved. Old row-clipping findings covered by pending row fitting remain subject to fresh rendered acceptance, rather than being mistaken for measurements of the new candidate. Repairs cannot introduce unrelated blank-cell targets.

Compound values in limited edits receive a focused source-fact audit when source fragments disappear. `_workbook_evidence.php` identifies candidate fragments; `ai/message/document-patches.php` asks the existing strict review contract to judge preservation or explicit authorization using only the request and affected values. Exact wording differences do not automatically reject an edit. Unavailable or malformed review stays incomplete without triggering speculative content repair; the cached acceptance key includes the gate version.

Existing-workbook planning with a confirmed workbook delivery uses `existing_edit_plan`, forwarded through Devlune to the Codex provider strict output schema. It cannot return execution operations during planning; the existing host validation and final semantic/visual/persistence gates still apply. Read-only and fresh-document planning retain their existing contracts. Hosted HTML imports await a durable snapshot before reporting ready, so an immediate reload cannot restore the previous workbook.

For a CWS Codex task, `api/ai/message/local-cell-route.php` first checks whether the entire request is a supported sheet-qualified literal assignment with only preservation or delivery clauses. Attachments, unknown workbook dependencies, ambiguous bindings and unsupported clauses use the model router. Either route can identify one qualified cell and quoted text. `api/ai/message/literal-cell-intent.php` binds the candidate to the exact request and complete single-sheet evidence. Only an empty cell outside structured tables and merged regions is materialized locally. Hints are task-scoped, do not count as progress, and are rechecked at execution; all ordinary semantic, operation, render, persistence and completion gates remain active. Other edits retain ordinary model planning and generation.

A locally parsed literal assignment on a compatible host can use `api/ai/message/literal-diff.php` for semantic acceptance. It rebinds the request, plan and original cell evidence, then compares the complete submitted and normalized operations. The browser independently reconstructs the expected workbook and measured row fit from its baseline. `api/_incremental_render.php` binds visual review to the current ready result and original request. For a proved cell, the host supplies its real pixel crop, value and bounds; an issue naming a different cell cannot authorize repair or completion.

CWS edit routing, semantic review and rendered review use one native structured-output schema in `ChtCortexV2/plugins/cht.provider.codex/cws/edit-review-wire.mjs`. The provider verifies and unwraps only the current stage before returning its original contract. The provider usage ledger records the native schema hash. Rendered review supplies actual pixels for each distinct image in the current batch, including when a previous turn saw identical pixels.

`api/_ai_backend.php` aligns only CWS Codex inference POST transport timeouts to the relay’s bounded end-to-end deadline (570 seconds plus a 30-second HTTP margin). Retrieval planning shares that upstream budget. Other applications, legacy backends and control/checkpoint calls keep their original deadlines.

`api/_quota.php` keeps the logical call ID stable and adds a claim-fence recovery token only after an uncertain-usage quota block. Devlune retains the original pending reservation and atomically funds a separate attempt; concurrent recovery cannot duplicate inference, successful attempt results replay without a new charge, and later reconciliation of the original attempt cannot replace that recovered result. The token is quota control metadata, not model context.

The Codex provider's `cws/usage.mjs` accepts a zero baseline for a recovered turn only when its first observed complete cumulative counters equal its last-request counters. This measures the new turn without assigning zero usage to an older interrupted request. Missing counters, unknown prior totals and counter resets remain partial or unavailable; audited reconciliation can replay an already saved result instead of repeating inference.

Provider isolation disables `agents.enabled` as well as legacy multi-agent feature flags in the CWS child process. The native isolation check verifies the effective configuration without model inference; the owner's desktop settings and other Codex sessions are not modified.

`history-and-task-state.js` marks resumed SSE operations that are received but deferred from the native preview. Final reconciliation in `chat-and-application.js` rebuilds such previews atomically rather than mistaking the logical operation list for applied content. `src/modules/agent/contracts.js` checks the expected operation count before committing; the Agent also verifies the host's actual applied count before saving or acknowledging completion.

Existing-edit review ignores known blank-to-blank cell writes and groups consecutive new blank-cell writes by row, retaining every exact value. Existing facts retain individual checks. Both passing and failing reviews must bind to the host-extracted operation/target coverage; malformed rejection evidence retries only the read-only review and never authorizes a content rewrite. Requested creation permits necessary presentation operations, while actual rendering and persistence remain separate completion gates.

Generated connector routing rejects shared middle segments even when their directions agree. Only genuine same-port fan-out and final fan-in trunks may overlap. Fractional bend candidates follow the native renderer's interpolation and separate lanes without changing directed topology; existing user diagrams are not silently rerouted.
