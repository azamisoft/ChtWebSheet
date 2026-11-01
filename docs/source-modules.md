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
- `cws-agent/src/modules/`: panel/session UI, task history/recovery, chat/application, streaming/continuation, context/inventory, shape validation and operation contracts.
- `cws-agent/api/_agent_ops.php`: operation constants and feature-library loader.
- `cws-agent/api/agent-ops/`: response JSON, artifact contracts, intent validation, shape layout, normalization and value/address utilities.
- `cws-agent/api/ai/message.php`: HTTP authentication, request handling and response emission.
- `cws-agent/api/ai/message/`: content routing, planning, document patches, template contracts, semantic review, source/repair, quality checks, provider sessions and prompts.
- `cws-agent/api/_task_state.php`: fenced execution, persisted checkpoints, delivery intent (`cws.delivery_contract`) and completion gates. A confirmed workbook delivery cannot become an answer-only result; unfinished checkpoints and empty artifact results remain incomplete. The model's `_readOnlyAnswer` flag is never completion authority by itself.
- `cws-agent/api/ai/message/attachment-records.php`: independent source-region/record inventory for image-derived information lists, cached against the request, headers and image fingerprints. Source document numbers constrain deduplication; deterministic row/field and requested summary-formula checks run before semantic acceptance and again at the final delivery gate. Missing/unreadable states use task-defined display markers, never uncertain OCR fragments or empty form labels. Unverified inventory stays incomplete. `cws.plan_candidate` preserves a validated plan across downstream extraction failures without claiming execution progress. The Codex provider adds cached native OCR text/coordinates only to the source extraction phase; this is fallible evidence alongside the original images, not an authority that may override readable source content.
- `cws-agent/api/ai/message/attachment-materializer.php`: independently reviews extracted source fields, then compiles one information-list owner into native numeric cells, filters, readable widths and exact summary formulas. A separately planned summary cannot regenerate/overwrite that same table. Rejected source/review candidates remain durable diagnostic evidence; only reviewed facts may be materialized. Both OCR stages reuse task-owned provider threads, and unchanged audited evidence is reused without another model call.
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
