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

PHP feature libraries are loaded before request handling. Deploy the new libraries before replacing their loader endpoints. The Agent build copies those directories recursively and bundles the frontend as a classic script, preserving hosted and standalone integration.

## Tests and size limit

- `scripts/e2e-real-upload.mjs`: CLI flags and sequential orchestration.
- `scripts/e2e/helpers/`: fixtures and feature smoke-test helpers.
- `scripts/e2e/scenarios/`: ordered integration phases sharing explicit scenario results. Original statement order is retained.
- `cws-agent/test/helpers/source.mjs`: test-only assembly for existing VM/function contract tests; production does not use it.
- `cws-agent/test/helpers/endpoint-source.php`: test-only access to endpoint functions without running HTTP/authentication/provider side effects.

Run `npm run check:source-size`. Authored application, Agent, API, script and test source files must have fewer than 10,000 lines. The main build runs this check; the Agent build also checks its own source tree. Dependencies, generated `dist`/standalone runtime files and packaged archives are not authored source and are excluded.

For targeted work, search the function name under the relevant module directory and read only the matching section. Reading a file inside a local script does not itself send its contents to a model; text returned to the model becomes input context. Splitting improves navigation and limits unnecessary source reads, but does not by itself reduce CWS inference prompts, screenshot inputs or retry calls.

## Inference protocol boundary

CWS source-image observations are constrained by the Harness provider `visual_source` contract (`ChtCortexV2/plugins/cht.provider.codex/cws/output-contracts.mjs`). Keep its required fields synchronized with the JSON example and validation in `cws-agent/api/_visual_review.php`. Devlune `cws-codex.mjs` forwards the named contract; it does not define the observation fields. Update and test the provider schema before a coordinated plugin reload. Confirm the installed file matches the tested source and the bridge is ready before E2E.
