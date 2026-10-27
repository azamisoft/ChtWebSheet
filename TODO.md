# CWS Memory Integration TODO

Status: deferred. This is a record for later review, not an instruction to enable these features now.

The current architecture is CWS -> Devlune API -> local Agent -> ChtCortexV2 Harness -> Codex. Memory belongs in the Harness plugin; Devlune remains a transport. Keep the original Cortex interface and existing behavior intact. Memory writes remain disabled.

Review the following legacy Cortex behavior before deciding whether further migration is needed:

- Extract the actual user question from CWS requests instead of searching with the entire protocol envelope.
- Select memory scopes and limit search results and injected context length.
- Read `artifactRef` and turn templates into planning summaries without losing required sheets, headers, or sample-cell positions.
- Handle direct template responses, `_templateArtifact`, and related special cases.
- Generate concise memory records from completed results. Do not enable writes without a separate decision.

The earlier proposed location was an independent `mac-agent/cws/memory/` module. Preserve that proposal for reference only: re-evaluate it against the current Harness-plugin ownership before implementation. Some items may already be covered by the read-only Harness plugin and may need no further changes.
