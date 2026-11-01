# Execution workflow

- Preserve the user's selected model and reasoning effort. Reduce unnecessary work, not required acceptance checks.
- Read `docs/source-modules.md` to locate a feature. Use targeted searches, then read the relevant functions. Do not dump whole entry files, task snapshots, logs, or chat histories into model context.
- Batch independent checks and browser actions. Summarize structured outputs locally before returning them to the model; keep full evidence on disk.
- Observe long-running CWS tasks with `scripts/observe-cws-tasks.py`. Do not repeatedly ask the model to poll unchanged status. The observer must not resume, repair, or mark tasks complete.
- Wait for a stable terminal snapshot: visual rejection can briefly pass through `failed` before its evidence is saved and the task becomes `interrupted`. Do not diagnose that intermediate state as the final outcome.
- For a resumed acceptance run, start the observer before clicking resume with `--wait-for-change --stop-on-interrupted`; consume its initial state and terminal result. Do not wait only for a future log change when a failure may already have arrived.
- Reproduce a saved failure offline before another paid model run. After a fix, run focused checks first, then the required integration/visual acceptance. Reuse passing evidence only when its dependencies are unchanged.
- Stop repeated runs of the same failure and inspect the cause. Do not skip template coverage, content/diagram consistency, actual rendering, persistence, recovery, or partial-failure completion guards to save tokens.
- Coordinate shared Harness restarts and browser use with already-authorized tasks. Preserve unrelated dirty files and existing checkpoints.
- For inference protocol changes, verify every hop (CWS, Devlune validation/forwarding, Harness and Codex output schema) before resuming a paid acceptance case. After reload, wait for an observed ready state before reporting deployment success.


- Before any CWS API deployment, run `npm --prefix cws-agent run build` and `node scripts/check-cws-api-build.mjs`. The deploy source is `cws-agent/dist`, not authored `api` files. Require actual changed-file upload/hash verification; an empty deployment with only a backup directory is not proof that a source fix is live.
- When changing an inference response shape, inspect the provider contract in `ChtCortexV2/plugins/cht.provider.codex/cws/output-contracts.mjs` as well as CWS PHP validation and Devlune forwarding. A prompt example cannot add fields forbidden by the active strict schema. Verify the installed provider file and observed `bridge.ready` after coordinated reload before paid E2E.
