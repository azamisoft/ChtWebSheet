#!/usr/bin/env python3
"""Read-only, change-driven observer for CWS durable task checkpoints."""
import argparse
import json
import sqlite3
import tempfile
import time
from pathlib import Path


def failure_codes(value):
    codes = set()
    if isinstance(value, dict):
        if isinstance(value.get("code"), str):
            codes.add(value["code"])
        for child in value.values():
            codes.update(failure_codes(child))
    elif isinstance(value, list):
        for child in value:
            codes.update(failure_codes(child))
    return codes


def snapshot(db, task_id, output_dir):
    row = db.execute("select body from tasks where task_id=?", (task_id,)).fetchone()
    if not row:
        return {"taskId": task_id, "status": "not_found"}
    task = json.loads(row[0])
    if task.get("app_id") != "cws_agent":
        raise ValueError("Task is not owned by cws_agent")
    checkpoint = task.get("checkpoint") or {}
    resources = task.get("resources") or {}
    codes = set()
    for key, resource in resources.items():
        if key.startswith(("cws.step_draft.", "cws.sheet_repair.")):
            codes.update(failure_codes(resource.get("validation")))
    visual = resources.get("cws.visual_failure") or {}
    if output_dir:
        output_dir.mkdir(parents=True, exist_ok=True)
        target = output_dir / (task_id + "-task.json")
        # A resume observer and the original read-only watcher may overlap.
        with tempfile.NamedTemporaryFile("w", dir=output_dir, suffix=".tmp", delete=False) as stream:
            temp = Path(stream.name)
            try:
                json.dump(task, stream, ensure_ascii=False, indent=2)
            except BaseException:
                temp.unlink(missing_ok=True)
                raise
        temp.replace(target)
    return {"taskId": task_id, "status": task.get("status"),
            "leaseExpired": bool(task.get("status") == "running" and task.get("lease_until")
                                 and task["lease_until"] <= time.time() * 1000),
            "completed": checkpoint.get("completed"), "total": checkpoint.get("total"),
            "failure": task.get("failure"), "draftFailureCodes": sorted(codes),
            "savedVisualIssues": len(visual.get("issues") or [])}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("task_ids", nargs="+")
    parser.add_argument("--db", type=Path, default=Path.home() / "Library/Application Support/ChtCortexV2/plugin-data/cht.memory.task_state/tasks.sqlite")
    parser.add_argument("--snapshot-dir", type=Path)
    parser.add_argument("--once", action="store_true")
    parser.add_argument("--wait-for-change", action="store_true", help="Ignore the initial terminal state until a resumed task changes")
    parser.add_argument("--stop-on-interrupted", action="store_true", help="Return when a visual rejection interrupts the task so its saved defects can be inspected")
    parser.add_argument("--stop-on-expired-lease", action="store_true",
                        help="Return a diagnostic when a running task lease stays expired; do not change task state")
    parser.add_argument("--interval", type=float, default=5)
    parser.add_argument("--timeout", type=float, default=900)
    parser.add_argument("--terminal-settle", type=float, default=6, help="Require a non-success terminal snapshot to stay unchanged before returning")
    args = parser.parse_args()
    if args.interval <= 0 or args.timeout <= 0 or args.terminal_settle < 0:
        parser.error("interval and timeout must be positive; terminal-settle must be nonnegative")
    db = sqlite3.connect(args.db.resolve().as_uri() + "?mode=ro", uri=True, timeout=5)
    previous = None
    observed_change = not args.wait_for_change
    started = time.monotonic()
    terminal_signature = None
    terminal_since = None
    while True:
        states = [snapshot(db, task_id, args.snapshot_dir) for task_id in args.task_ids]
        signature = json.dumps(states, sort_keys=True)
        if signature != previous:
            if previous is not None:
                observed_change = True
            print(json.dumps({"event": "task_state", "tasks": states}, ensure_ascii=False), flush=True)
            previous = signature
        if args.once:
            return 0
        terminal = {"completed", "failed", "cancelled", "not_found"}
        if args.stop_on_interrupted:
            terminal.add("interrupted")
        # Read-scope routing releases its lease before execution claims the
        # same task. This handoff is not a failed or completed workflow.
        preflight_handoff = any(state["status"] == "interrupted" and state.get("failure") == "preflight_released" for state in states)
        expired = args.stop_on_expired_lease and any(s.get("leaseExpired") for s in states)
        settled = all(s["status"] in terminal or (expired and s.get("leaseExpired")) for s in states)
        if observed_change and not preflight_handoff and settled:
            if all(state["status"] == "completed" for state in states):
                return 0
            # Visual rejection currently records reject/claim/evidence/interrupt
            # as separate durable mutations. Do not stop on its transient
            # "failed" state before the final repairable interruption is saved.
            if terminal_signature != signature:
                terminal_signature, terminal_since = signature, time.monotonic()
            elif time.monotonic() - terminal_since >= args.terminal_settle:
                if expired:
                    print(json.dumps({"event": "worker_lease_expired", "tasks": states}, ensure_ascii=False), flush=True)
                    return 4
                return 2
        else:
            terminal_signature = terminal_since = None
        if time.monotonic() - started >= args.timeout:
            print(json.dumps({"event": "observer_timeout", "tasks": states}, ensure_ascii=False), flush=True)
            return 3
        time.sleep(args.interval)


if __name__ == "__main__":
    raise SystemExit(main())
