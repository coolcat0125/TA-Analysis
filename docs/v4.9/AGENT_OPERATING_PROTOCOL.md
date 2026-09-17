# TA-Analysis Multi-Agent Operating Protocol

> **Purpose:** make every Agent operate from the same repository state, with consistent sequencing, evidence standards, QA gates and handoff rules across different times, sessions and Agents.

## 1. Single source of truth

The GitHub repository is the authoritative project workspace.

Agents must not treat chat history, private notes, memory, or verbal status as authoritative when repository state is available.

Authoritative order:

1. current Git commit / branch;
2. `docs/v4.9/V4.9_AGENT_STATE.md`;
3. `docs/v4.9/V4.9_BASELINE.md`;
4. `docs/v4.9/V4.9_DATA_CONTRACT.md`;
5. `AGENTS.md`;
6. latest `PROJECT_STATUS.md` and `CHANGELOG.md`;
7. task-specific reports and artifacts.

## 2. Every Agent startup protocol

Before any work:

### STEP 01 — Sync
Fetch/read current repository state and identify the current HEAD.

### STEP 02 — Read
Read all mandatory governance documents listed above.

### STEP 03 — Compare
Compare current HEAD with the Agent's recorded input HEAD.

### STEP 04 — Resolve
If HEAD changed, inspect the intervening commits and re-evaluate task assumptions.

### STEP 05 — Claim
Set the assigned task to `RUNNING` in `V4.9_AGENT_STATE.md` before substantive work.

### STEP 06 — Execute
Work only within the declared task scope.

## 3. No blind continuation

If an Agent resumes after a time gap, it must assume that the repository may have changed.

Never:
- reuse stale counts without recomputation;
- overwrite newer files;
- assume another Agent's task is still current;
- treat a previous conversation as a substitute for repository inspection.

## 4. Parallelism model

Parallel work is allowed only when outputs do not concurrently modify the same canonical artifact.

Safe parallel pattern:

`Discovery Agent -> candidate ledger`

`Independent Discovery Agent -> separate candidate ledger`

Then:

`Verification -> QA -> Orchestrator -> Master`

Canonical master writes are serialized.

## 5. Task ownership

Each task has exactly one accountable Agent.

An Agent may consume outputs from other Agents but must not silently take ownership of another active task.

The Orchestrator may reassign, split, pause, supersede or merge tasks.

## 6. Data promotion gate

No candidate value may become master/verified merely because it looks plausible.

Promotion requires:

1. source/evidence recorded;
2. status assigned;
3. applicable validation passed;
4. conflict check passed;
5. before/after impact known;
6. task report completed;
7. Orchestrator acceptance.

## 7. Large-batch safety

Any operation affecting more than 20 records must produce:

- input count;
- output count;
- changed count;
- rejected count;
- unchanged count;
- before/after audit;
- rollback/recovery reference.

Any unexplained record loss blocks integration.

## 8. Evidence rules

Every externally sourced material value must preserve:

- source;
- source URL;
- source date;
- retrieval date where practical;
- extraction method;
- evidence level;
- verification status.

Conflicting evidence is not silently discarded.

## 9. Fact / inference / scenario separation

All analytical outputs must distinguish:

- **Fact:** directly supported observation.
- **Inference:** derived from evidence under a documented rule.
- **Scenario:** forward-looking assumption or projection.

Scenario values must never enter factual parameter fields without an explicit scenario field.

## 10. QA gates

### Gate G0 — Repository integrity
No accidental deletion or unrelated change.

### Gate G1 — Schema
Data contract remains consistent.

### Gate G2 — Evidence
New material values have traceable evidence.

### Gate G3 — Data quality
No new CRITICAL issue; no unexplained data loss.

### Gate G4 — Business logic
No cross-brand/vehicle contamination; units and physical relationships are valid.

### Gate G5 — Regression
Existing v4.x functionality remains operational.

### Gate G6 — Integration
Orchestrator reviews the complete task package before main integration.

## 11. Handoff protocol

A completed task must leave the repository in a state from which another Agent can continue without private context.

Required handoff:

1. update `V4.9_AGENT_STATE.md`;
2. create/update task report from `AGENT_REPORT_TEMPLATE.md`;
3. record input/output SHA;
4. record changed files;
5. record data/evidence counts;
6. record QA results;
7. record blockers;
8. define the exact next task.

## 12. Status definitions

- TODO: queued and not started.
- RUNNING: actively being executed.
- CANDIDATE: output produced but not yet validated.
- VALIDATING: validation in progress.
- VERIFIED: task output passed its task-level QA.
- MERGED: integrated into main.
- PUBLISHED: released for downstream use.
- BLOCKED: cannot proceed because a dependency/source/access issue is unresolved.
- REJECTED: output failed QA and must not enter master.
- SUPERSEDED: replaced by a newer valid task/output.

## 13. Stale-task rule

A task is stale when its recorded input HEAD differs from current HEAD.

A stale Agent must not continue modifying canonical data until it:

1. reads intervening commits;
2. checks affected files;
3. re-runs relevant baseline checks;
4. records the new input HEAD;
5. confirms that its task remains valid.

## 14. Orchestrator monitoring

At every checkpoint the Orchestrator checks:

- current HEAD;
- active branches;
- open PRs;
- task states;
- stale tasks;
- changed files;
- record counts;
- evidence counts;
- QA results;
- blockers;
- next task.

The Orchestrator's acceptance is required before a task is treated as integrated.

## 15. Copyright

Existing repository copyright descriptions remain in force. Agents must preserve them in applicable repository files and released deliverables.
