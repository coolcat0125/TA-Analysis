# TA-Analysis Agent Collaboration Contract

## 1. Purpose
This repository is a shared, continuously evolving workspace for NEV announcement intelligence, supplier mapping, technical parameters, product portfolios and technology-trend analysis.

All Agents must treat the repository as the single source of truth for project state.

## 2. Operating principle
Read -> Inspect -> Work -> Validate -> Record -> Commit -> Handoff.

An Agent must never assume that a previous Agent's verbal status is authoritative. The authoritative state is the Git history plus current project-state documents and validated artifacts.

## 3. Mandatory startup sequence
1. Read AGENTS.md.
2. Read docs/v4.9/V4.9_AGENT_STATE.md.
3. Read docs/v4.9/V4.9_BASELINE.md.
4. Read docs/v4.9/V4.9_DATA_CONTRACT.md.
5. Read the latest relevant PROJECT_STATUS.md and CHANGELOG.md.
6. Inspect the current branch and HEAD.
7. Check whether another Agent has already changed the same data, schema, script, or output.
8. Never overwrite newer work merely to restore an older local assumption.

## 4. Branch and write policy
- Work normally occurs on a task branch named agent/<phase>-<task>.
- Do not make unrelated changes in the same task.
- Do not rewrite or delete historical archives.
- Do not modify the canonical master dataset without passing the applicable validation gate.
- Candidate data must remain separated from verified/master data.
- Every material change must have a Git commit.
- Prefer small, auditable commits over large opaque commits.
- main is the release/integration line.

## 5. Data integrity rules
1. Never invent a parameter.
2. Every material externally sourced parameter must retain evidence metadata.
3. Distinguish verified, candidate, inferred, conflicted, rejected, and unknown.
4. Preserve source values where possible; normalize only in derived fields.
5. Do not silently overwrite an existing verified value with a weaker source.
6. Any batch operation changing more than 20 records must generate a before/after audit artifact.
7. Any schema or field-semantic change must update V4.9_DATA_CONTRACT.md.
8. Any suspected cross-vehicle or cross-brand propagation must stop the pipeline and enter QA.

## 6. Evidence hierarchy
- A — official government/company primary source.
- B — high-quality industry/database/research source.
- C — professional media/platform.
- D — inference/model-based estimate.

D-level values must never be presented as verified facts.

## 7. Quality gates
A task is not complete until:
- no new CRITICAL issue is introduced;
- affected records have been audited;
- evidence status is recorded;
- source/date are retained where applicable;
- before/after counts are known;
- output is reproducible;
- handoff state is written to V4.9_AGENT_STATE.md.

## 8. Handoff protocol
Each completed Agent task must record:
- Task ID
- Agent role
- Start/end time
- Input HEAD
- Output HEAD
- Files changed
- Records added/changed/rejected
- Evidence added
- Validation commands/results
- Known limitations
- Open blockers
- Recommended next task

Use docs/v4.9/AGENT_REPORT_TEMPLATE.md.

## 9. Conflict protocol
If the current repository differs from the Agent's expected baseline:
STOP -> INSPECT -> REBASE/REPLAN -> CONTINUE.

Do not force-update or overwrite another Agent's newer work.

If two Agents need the same canonical file, the Orchestrator must serialize the integration.

## 10. Orchestrator role
The Orchestrator is the final project-state authority for this workflow. It:
- selects the next task;
- checks dependencies;
- verifies Agent outputs;
- rejects unsupported changes;
- updates project state;
- performs final integration checks;
- tracks blockers and stale tasks.

## 11. Copyright
Existing repository copyright descriptions remain in force. Agents must not remove or conceal them.

Current repository wording includes:
Copyright (c) 2026 David YEAH
and
Copyright © 2026 David YE

Do not alter these statements unless a future explicit project decision changes the standard.

Copyright information must remain in applicable released deliverables.
