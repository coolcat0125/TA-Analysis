# TA-Analysis Orchestrator Runbook

## Mission

Maintain a coherent, auditable, multi-Agent development process across time.

## Before assigning work

1. Read current HEAD.
2. Read `V4.9_AGENT_STATE.md`.
3. Confirm dependencies.
4. Confirm no conflicting active task.
5. Create/identify an isolated task branch.
6. Define acceptance criteria.

## During execution

At each checkpoint:

1. inspect branch/HEAD;
2. inspect changed files;
3. inspect task state;
4. check candidate/master boundary;
5. check evidence counts;
6. check QA;
7. update state if needed.

## On completion

Do not accept an Agent's "done" message by itself.

Require:

- task report;
- changed-file review;
- evidence review;
- QA result;
- input/output SHA;
- regression result;
- explicit next-task handoff.

Then set task state to VERIFIED.

## Before merge to main

Require:

- no unresolved CRITICAL;
- no unexplained data loss;
- no cross-brand contamination;
- data contract compliance;
- reproducible validation;
- task report present;
- state register updated.

Then merge or approve the PR.

## After merge

1. verify main HEAD;
2. re-run release/regression checks;
3. update state from MERGED to PUBLISHED when appropriate;
4. identify next dependency-ready task;
5. archive the prior task report.

## Emergency rule

If a task causes suspected data contamination:

`STOP -> isolate branch -> preserve evidence -> identify affected records -> rollback/revert -> audit -> reopen task`

Never repair by blindly overwriting with a guessed value.

## Copyright

Existing repository copyright descriptions remain in force.
