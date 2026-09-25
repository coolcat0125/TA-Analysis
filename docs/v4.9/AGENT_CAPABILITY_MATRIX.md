# TA-Analysis Agent Capability & Assignment Matrix

## 1. Purpose

Agents are selected by task characteristics, not by model name alone. The Orchestrator should match the task to the capability profile below and require independent QA for consequential outputs.

## 2. Capability profiles

| Profile | Required traits | Suitable work |
|---|---|---|
| Research Scout | broad web retrieval, primary-source finding, recall, query decomposition | supplier/model/technology discovery |
| Evidence Auditor | source criticism, contradiction detection, provenance discipline, conservative judgment | evidence verification and conflict resolution |
| Data Engineer | structured transformation, schema discipline, deterministic processing, reproducibility | dataset construction, normalization, migration |
| Quant Analyst | statistics, aggregation, anomaly detection, uncertainty handling | gap inventory, trend analysis, benchmarking |
| OCR / Vision Analyst | image/table reading, layout understanding, parameter extraction, cross-checking | motor/battery nameplate and announcement OCR |
| Systems Analyst | entity modeling, relationship reasoning, architecture thinking | supplier-product-vehicle mapping |
| Technology Analyst | engineering-domain reasoning, chronology, mechanism analysis | battery/e-drive technology evolution |
| QA / Red Team | adversarial checking, edge cases, regression thinking, skepticism | final verification and contamination detection |
| Report Synthesizer | evidence-grounded synthesis, concise executive communication | DOC/PPT/dashboard narrative |

## 3. Model selection guidance

Use the strongest available reasoning model for:
- architecture;
- schema changes;
- conflict resolution;
- final QA;
- technology inference;
- Orchestrator decisions.

Use a high-context research-capable model for:
- broad source discovery;
- long documents;
- multi-source evidence collection.

Use a fast/low-cost model for:
- deterministic cleanup;
- repetitive candidate extraction;
- formatting;
- simple classification.

Use vision-capable models for:
- OCR;
- screenshots;
- PDF tables;
- vehicle/battery/motor parameter images.

Do not select an Agent merely because it is fast if the task requires source judgment or engineering inference.

## 4. Recommended Agent allocation for P1

### A03 — Battery Supplier Discovery
Profile: Research Scout + Systems Analyst.
Model capability: high web-retrieval recall + strong entity disambiguation.
Output: candidate supplier/evidence ledger only.

### A04 — Battery Supplier Verification
Profile: Evidence Auditor + QA/Red Team.
Model capability: strong reasoning + source criticism + contradiction detection.
Rule: preferably use a model/Agent independent from A03 to reduce correlated errors.

### A05 — Battery Parameter Extraction
Profile: Data Engineer + OCR/Vision Analyst.
Model capability: multimodal extraction + structured output + unit normalization.
Output: candidate parameters with evidence references.

### A06 — Motor Supplier Discovery
Profile: Research Scout + Systems Analyst.
Model capability: broad retrieval + automotive supplier/entity knowledge.

### A07 — Motor Parameter / OCR
Profile: OCR/Vision Analyst + Data Engineer.
Model capability: strong multimodal reading + engineering consistency checks.

### A08 — Evidence Audit
Profile: Evidence Auditor + QA/Red Team.
Model capability: high reasoning reliability and adversarial verification.

### A09 — Battery Product Portfolio
Profile: Systems Analyst + Technology Analyst.
Model capability: long-context synthesis and product-family/generation reasoning.

### A10 — Technology Evolution
Profile: Technology Analyst + Quant Analyst.
Model capability: strong engineering reasoning, chronology and uncertainty separation.

### A11 — QA Automation
Profile: Data Engineer + QA/Red Team.
Model capability: code generation/review, deterministic logic, regression design.

### A12 — Dashboard / Report Integration
Profile: Data Engineer + Report Synthesizer.
Model capability: structured data-to-visual narrative transformation.

## 5. Independence rule

For high-impact fields, discovery and verification should use materially independent reasoning passes.

Example:

A03 finds a battery supplier -> A04 independently verifies it -> A08 audits conflicts -> Orchestrator decides promotion.

Do not allow one Agent to both introduce and unilaterally approve a disputed value.

## 6. Model fallback rule

If the preferred model is unavailable:

1. retain the same capability profile;
2. downgrade task scope if necessary;
3. increase verification;
4. never relax evidence requirements merely because a weaker model is being used.

## 7. Agent performance record

Task reports should record capability profile and model identifier when available. The purpose is reproducibility and post-task quality analysis, not model branding.

## 8. Human / Orchestrator override

Capability matching is advisory. The Orchestrator may reassign work when:
- evidence quality is poor;
- model output shows systematic errors;
- task scope changes;
- repository dependencies change.

## 9. Copyright

Existing repository copyright descriptions remain in force.
