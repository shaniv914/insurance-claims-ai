# Shared Workflow Engine Design

## Overview

The workflow engine is the orchestrating layer that processes any claim, against any playbook, through a fixed sequence of stages. It is **peril-agnostic**: it receives a playbook config and a claim input, and returns a fully populated `ClaimAssessment` object.

The engine contains no storm-specific logic. Storm-specific knowledge lives entirely in the playbook config loaded by the registry.

---

## Workflow Stages

```
ClaimInput
    │
    ▼
┌─────────────────────────────────────────────────────────────┐
│  Stage 1: Input Validation                                  │
│  Check all mandatory fields present; raise ValidationError  │
│  if any required input missing                              │
└──────────────────────────────┬──────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────┐
│  Stage 2: Playbook Resolution                               │
│  PlaybookRegistry.resolve(peril_type) → PlaybookConfig      │
│  Load rules, scoring dimensions, evidence matrix, flags     │
└──────────────────────────────┬──────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────┐
│  Stage 3: Pre-Assessment Gate Rules                         │
│  Run halt_on_trigger=True rules first (exclusion checks)    │
│  Rule 1: Policy Exclusion Gate                              │
│  Rule 10: Fence Exclusion Check                             │
│                                                             │
│  IF any halt rule triggers → return ClaimAssessment with    │
│  status=HALTED, flag=EXCLUDED_CLAIM, halt_by_rule set       │
│  (skip Stages 4–9)                                          │
└──────────────────────────────┬──────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────┐
│  Stage 4: Evidence Classification                           │
│  EvidenceProcessor.classify(claim_input, playbook)          │
│  For each evidence type in playbook evidence matrix:        │
│  - Mark present / absent                                    │
│  - Assign weight from matrix                                │
│  Returns: List[EvidenceItem]                                │
└──────────────────────────────┬──────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────┐
│  Stage 5: Rules Execution                                   │
│  RulesEngine.evaluate(claim_state, playbook.rules)          │
│  Iterate rules in priority order                            │
│  Evaluate each rule's condition_json against claim_state    │
│  Apply actions for triggered rules:                         │
│  - Add flags                                                │
│  - Add checklist items                                      │
│  - Apply score modifiers                                    │
│  Returns: List[RuleTriggerResult], accumulated flags,       │
│           score modifiers, checklist items                  │
└──────────────────────────────┬──────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────┐
│  Stage 6: Scoring                                           │
│  ScoringEngine.score(claim_state, evidence, rule_results,   │
│                      playbook.scoring_dimensions)           │
│  Compute each dimension score (0–max_score)                 │
│  Apply any rule-triggered score modifiers                   │
│  Sum dimensions → total score                               │
│  Map total to strength band and rejection risk              │
│  Returns: List[DimensionScore], total, band, risk           │
└──────────────────────────────┬──────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────┐
│  Stage 7: Missing Evidence Generation                       │
│  EvidenceProcessor.generate_gaps(evidence_items, playbook,  │
│                                  checklist_from_rules)      │
│  For each absent evidence type:                             │
│  - Look up gap_risk_level and gap_instruction               │
│  - Merge with rule-generated checklist items                │
│  - Deduplicate and sort by priority (critical → low)        │
│  Returns: List[MissingEvidenceItem]                         │
└──────────────────────────────┬──────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────┐
│  Stage 8: Deterministic Output Generation                   │
│  OutputGenerator.generate_deterministic(assessment_state)   │
│  Produce without AI calls:                                  │
│  - rejection_classification (rule-based lookup)             │
│  - evidence_gap_analysis (structured list)                  │
│  - recommended_actions (rule + score derived)               │
│  - missing_evidence_checklist (sorted, prioritised)         │
└──────────────────────────────┬──────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────┐
│  Stage 9: AI Output Generation (Claude)                     │
│  OutputGenerator.generate_ai(assessment_state, rag_context) │
│  RAG: retrieve top-k relevant cases from vector store       │
│  Build structured prompt from assessment state + cases      │
│  Call Claude for:                                           │
│  - claim_summary                                            │
│  - timeline (from input dates and events)                   │
│  - complaint_draft (if applicable)                          │
│  - fos_draft (if applicable)                                │
└──────────────────────────────┬──────────────────────────────┘
                               │
                               ▼
                      ClaimAssessment
```

---

## ClaimInput Schema

The input to the workflow. All fields used directly by rules engine conditions.

```python
class ClaimInput(BaseModel):
    # Identity
    claim_ref: str
    peril_type: str                    # "storm", "flood", etc.
    playbook_key: str | None = None    # Override auto-resolution if needed

    # Core claim details
    damage_type: str                   # e.g. "roof_tile_damage"
    date_of_loss: date
    property_postcode: str
    insurer: str
    policy_ref: str | None = None

    # Rejection context
    rejection_reason_raw: str | None = None
    rejection_category_key: str | None = None

    # Evidence flags (keyed to playbook evidence_key values)
    evidence_present: dict[str, bool]  # e.g. {"weather_data": True, "photos": False}
    evidence_notes: dict[str, str]     # Optional notes per evidence type

    # Claim-specific indicators (used by rule conditions)
    indicators: dict[str, Any]         # e.g. {"has_lifted_tiles": True, "water_ingress": True}

    # Documents attached
    document_refs: list[str] = []
```

---

## ClaimState (Internal)

Built from ClaimInput + evidence classification results; passed to rules engine.

```python
class ClaimState(BaseModel):
    claim_input: ClaimInput
    evidence_items: list[EvidenceItem]
    
    # Convenience accessors used in rule conditions
    def evidence_present(self, key: str) -> bool: ...
    def evidence_weight(self, key: str) -> str: ...
    def indicator(self, key: str) -> Any: ...
```

---

## Rules Engine Condition Format

Conditions are expressed as JSONLogic objects. The evaluator resolves `{"var": "path"}` references against `ClaimState`.

### Examples

```json
// Rule 1 — Policy Exclusion Gate
{
  "condition": {
    "in": [{"var": "damage_type"}, ["fence", "gate", "hedge"]]
  }
}

// Rule 2 — Storm Verification Gate
{
  "condition": {
    "or": [
      {"==": [{"var": "evidence_present.weather_data"}, false]},
      {"==": [{"var": "indicators.weather_below_threshold"}, true]}
    ]
  }
}

// Rule 14 — Lifted Tiles + Weather Data Combination
{
  "condition": {
    "and": [
      {"==": [{"var": "indicators.has_lifted_tiles"}, true]},
      {"==": [{"var": "indicators.water_ingress"}, true]},
      {"==": [{"var": "evidence_present.weather_data"}, true]}
    ]
  }
}
```

---

## Score Modifier System

Rules can modify dimension scores without the scoring engine having peril-specific logic.

```json
// Rule 6 — Maintenance Records Boost
{
  "actions": {
    "score_modifiers": [
      {
        "dimension_key": "maintenance_history",
        "delta": 3,
        "reason": "Maintenance records + contractor confirmation"
      },
      {
        "dimension_key": "causation_strength",
        "delta": 2,
        "reason": "Reduced wear and tear risk"
      }
    ]
  }
}
```

Modifiers are applied by the scoring engine after base dimension scoring. Scores are clamped to `[0, max_score]`.

---

## Halt Behaviour

When a `halt_on_trigger=True` rule fires:

1. Workflow records `halted_by_rule` on the assessment
2. All flags triggered by the halt rule are recorded
3. Missing evidence engine still runs (policy wording may be flagged as critical)
4. A minimal output set is generated: flags, halt reason, policy check instruction
5. AI outputs are **not** generated for halted assessments
6. Assessment status is set to `HALTED`

---

## Concurrency and Statefulness

- The workflow orchestrator is **stateless**: each `assess(claim_input)` call is independent
- Rule evaluation order is deterministic: sorted by `priority` ascending, then by `rule_code` for ties
- No shared mutable state between assessments
- Thread-safe: multiple concurrent assessments can run against the same loaded playbook config

---

## Adding a New Peril

To support a new peril (e.g. Flood):

1. Author `playbooks/flood/` JSON configs with flood-specific rules, scoring, evidence matrix
2. Register via `PlaybookRegistry.register("playbooks/flood/playbook.json")`
3. Pass `peril_type="flood"` in `ClaimInput` — the workflow resolves the rest
4. Zero changes to `Orchestrator`, `RulesEngine`, `ScoringEngine`, or `EvidenceProcessor`

The only code addition required: if a flood claim needs a new **indicator** type (e.g. `flood_depth_cm`), add it to `ClaimInput.indicators` — the engine already handles arbitrary indicator keys.
