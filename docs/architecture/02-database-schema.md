# Database Schema — Multi-Playbook Insurance Claims Platform

**Database:** PostgreSQL  
**ORM:** SQLAlchemy 2.x  
**Convention:** snake_case, UUID primary keys, JSONB for flexible structured data

---

## Schema Diagram (Logical)

```
playbooks ──< playbook_rules
          ──< playbook_scoring_dimensions
          ──< playbook_evidence_types
          ──< playbook_flags
          ──< playbook_claim_types
          ──< playbook_rejection_categories
          ──< playbook_ombudsman_patterns
          ──< case_references
          ──< claims

claims ──< claim_evidence
       ──< claim_dimension_scores
       ──< claim_rules_triggered
       ──< claim_flags
       ──< claim_missing_evidence
       ──< claim_outputs
       ──< documents
```

---

## Playbook Tables

### `playbooks`
Master record for each loaded peril playbook.

| Column | Type | Notes |
|---|---|---|
| id | UUID PK | |
| playbook_key | VARCHAR(64) UNIQUE | e.g. `storm-v1`, `flood-v1` |
| peril_type | VARCHAR(64) | e.g. `storm`, `flood`, `fire` |
| version | VARCHAR(16) | e.g. `1.0` |
| jurisdiction | VARCHAR(64) | e.g. `UK` |
| display_name | VARCHAR(128) | e.g. `Storm Damage` |
| status | ENUM | `draft`, `active`, `deprecated` |
| source_file | VARCHAR(256) | Path to source JSON |
| case_count | INT | Number of cases in case database |
| created_at | TIMESTAMPTZ | |
| updated_at | TIMESTAMPTZ | |

---

### `playbook_rules`
Deterministic rules that the rules engine evaluates per claim.

| Column | Type | Notes |
|---|---|---|
| id | UUID PK | |
| playbook_id | UUID FK → playbooks | |
| rule_code | VARCHAR(32) | e.g. `RULE-001` |
| name | VARCHAR(128) | e.g. `Policy Exclusion Gate` |
| priority | INT | Evaluation order (lower = earlier) |
| halt_on_trigger | BOOLEAN | Halt full assessment if triggered |
| condition_json | JSONB | JSONLogic condition object |
| action_json | JSONB | Actions to apply when triggered |
| flags_triggered | JSONB | Array of flag codes |
| checklist_items | JSONB | Array of checklist item strings |
| source_section | VARCHAR(64) | e.g. `Section 9, Rule 1` |

---

### `playbook_scoring_dimensions`
The scoring dimensions for a playbook (Storm has 5 × 20 = 100 total).

| Column | Type | Notes |
|---|---|---|
| id | UUID PK | |
| playbook_id | UUID FK → playbooks | |
| dimension_key | VARCHAR(64) | e.g. `storm_verification` |
| display_name | VARCHAR(128) | e.g. `Storm Verification` |
| max_score | INT | e.g. 20 |
| sort_order | INT | Display and calculation order |
| criteria_json | JSONB | Array of `{min, max, label, description}` |

---

### `playbook_evidence_types`
The evidence matrix for a playbook — types, weights, and requirements.

| Column | Type | Notes |
|---|---|---|
| id | UUID PK | |
| playbook_id | UUID FK → playbooks | |
| evidence_key | VARCHAR(64) | e.g. `structural_engineer_report` |
| display_name | VARCHAR(128) | |
| weight | ENUM | `very_high`, `high`, `medium`, `low` |
| required | BOOLEAN | Mandatory input |
| recommended | BOOLEAN | Strongly recommended |
| gap_risk_level | ENUM | `critical`, `high`, `medium`, `low` |
| gap_instruction | TEXT | What to do if missing |
| notes | TEXT | |

---

### `playbook_flags`
Flag reference table — all possible flags a playbook can raise.

| Column | Type | Notes |
|---|---|---|
| id | UUID PK | |
| playbook_id | UUID FK → playbooks | |
| flag_code | VARCHAR(64) | e.g. `EXCLUDED_CLAIM` |
| meaning | TEXT | |
| severity | ENUM | `info`, `warning`, `critical` |

---

### `playbook_claim_types`
Supported damage/claim types for a playbook.

| Column | Type | Notes |
|---|---|---|
| id | UUID PK | |
| playbook_id | UUID FK → playbooks | |
| claim_type_key | VARCHAR(64) | e.g. `roof_tile_damage` |
| display_name | VARCHAR(128) | |
| scrutiny_level | ENUM | `standard`, `high`, `very_high` |
| exclusion_check_required | BOOLEAN | |
| notes | TEXT | |

---

### `playbook_rejection_categories`
Rejection taxonomy for a playbook, with historical frequency counts.

| Column | Type | Notes |
|---|---|---|
| id | UUID PK | |
| playbook_id | UUID FK → playbooks | |
| category_key | VARCHAR(64) | e.g. `wear_and_tear_deterioration` |
| display_name | VARCHAR(128) | |
| frequency_count | INT | From case database |
| rank | INT | Rank by frequency |

---

### `playbook_ombudsman_patterns`
Named reasoning patterns from Ombudsman decisions (Section 8 of Storm Playbook).

| Column | Type | Notes |
|---|---|---|
| id | UUID PK | |
| playbook_id | UUID FK → playbooks | |
| pattern_number | INT | |
| title | VARCHAR(128) | e.g. `Storm Occurred ≠ Claim Succeeds` |
| description | TEXT | |

---

## Claim Tables

### `claims`
Master claim record. One row per assessment.

| Column | Type | Notes |
|---|---|---|
| id | UUID PK | |
| claim_ref | VARCHAR(64) | User-supplied reference |
| playbook_id | UUID FK → playbooks | Active playbook at assessment time |
| peril_type | VARCHAR(64) | Denormalised for query convenience |
| damage_type | VARCHAR(128) | e.g. `roof_tile_damage` |
| date_of_loss | DATE | |
| property_postcode | VARCHAR(16) | |
| insurer | VARCHAR(128) | |
| policy_ref | VARCHAR(64) | |
| rejection_reason_raw | TEXT | Insurer's stated rejection reason |
| rejection_category_id | UUID FK → playbook_rejection_categories | Nullable until classified |
| status | ENUM | `draft`, `in_progress`, `complete`, `escalated` |
| total_score | INT | 0–100 |
| claim_strength_band | ENUM | `very_weak`, `weak`, `moderate`, `strong`, `very_strong` |
| rejection_risk | ENUM | `very_low`, `low`, `moderate`, `high`, `very_high` |
| halted_by_rule | UUID FK → playbook_rules | Nullable; set if assessment halted early |
| created_at | TIMESTAMPTZ | |
| updated_at | TIMESTAMPTZ | |

---

### `claim_evidence`
Evidence items submitted for a claim, keyed against playbook evidence types.

| Column | Type | Notes |
|---|---|---|
| id | UUID PK | |
| claim_id | UUID FK → claims | |
| evidence_type_id | UUID FK → playbook_evidence_types | |
| present | BOOLEAN | Was this evidence provided? |
| notes | TEXT | Assessor notes |
| file_ref | VARCHAR(256) | Path or storage key |
| added_at | TIMESTAMPTZ | |

---

### `claim_dimension_scores`
Per-dimension scores computed by the scoring engine.

| Column | Type | Notes |
|---|---|---|
| id | UUID PK | |
| claim_id | UUID FK → claims | |
| dimension_id | UUID FK → playbook_scoring_dimensions | |
| score | INT | 0–max_score |
| rationale | TEXT | Why this score was assigned |
| scored_at | TIMESTAMPTZ | |

---

### `claim_rules_triggered`
Audit trail of all rules evaluated for a claim.

| Column | Type | Notes |
|---|---|---|
| id | UUID PK | |
| claim_id | UUID FK → claims | |
| rule_id | UUID FK → playbook_rules | |
| triggered | BOOLEAN | Did the rule fire? |
| trigger_detail | TEXT | Which condition matched |
| evaluated_at | TIMESTAMPTZ | |

---

### `claim_flags`
Flags raised during assessment, with source traceability.

| Column | Type | Notes |
|---|---|---|
| id | UUID PK | |
| claim_id | UUID FK → claims | |
| flag_id | UUID FK → playbook_flags | |
| rule_id | UUID FK → playbook_rules | Nullable — which rule raised this flag |
| triggered_at | TIMESTAMPTZ | |

---

### `claim_missing_evidence`
Generated missing-evidence checklist items.

| Column | Type | Notes |
|---|---|---|
| id | UUID PK | |
| claim_id | UUID FK → claims | |
| evidence_type_id | UUID FK → playbook_evidence_types | |
| checklist_item | TEXT | Specific instruction (e.g. "Obtain Met Office data for SW1A 2AA on 2024-01-15") |
| priority | ENUM | `critical`, `high`, `medium`, `low` |
| resolved | BOOLEAN | Default false |
| resolved_at | TIMESTAMPTZ | Nullable |

---

### `claim_outputs`
All generated outputs for a claim (structured text and AI-generated prose).

| Column | Type | Notes |
|---|---|---|
| id | UUID PK | |
| claim_id | UUID FK → claims | |
| output_type | ENUM | `claim_summary`, `timeline`, `rejection_classification`, `evidence_gap_analysis`, `recommended_actions`, `missing_evidence_checklist`, `complaint_draft`, `fos_draft` |
| content | TEXT | Full output text or JSON string |
| model_used | VARCHAR(64) | e.g. `claude-sonnet-4-6`; NULL for deterministic outputs |
| prompt_tokens | INT | Nullable |
| completion_tokens | INT | Nullable |
| generated_at | TIMESTAMPTZ | |

---

## Case Reference Tables

### `case_references`
Historical case database entries, one row per decided case.

| Column | Type | Notes |
|---|---|---|
| id | UUID PK | |
| playbook_id | UUID FK → playbooks | |
| case_ref | VARCHAR(32) | e.g. `STORM-001`, `DRN-1207086` |
| damage_type | VARCHAR(128) | |
| rejection_reason | TEXT | |
| rejection_category_id | UUID FK → playbook_rejection_categories | |
| outcome | ENUM | `rejected`, `upheld`, `partial`, `mixed` |
| key_insight | TEXT | One-sentence workflow insight |
| source_file | VARCHAR(256) | e.g. `DRN-1207086.pdf` |
| embedding_id | VARCHAR(256) | Reference to vector store entry |

---

## Document Tables

### `documents`
Raw documents — claim files, case PDFs, playbook sources.

| Column | Type | Notes |
|---|---|---|
| id | UUID PK | |
| claim_id | UUID FK → claims | Nullable |
| case_ref_id | UUID FK → case_references | Nullable |
| playbook_id | UUID FK → playbooks | Nullable |
| file_name | VARCHAR(256) | |
| file_path | VARCHAR(512) | Relative to project root |
| file_type | ENUM | `pdf`, `xlsx`, `md`, `txt`, `image` |
| text_extracted | TEXT | Full extracted text |
| page_count | INT | Nullable |
| embedding_id | VARCHAR(256) | Reference to vector store |
| ingested_at | TIMESTAMPTZ | |

---

## Indexes

```sql
-- Claims lookup
CREATE INDEX idx_claims_playbook_id ON claims(playbook_id);
CREATE INDEX idx_claims_status ON claims(status);
CREATE INDEX idx_claims_peril_type ON claims(peril_type);
CREATE INDEX idx_claims_date_of_loss ON claims(date_of_loss);

-- Evidence lookup
CREATE INDEX idx_claim_evidence_claim_id ON claim_evidence(claim_id);
CREATE INDEX idx_claim_evidence_present ON claim_evidence(claim_id, present);

-- Rules audit
CREATE INDEX idx_rules_triggered_claim ON claim_rules_triggered(claim_id);
CREATE INDEX idx_rules_triggered_fired ON claim_rules_triggered(claim_id, triggered);

-- Flags lookup
CREATE INDEX idx_claim_flags_claim ON claim_flags(claim_id);

-- Case DB lookup
CREATE INDEX idx_case_refs_playbook ON case_references(playbook_id);
CREATE INDEX idx_case_refs_outcome ON case_references(outcome);

-- Outputs
CREATE INDEX idx_claim_outputs_claim ON claim_outputs(claim_id);
CREATE INDEX idx_claim_outputs_type ON claim_outputs(claim_id, output_type);
```
