# MVP Roadmap — Multi-Playbook Insurance Claims Platform

## Guiding Principles

- **Storm first, abstract second** — get a working single-playbook system before building the multi-playbook abstraction
- **Rules engine before AI** — deterministic assessment must be correct before AI outputs are added
- **Test fixtures drive quality** — use real case outcomes from the 38-case database to validate engine correctness
- **CLI before API** — build and validate the core engine as a CLI tool before adding a web API

---

## Phase 1 — Foundation (Weeks 1–3)

**Goal:** Project skeleton, database, and Storm playbook as structured JSON.

### Deliverables

| Task | Output |
|---|---|
| Repo setup | `pyproject.toml`, `requirements.txt`, `config/.env.example` |
| Database setup | PostgreSQL schema via Alembic; all tables from database schema doc |
| Storm playbook JSON | `playbooks/storm/` — all 5 JSON files populated from Markdown playbook |
| Playbook loader | `src/playbooks/loader.py` — load and validate `playbook.json` against schema |
| Playbook registry | `src/playbooks/registry.py` — resolve peril → PlaybookConfig |
| Pydantic models | ClaimInput, ClaimState, PlaybookConfig, Rule, EvidenceItem, DimensionScore |
| CLAUDE.md | Project instructions for Claude Code sessions |

### Definition of Done
- `PlaybookRegistry.resolve("storm")` returns a valid `PlaybookConfig` with 27 rules, 5 dimensions, 12 flags
- All playbook JSON files pass schema validation
- Database migrations run cleanly

---

## Phase 2 — Rules Engine + Scoring (Weeks 4–6)

**Goal:** Deterministic assessment pipeline working end-to-end against Storm playbook.

### Deliverables

| Task | Output |
|---|---|
| JSONLogic evaluator | `src/core/rules_engine/evaluator.py` — evaluate condition objects against ClaimState |
| Rules engine | `src/core/rules_engine/engine.py` — iterate rules, return RuleTriggerResult list |
| Scoring engine | `src/core/scoring_engine/engine.py` — dimension scoring + total + band + risk |
| Evidence processor | `src/core/evidence_processor/processor.py` — classify evidence, apply weights |
| Missing evidence engine | `src/core/evidence_processor/gap_checker.py` — generate prioritised checklist |
| Workflow orchestrator | `src/core/workflow/orchestrator.py` — Stages 1–7 (no AI yet) |
| Test fixtures | `tests/fixtures/sample_claims/` — 5 claim inputs (strong/moderate/weak/halted/partial) |
| Rules engine unit tests | `tests/unit/test_rules_engine.py` — all 27 rules with fixture inputs |
| Scoring unit tests | `tests/unit/test_scoring_engine.py` — all 5 dimensions |

### Definition of Done
- 5 fixture claims produce correct flags, scores, and missing evidence checklists
- All 27 Storm rules trigger / do not trigger correctly on fixture inputs
- Gate rules (Rule 1, Rule 10) halt assessment correctly for fence/excluded damage types
- 0 regressions on subsequent fixture runs

---

## Phase 3 — Deterministic Outputs + CLI (Weeks 7–8)

**Goal:** Complete assessment → structured output pipeline accessible via CLI.

### Deliverables

| Task | Output |
|---|---|
| Output generator (deterministic) | `src/core/output_generator/generator.py` — Stages 8 outputs |
| CLI assess command | `scripts/assess_claim.py` — accepts JSON input, prints full ClaimAssessment |
| Rejection classifier | Map claim inputs → rejection taxonomy category |
| Recommended actions builder | Rule + score derived action list |
| Case ingestion (structured) | `scripts/ingest_cases.py` — parse Section 13 table → case_references rows |
| Integration test | `tests/integration/test_workflow.py` — end-to-end pipeline on all 5 fixtures |

### Sample CLI usage
```bash
python scripts/assess_claim.py --input tests/fixtures/sample_claims/strong_claim.json
# Outputs: ClaimAssessment JSON with all deterministic outputs populated
```

### Definition of Done
- CLI produces valid ClaimAssessment JSON for all 5 fixtures
- rejection_classification, evidence_gap_analysis, recommended_actions, missing_evidence_checklist all populated
- Outputs match expected_outputs fixtures
- Case database loaded: 38 storm cases in case_references table

---

## Phase 4 — AI Layer (Weeks 9–11)

**Goal:** Claude integration for prose outputs. RAG over case database.

### Deliverables

| Task | Output |
|---|---|
| Anthropic SDK client | `src/ai/client.py` — wrapped Claude calls with prompt caching |
| PDF case ingestion | `scripts/ingest_cases.py` extended — pdfplumber + embedding |
| Vector store setup | `src/ai/rag/vector_store.py` — ChromaDB, `storm_cases` collection |
| RAG retriever | `src/ai/rag/retriever.py` — query by claim profile, return top-k cases |
| Prompt builder | `src/ai/prompts/builder.py` — structured prompt from ClaimAssessment state + RAG context |
| Output templates | `src/core/output_generator/templates/` — claim_summary, complaint_draft, fos_draft, timeline |
| AI output generator | `src/core/output_generator/generator.py` — Stage 9 (AI outputs) |
| Updated workflow | Stages 8–9 producing all output types |

### Definition of Done
- claim_summary generated with accurate reflection of assessment state
- complaint_draft grounded in correct rejection category and missing evidence
- fos_draft references relevant precedent cases from the 38-case database
- All AI outputs stored in claim_outputs table with model and token counts
- Prompt caching implemented: playbook knowledge cached across requests

---

## Phase 5 — REST API (Week 12)

**Goal:** HTTP API to support future UI and external integrations.

### Deliverables

| Task | Output |
|---|---|
| FastAPI app | `src/api/main.py` |
| POST /claims | Accept ClaimInput, return ClaimAssessment |
| GET /claims/{id} | Retrieve stored assessment |
| GET /playbooks | List active playbooks |
| GET /health | Health check |
| API request/response schemas | `src/api/schemas/` |
| Integration tests | `tests/integration/test_api.py` |

### Definition of Done
- All 5 fixture claims produce correct assessments via HTTP API
- API documented via FastAPI auto-docs at /docs

---

## Phase 6 — Multi-Playbook Abstraction (Weeks 13–16)

**Goal:** Second peril playbook (Flood or Escape of Water). Proves the architecture is truly peril-agnostic.

### Deliverables

| Task | Output |
|---|---|
| Second playbook JSON | `playbooks/flood/` or `playbooks/escape_of_water/` |
| Playbook ingestion pipeline | `src/ingestion/pipeline.py` — full automated ingestion from Markdown |
| Ingestion CLI | `scripts/ingest_playbook.py` |
| Smoke test framework | Fixture-based smoke test run as part of playbook activation |
| Multi-playbook registry | Registry handles multiple active playbooks simultaneously |
| API peril routing | `POST /claims` routes to correct playbook by `peril_type` |
| Second playbook tests | Fixture claims for second peril |

### Definition of Done
- Two active playbooks running simultaneously
- Assessments for both perils produce correct outputs
- Adding the second playbook required zero changes to `src/core/`
- Ingestion CLI produces a deployable playbook from a new Markdown source document

---

## Future Phases (Post-MVP)

| Phase | Description |
|---|---|
| Web UI | Claims handler interface — input form, assessment display, output export |
| Document ingestion | PDF claim files → extracted ClaimInput fields via Claude |
| Additional playbooks | Escape of Water, Fire, Theft, Subsidence, Accidental Damage |
| Case database expansion | Ongoing ingestion of new FOS decisions |
| Batch assessment | Multiple claims via CSV upload |
| Audit trail | Full change history per claim |
| Export | PDF report generation for complaint/FOS submissions |

---

## Risk Register

| Risk | Mitigation |
|---|---|
| Rule extraction from Markdown is imprecise | Manual review step before playbook activation; smoke tests gate activation |
| AI outputs drift from deterministic assessment | Structured prompt includes full ClaimAssessment state; outputs are post-deterministic |
| JSONLogic condition errors | Unit test every rule condition against fixture inputs |
| Vector search misses relevant precedents | Supplement with metadata filters (damage_type, outcome) alongside semantic search |
| New perils need indicator types not in schema | `ClaimInput.indicators` is an open dict — new keys require no code change |

---

## Key Metrics

| Metric | Target |
|---|---|
| Rules engine accuracy | 100% match on fixture expected outputs |
| Assessment latency (deterministic only) | < 200ms |
| Assessment latency (with AI outputs) | < 8s |
| Relevant case retrieval | Top-3 cases should include correct precedent in ≥ 80% of test assessments |
| Playbook ingestion time | < 5 minutes end-to-end for a new peril |
