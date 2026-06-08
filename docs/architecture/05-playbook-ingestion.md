# Playbook Ingestion Design

## Overview

Playbook ingestion is the process of taking raw source material (Excel, Markdown, PDFs) and producing two outputs:
1. A validated `playbook.json` config ready for the rules engine
2. Vector-embedded case records ready for RAG retrieval

Ingestion runs offline (not during claim assessment). A new playbook is fully authorable from a spreadsheet and a set of case PDFs — no code changes required.

---

## Ingestion Sources

| Source Type | Content | Current Example |
|---|---|---|
| Markdown playbook | Full structured playbook (rules, scoring, evidence, patterns) | `Storm_Damage_Playbook_v1.0.md` |
| Excel playbook | Condensed version of playbook | `Storm Damage Playbook.xlsx` |
| Excel case database | Structured case summary table | `Storm case database.xlsx` |
| PDF raw cases | Full FOS/Ombudsman decision documents | `DRN-*.pdf` |

The Markdown playbook is the authoritative source for structured extraction. Excel is a fallback / cross-check. PDFs are ingested for semantic search (RAG), not for rule extraction.

---

## Ingestion Pipeline Stages

```
Source Documents (knowledge/)
        │
        ▼
┌───────────────────────────────────────────────────────────────┐
│  Stage 1: Document Parsing                                    │
│  IngestionPipeline.parse_sources(peril_type, source_dir)      │
│                                                               │
│  Markdown → section_map: {section_title: section_text}        │
│  Excel → sheet_map: {sheet_name: List[Row]}                   │
│  PDF → text_map: {filename: full_text}                        │
└──────────────────────────────┬────────────────────────────────┘
                               │
                               ▼
┌───────────────────────────────────────────────────────────────┐
│  Stage 2: Structured Extraction                               │
│  PlaybookExtractor.extract(section_map) → RawPlaybookData     │
│                                                               │
│  Extract per section:                                         │
│  § 1  → definition_framework (three-question framework)       │
│  § 2  → claim_types[] (supported + exclusions)                │
│  § 3  → rejection_taxonomy{} (categories + frequencies)       │
│  § 4  → evidence_matrix[] (types + weights)                   │
│  § 5  → scoring_dimensions[] (criteria tables)               │
│  § 6  → missing_evidence_rules[] (gap check logic)            │
│  § 7  → preservation_checklist[]                              │
│  § 8  → ombudsman_patterns[]                                  │
│  § 9  → rules[] (rule blocks → Rule objects)                  │
│  § 10 → input_config{} (required/recommended/optional)        │
│  § 11 → output_config{} (required output types)               │
│  § 12 → flags{} (flag reference table)                        │
│  § 13 → case_references[] (case table rows)                   │
│                                                               │
│  CaseExtractor.extract(pdf_texts) → List[CaseSummary]         │
└──────────────────────────────┬────────────────────────────────┘
                               │
                               ▼
┌───────────────────────────────────────────────────────────────┐
│  Stage 3: Schema Validation                                   │
│  PlaybookValidator.validate(raw_data) → ValidationResult      │
│                                                               │
│  Validate against playbook.schema.json:                       │
│  - All required fields present                                │
│  - Rule condition_json is valid JSONLogic                      │
│  - Scoring dimension scores sum ≤ 100                         │
│  - All flag codes referenced in rules exist in flags table    │
│  - Evidence keys referenced in rules exist in evidence matrix │
│                                                               │
│  Raise PlaybookValidationError with full diff if invalid      │
└──────────────────────────────┬────────────────────────────────┘
                               │
                               ▼
┌───────────────────────────────────────────────────────────────┐
│  Stage 4: JSON Config Serialisation                           │
│  Serialise validated playbook → playbooks/<peril>/            │
│  Write:                                                       │
│  - playbook.json   (master config with metadata)              │
│  - rules.json      (rules array)                              │
│  - scoring.json    (dimensions + criteria)                    │
│  - evidence_matrix.json                                       │
│  - flags.json                                                 │
│                                                               │
│  Write structured cases → data/processed/<peril>_cases.jsonl  │
└──────────────────────────────┬────────────────────────────────┘
                               │
                               ▼
┌───────────────────────────────────────────────────────────────┐
│  Stage 5: Database Registration                               │
│  PlaybookRepo.register(playbook_config) → playbook_id         │
│                                                               │
│  Write rows to:                                               │
│  - playbooks (status=draft)                                   │
│  - playbook_rules                                             │
│  - playbook_scoring_dimensions                                │
│  - playbook_evidence_types                                    │
│  - playbook_flags                                             │
│  - playbook_claim_types                                       │
│  - playbook_rejection_categories                              │
│  - playbook_ombudsman_patterns                                │
│  - case_references                                            │
└──────────────────────────────┬────────────────────────────────┘
                               │
                               ▼
┌───────────────────────────────────────────────────────────────┐
│  Stage 6: Embedding Generation                                │
│  Embedder.embed_cases(case_refs, playbook_id)                 │
│  Embedder.embed_playbook_knowledge(playbook_config)           │
│                                                               │
│  For each case:                                               │
│  Chunk: case_ref + damage_type + outcome + key_insight        │
│  Embed → ChromaDB collection: "{peril}_cases"                 │
│  Store embedding_id back to case_references row               │
│                                                               │
│  For each playbook knowledge chunk (rules, patterns):         │
│  Embed → ChromaDB collection: "{peril}_knowledge"             │
└──────────────────────────────┬────────────────────────────────┘
                               │
                               ▼
┌───────────────────────────────────────────────────────────────┐
│  Stage 7: Smoke Test                                          │
│  Run 3 fixture claim inputs (strong / moderate / weak)        │
│  Assert expected flags, score bands, and rules triggered      │
│  If any assertion fails: set status=draft, alert operator     │
│  If all pass: PlaybookRepo.activate(playbook_id)              │
└──────────────────────────────┬────────────────────────────────┘
                               │
                               ▼
                  Playbook status = ACTIVE
           PlaybookRegistry picks it up on next load
```

---

## Rule Extraction Logic

Section 9 of the Markdown playbook uses a consistent format:
```
### Rule N — <Name>
```
followed by a fenced code block containing pseudocode conditions and actions.

The rule extractor:
1. Splits the section on `### Rule` headers
2. Extracts rule number and name from the header
3. Parses the code block into condition and action components
4. Translates pseudocode conditions into JSONLogic objects
5. Extracts flag codes and checklist items from actions

**Manual review step:** The initial Storm playbook ingestion requires a human to verify the JSONLogic translation is correct before activation. Subsequent playbooks should follow the same format, so the extraction can be automated with high confidence.

---

## Case PDF Extraction

Raw case PDFs (FOS/Ombudsman decisions) are ingested for semantic search only — the structured case table in Section 13 of the Markdown is the authoritative source for structured case data.

PDF ingestion:
1. pdfplumber extracts full text per page
2. Text is chunked into ~500-token overlapping windows
3. Each chunk is embedded and stored in `{peril}_cases` ChromaDB collection
4. Metadata: `case_ref`, `page`, `playbook_id`

This allows the RAG pipeline to retrieve specific passages from actual decisions when generating FOS draft submissions.

---

## Ingestion CLI

```bash
# Ingest Storm playbook from knowledge/
python scripts/ingest_playbook.py --peril storm --source knowledge/playbooks/Storm_Damage_Playbook_v1.0.md

# Ingest case PDFs for Storm
python scripts/ingest_cases.py --peril storm --source knowledge/raw-cases/

# Ingest both in sequence
python scripts/ingest_playbook.py --peril storm --with-cases
```

---

## Adding a New Playbook (Future Peril)

1. Prepare source Markdown following the same 13-section structure as `Storm_Damage_Playbook_v1.0.md`
2. Place in `knowledge/playbooks/<Peril>_Playbook_v1.0.md`
3. Place raw case PDFs in `knowledge/raw-cases/<peril>/`
4. Run: `python scripts/ingest_playbook.py --peril <peril> --with-cases`
5. Review smoke test results
6. Activate: `python scripts/ingest_playbook.py --peril <peril> --activate`

No code changes required unless the new peril introduces a structurally new concept not expressible in the current playbook schema (e.g. a peril that requires multi-property assessment).

---

## Playbook Versioning

- Each playbook has a `version` field (e.g. `1.0`, `1.1`, `2.0`)
- Old versions are set to `status=deprecated` but not deleted
- Claims retain a reference to the `playbook_id` at time of assessment
- The registry always resolves to the highest active version for a given `peril_type`
- Historical claims can be re-assessed against a newer version explicitly
