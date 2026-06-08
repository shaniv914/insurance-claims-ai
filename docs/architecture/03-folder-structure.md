# Folder Structure — Multi-Playbook Insurance Claims Platform

## Full Directory Tree

```
Insurance-Claims-AI/
│
├── src/                            # All application source code
│   ├── core/                       # Peril-agnostic engine components
│   │   ├── __init__.py
│   │   ├── workflow/               # Claim assessment orchestrator
│   │   │   ├── __init__.py
│   │   │   ├── orchestrator.py     # Main workflow: input → assessment output
│   │   │   └── models.py           # ClaimInput, ClaimAssessment Pydantic models
│   │   ├── rules_engine/           # Deterministic rule evaluator
│   │   │   ├── __init__.py
│   │   │   ├── engine.py           # RulesEngine class
│   │   │   ├── evaluator.py        # JSONLogic condition evaluator
│   │   │   └── models.py           # Rule, RuleTriggerResult models
│   │   ├── scoring_engine/         # 0–100 multi-dimension scorer
│   │   │   ├── __init__.py
│   │   │   ├── engine.py           # ScoringEngine class
│   │   │   └── models.py           # DimensionScore, TotalScore models
│   │   ├── evidence_processor/     # Evidence classification + gap analysis
│   │   │   ├── __init__.py
│   │   │   ├── processor.py        # EvidenceProcessor class
│   │   │   ├── gap_checker.py      # Missing evidence checklist generation
│   │   │   └── models.py           # EvidenceItem, EvidenceGap models
│   │   └── output_generator/       # Structured output production
│   │       ├── __init__.py
│   │       ├── generator.py        # OutputGenerator class
│   │       ├── templates/          # Output prompt templates per output type
│   │       │   ├── claim_summary.txt
│   │       │   ├── complaint_draft.txt
│   │       │   ├── fos_draft.txt
│   │       │   └── timeline.txt
│   │       └── models.py           # ClaimOutput model
│   │
│   ├── playbooks/                  # Playbook management
│   │   ├── __init__.py
│   │   ├── registry.py             # PlaybookRegistry: load, validate, resolve
│   │   ├── loader.py               # Load playbook JSON → PlaybookConfig
│   │   ├── validator.py            # Validate playbook config against schema
│   │   └── models.py               # PlaybookConfig, Rule, ScoringDimension, etc.
│   │
│   ├── ingestion/                  # Playbook and document ingestion pipeline
│   │   ├── __init__.py
│   │   ├── pipeline.py             # IngestionPipeline orchestrator
│   │   ├── parsers/
│   │   │   ├── __init__.py
│   │   │   ├── pdf_parser.py       # pdfplumber-based PDF text extraction
│   │   │   ├── excel_parser.py     # openpyxl-based Excel extraction
│   │   │   └── markdown_parser.py  # Markdown section extractor
│   │   ├── extractors/
│   │   │   ├── __init__.py
│   │   │   ├── playbook_extractor.py   # Raw doc → structured PlaybookConfig
│   │   │   └── case_extractor.py       # Raw PDF → CaseReference record
│   │   └── embedder.py             # Chunk text → vector embeddings
│   │
│   ├── ai/                         # Claude integration
│   │   ├── __init__.py
│   │   ├── client.py               # Anthropic SDK client wrapper
│   │   ├── rag/
│   │   │   ├── __init__.py
│   │   │   ├── retriever.py        # Vector store query → relevant cases
│   │   │   └── vector_store.py     # ChromaDB adapter
│   │   └── prompts/
│   │       ├── __init__.py
│   │       └── builder.py          # Build structured prompts from assessment state
│   │
│   ├── api/                        # FastAPI application
│   │   ├── __init__.py
│   │   ├── main.py                 # App entry point
│   │   ├── routes/
│   │   │   ├── __init__.py
│   │   │   ├── claims.py           # POST /claims, GET /claims/{id}
│   │   │   ├── playbooks.py        # GET /playbooks, GET /playbooks/{key}
│   │   │   └── health.py           # GET /health
│   │   ├── schemas/                # FastAPI request/response Pydantic schemas
│   │   │   ├── __init__.py
│   │   │   ├── claim_request.py
│   │   │   └── claim_response.py
│   │   └── dependencies.py         # Shared FastAPI dependencies (DB session, registry)
│   │
│   ├── db/                         # Database layer
│   │   ├── __init__.py
│   │   ├── connection.py           # SQLAlchemy engine + session factory
│   │   ├── models/                 # SQLAlchemy ORM models
│   │   │   ├── __init__.py
│   │   │   ├── playbook.py         # Playbook, PlaybookRule, etc.
│   │   │   └── claim.py            # Claim, ClaimEvidence, ClaimOutput, etc.
│   │   ├── repositories/           # Data access layer
│   │   │   ├── __init__.py
│   │   │   ├── claim_repo.py
│   │   │   └── playbook_repo.py
│   │   └── migrations/             # Alembic migration scripts
│   │       ├── env.py
│   │       └── versions/
│   │
│   └── utils/
│       ├── __init__.py
│       ├── logging.py
│       └── config.py               # Pydantic settings (env vars)
│
├── playbooks/                      # Playbook configuration files (data, not code)
│   ├── _template/                  # Template for authoring a new playbook
│   │   ├── playbook.json           # Master playbook config
│   │   ├── rules.json              # Rules array
│   │   ├── scoring.json            # Scoring dimensions + criteria
│   │   ├── evidence_matrix.json    # Evidence types + weights
│   │   └── flags.json              # Flag definitions
│   └── storm/
│       ├── playbook.json           # Storm playbook master config
│       ├── rules.json              # 27 rules from Section 9
│       ├── scoring.json            # 5 dimensions from Section 5
│       ├── evidence_matrix.json    # Evidence matrix from Section 4
│       └── flags.json              # 12 flags from Section 12
│
├── knowledge/                      # Source knowledge (unchanged from current)
│   ├── case-databases/
│   │   └── Storm case database.xlsx
│   ├── playbooks/
│   │   ├── Storm Damage Playbook.xlsx
│   │   └── Storm_Damage_Playbook_v1.0.md
│   └── raw-cases/
│       └── DRN-*.pdf               # 35 raw case PDFs
│
├── data/                           # Extracted/processed data
│   ├── case_db_raw.txt
│   ├── playbook_raw.txt
│   └── processed/                  # Ingestion outputs
│       ├── storm_cases.jsonl       # Structured case records
│       └── storm_playbook.json     # Validated playbook config
│
├── docs/                           # Architecture and design documents
│   ├── architecture/
│   │   ├── 01-system-architecture.md
│   │   ├── 02-database-schema.md
│   │   ├── 03-folder-structure.md  ← this file
│   │   ├── 04-workflow-engine.md
│   │   └── 05-playbook-ingestion.md
│   ├── schemas/
│   │   ├── playbook.schema.json
│   │   ├── claim-assessment.schema.json
│   │   ├── evidence-item.schema.json
│   │   └── rule.schema.json
│   └── roadmap/
│       └── mvp-roadmap.md
│
├── tests/
│   ├── unit/
│   │   ├── test_rules_engine.py
│   │   ├── test_scoring_engine.py
│   │   └── test_evidence_processor.py
│   ├── integration/
│   │   ├── test_workflow.py
│   │   └── test_api.py
│   └── fixtures/
│       ├── sample_claims/          # JSON test claim inputs
│       └── expected_outputs/       # Expected assessment outputs
│
├── scripts/
│   ├── ingest_playbook.py          # CLI: ingest a new playbook source doc
│   ├── ingest_cases.py             # CLI: ingest case PDFs into vector store
│   └── assess_claim.py             # CLI: run a claim assessment from JSON file
│
├── config/
│   ├── .env.example                # Environment variable template
│   └── logging.yaml
│
├── .gitignore
├── pyproject.toml                  # Project metadata + dependencies
├── requirements.txt
└── CLAUDE.md                       # Claude Code project instructions
```

---

## Key Design Decisions

### `playbooks/` is data, not code
Playbook JSON files live outside `src/`. A new playbook is added by dropping files into `playbooks/<peril>/` — no code changes required.

### `knowledge/` is read-only source material
Original Excel, Markdown, and PDF files stay in `knowledge/`. The ingestion pipeline reads from here and writes structured output to `data/processed/`.

### `core/` has no playbook awareness
The rules engine, scoring engine, and evidence processor receive their configuration from the playbook registry at runtime. They contain no storm-specific logic.

### `src/ingestion/` is separate from `src/core/`
Ingestion is a one-time / on-demand pipeline. It does not run during claim assessment.

### `tests/fixtures/` drives deterministic testing
Rules engine and scoring engine tests use fixed JSON claim inputs and compare against expected outputs — no AI calls needed for unit tests.
