# System Architecture — Multi-Playbook Insurance Claims Platform

## Overview

The platform is a **peril-agnostic AI-assisted claims assessment engine**. It applies a deterministic rules engine and a structured scoring model derived from a playbook, augmented by Claude for non-deterministic output generation (complaint drafts, FOS submissions, summaries).

The Storm Damage Playbook (v1.0) is the first playbook. Every architectural decision is made to support future perils (Flood, Escape of Water, Fire, Subsidence, Theft, etc.) without rewriting core engine logic.

---

## Architectural Principles

| Principle | Implication |
|---|---|
| Peril-agnostic engine | Rules, scoring, and evidence logic are loaded from playbook configs, not hardcoded |
| Deterministic rules first | Rules engine runs without AI; Claude only generates prose outputs |
| Playbook as data | A new peril = a new JSON playbook config, not new code |
| Evidence-first design | Every assessment is anchored in evidence classification and gap analysis |
| UK FOS jurisdiction | All frameworks and outputs are calibrated to Financial Ombudsman Service patterns |

---

## Component Map

```
┌─────────────────────────────────────────────────────────────────┐
│                        INTERFACE LAYER                          │
│  ┌──────────────┐   ┌──────────────┐   ┌──────────────────────┐ │
│  │   REST API   │   │     CLI      │   │   Web UI (future)    │ │
│  └──────┬───────┘   └──────┬───────┘   └──────────┬───────────┘ │
└─────────┼──────────────────┼──────────────────────┼─────────────┘
          │                  │                       │
          ▼                  ▼                       ▼
┌─────────────────────────────────────────────────────────────────┐
│                      ORCHESTRATION LAYER                        │
│                   Claim Assessment Workflow                      │
│  Input Validation → Playbook Resolution → Gate Checks →         │
│  Evidence Classification → Rules Execution → Scoring →          │
│  Flag Aggregation → Missing Evidence → Output Generation        │
└───────────────────┬─────────────────────────────────────────────┘
                    │
        ┌───────────┼───────────┐
        ▼           ▼           ▼
┌──────────────┐ ┌──────────┐ ┌──────────────────────┐
│  Rules       │ │ Scoring  │ │  Evidence Processor   │
│  Engine      │ │ Engine   │ │  (Classification +    │
│              │ │          │ │   Gap Analysis)       │
│  Evaluates   │ │  0-100   │ │                       │
│  27 rules    │ │  across  │ │  Weight, present/     │
│  per playbook│ │  5 dims  │ │  missing, checklist   │
└──────┬───────┘ └────┬─────┘ └──────────┬────────────┘
       │              │                   │
       └──────────────┼───────────────────┘
                      ▼
┌─────────────────────────────────────────────────────────────────┐
│                       AI LAYER                                  │
│  ┌─────────────────────────┐   ┌────────────────────────────┐   │
│  │    Claude API Client    │   │      RAG Pipeline          │   │
│  │  (Anthropic SDK)        │   │   (Case DB + Playbook KB)  │   │
│  │                         │   │                            │   │
│  │  - Claim summaries      │   │  - Semantic case search    │   │
│  │  - Complaint drafts     │   │  - Pattern matching        │   │
│  │  - FOS submission drafts│   │  - Precedent retrieval     │   │
│  │  - Timeline generation  │   │                            │   │
│  └─────────────────────────┘   └────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────┘
                      │
┌─────────────────────────────────────────────────────────────────┐
│                       DATA LAYER                                │
│  ┌───────────────┐  ┌──────────────┐  ┌──────────────────────┐  │
│  │  Relational   │  │ Vector Store │  │   Document Store     │  │
│  │  Database     │  │ (embeddings) │  │   (PDFs, Excel,      │  │
│  │  (PostgreSQL) │  │              │  │    Markdown)         │  │
│  │               │  │  Case DB     │  │                      │  │
│  │  Claims       │  │  Playbook KB │  │  Raw case PDFs       │  │
│  │  Evidence     │  │  Semantic    │  │  Playbook source     │  │
│  │  Scores       │  │  search      │  │  documents           │  │
│  │  Playbooks    │  │              │  │                      │  │
│  └───────────────┘  └──────────────┘  └──────────────────────┘  │
└─────────────────────────────────────────────────────────────────┘
                      │
┌─────────────────────────────────────────────────────────────────┐
│                     PLAYBOOK REGISTRY                           │
│                                                                 │
│  storm-v1 (active)   flood-v1 (draft)   fire-v1 (planned)      │
│  escape-of-water-v1 (planned)   subsidence-v1 (planned)        │
│                                                                 │
│  Each playbook: rules, scoring dimensions, evidence matrix,     │
│  claim types, rejection taxonomy, flags, output config         │
└─────────────────────────────────────────────────────────────────┘
```

---

## Key Services

### 1. Playbook Registry
- Loads and validates playbook JSON configs at startup
- Exposes `get_playbook(peril_type)` to the workflow orchestrator
- Tracks playbook versions; supports parallel active versions
- Validates new playbooks against the playbook JSON schema before activation

### 2. Claim Assessment Workflow (Orchestrator)
- Stateless workflow — receives a claim input, returns a full assessment object
- Delegates to Rules Engine, Scoring Engine, Evidence Processor, Output Generator
- Halts immediately on halt-on-trigger rule (e.g. EXCLUDED_CLAIM)
- Produces a complete `ClaimAssessment` object as its output

### 3. Rules Engine
- Iterates rules in `priority` order
- Evaluates JSONLogic-style condition objects against claim state
- Applies `actions` (set flags, add checklist items, update score modifiers)
- Returns list of all `RuleTriggerResult` objects (triggered + not triggered)
- Peril-agnostic: loaded entirely from playbook config

### 4. Scoring Engine
- Computes each scoring dimension (0–max_score each) per playbook config
- Aggregates dimension scores into total
- Classifies total into strength band (Very Weak → Very Strong)
- Assigns Rejection Risk rating
- Can apply rule-triggered score modifiers (boosts / penalties)

### 5. Evidence Processor
- Accepts submitted evidence items with `present` flags
- Looks up weight from playbook evidence matrix
- Generates missing-evidence checklist for all absent high/very-high weight items
- Assigns `priority` label (critical, high, medium, low) per playbook gap config

### 6. Output Generator
- Produces structured outputs: claim_summary, timeline, rejection_classification, evidence_gap_analysis, recommended_actions, missing_evidence_checklist
- For prose outputs (complaint_draft, fos_draft): calls Claude with a structured prompt built from the assessment state + relevant case precedents from the RAG pipeline
- All outputs stored in `claim_outputs` table with model and timestamp

### 7. RAG Pipeline
- At ingestion: embeds case summaries and playbook knowledge chunks
- At query: retrieves top-k relevant cases by similarity to claim profile
- Used by Output Generator to ground complaint drafts and FOS submissions in real precedents
- Vector store: ChromaDB (local, MVP) → Pinecone or pgvector (production)

### 8. Document Ingester
- Accepts PDF, Excel, Markdown source documents
- Extracts text (pdfplumber for PDFs, openpyxl for Excel)
- Classifies content type (playbook source, raw case, claim document)
- Routes to appropriate storage and triggers embedding

---

## Technology Stack (Proposed)

| Layer | Technology | Rationale |
|---|---|---|
| Language | Python 3.11+ | AI/ML ecosystem; existing playbook work |
| API Framework | FastAPI | Async, typed, auto-docs |
| Database | PostgreSQL + SQLAlchemy | Structured claims data; JSONB for flexible rule/score storage |
| Vector Store | ChromaDB (MVP) | Local, zero-infra for MVP; swap to pgvector or Pinecone later |
| AI Model | Claude (claude-sonnet-4-6) | Prose generation, summaries, drafts |
| PDF Extraction | pdfplumber | Reliable UK court/FOS document extraction |
| Excel Extraction | openpyxl | Playbook source format |
| Testing | pytest | Unit tests for rules engine, scoring engine |
| Config | Pydantic v2 | Schema validation for playbook configs and claim inputs |

---

## Multi-Playbook Extension Path

Adding a new peril (e.g. Flood) requires:
1. Author `playbooks/flood/playbook.json` (rules, scoring, evidence matrix, flags)
2. Ingest case database PDFs for flood cases
3. Register playbook via `PlaybookRegistry.register()`
4. Zero code changes to Rules Engine, Scoring Engine, or API

The only code change needed for a new peril is if it requires a new **output type** not already in the output config (e.g. a flood-specific regulatory letter format).
