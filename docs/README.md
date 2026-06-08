# Insurance Claims AI — Design Documentation

Architecture and design for a multi-playbook, peril-agnostic AI-assisted insurance claims assessment platform. UK jurisdiction / FOS framework. Storm Damage (v1.0) is the first playbook.

## Documents

### Architecture
| Document | Contents |
|---|---|
| [01 — System Architecture](architecture/01-system-architecture.md) | Component map, service descriptions, technology stack, multi-playbook extension path |
| [02 — Database Schema](architecture/02-database-schema.md) | Full PostgreSQL schema: all tables, columns, indexes |
| [03 — Folder Structure](architecture/03-folder-structure.md) | Full directory tree with file-level descriptions |
| [04 — Workflow Engine](architecture/04-workflow-engine.md) | 9-stage workflow, rule condition format, score modifiers, halt behaviour |
| [05 — Playbook Ingestion](architecture/05-playbook-ingestion.md) | 7-stage ingestion pipeline, rule extraction, embedding, CLI usage |

### Schemas
| Schema | Contents |
|---|---|
| [playbook.schema.json](schemas/playbook.schema.json) | Full playbook configuration — rules, scoring, evidence matrix, flags, outputs |
| [rule.schema.json](schemas/rule.schema.json) | Single rule — JSONLogic conditions, flags, score modifiers, checklist items |
| [evidence-item.schema.json](schemas/evidence-item.schema.json) | Evidence item with weight, presence, source, and conflict tracking |
| [claim-assessment.schema.json](schemas/claim-assessment.schema.json) | Full assessment output — scores, flags, missing evidence, all output types |

### Roadmap
| Document | Contents |
|---|---|
| [MVP Roadmap](roadmap/mvp-roadmap.md) | 6 phases, definition of done per phase, risk register, key metrics |

## Source Knowledge
- `knowledge/playbooks/Storm_Damage_Playbook_v1.0.md` — authoritative source for Storm playbook
- `knowledge/raw-cases/DRN-*.pdf` — 35 raw FOS/Ombudsman decision PDFs
- `knowledge/case-databases/Storm case database.xlsx` — 38-case structured summary
