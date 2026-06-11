# Insurance Claims AI — Project Status

**Last updated:** 2026-06-11 (STORM-011 to STORM-015 appended)
**Branch:** `master`
**Repository:** https://github.com/shaniv914/insurance-claims-ai

---

## Current Focus

All 57 Escape of Water (EOW) FOS case PDFs have been processed and classified.
The EOW case database is complete. EOW playbook authoring may now begin.

---

## Database: Escape_of_Water_Case_Database.xlsx

**Path:** `knowledge/case-databases/Escape_of_Water_Case_Database.xlsx`
**Schema version:** v2 (21 columns)
**Total cases processed:** 57 (EOW-001 to EOW-057)
**Total PDFs in repository:** 57 (`knowledge/raw-cases/escape-of-water/`)
**Remaining to process:** 0 — all PDFs processed

### Schema — Column Order (v2)

| # | Column | Type | Notes |
|---|---|---|---|
| 1 | Case ID | Text | Format: EOW-NNN |
| 2 | FOS Decision ID | Text | Matches PDF header (DRN-XXXXXXX) |
| 3 | Insurer Name | Text | Formal registered name from FOS decision |
| 4 | FOS Decision Date | Text | Accept-or-reject deadline as printed (DD Mon YYYY) |
| 5 | Claim Type | Free text | Physical incident and dispute in one sentence |
| 6 | Leak Source | Free text | Physical origin of the water |
| 7 | Property Type | Free text | Residential / Unoccupied / Leasehold / Commercial |
| 8 | Dispute Type | Controlled vocab | 7 values — see below |
| 9 | Coverage Decision | Controlled vocab | 5 values — see below |
| 10 | Rejection Reason | Free text | Insurer's stated reason for declining |
| 11 | Evidence Dispute | Free text | What evidence each party relied on |
| 12 | Outcome Category | Controlled vocab | 4 values — see below |
| 13 | Outcome | Free text | Full FOS remedy instructions |
| 14 | Compensation Awarded (£) | Integer | Distress/inconvenience only; 0 if none |
| 15 | Is Core Case | Controlled vocab | 5 values — see below |
| 16 | Key Policy Clause | Free text | Policy wording or FOS/FCA principle applied |
| 17 | Missing Evidence | Free text | Evidence absent that affected the outcome |
| 18 | Ombudsman Reasoning | Free text | How the ombudsman weighed the evidence |
| 19 | Workflow Insight | Free text | Operational rule for the claims workflow |
| 20 | AI Rule Candidate | Free text | Machine-evaluable rule for the rules engine |
| 21 | Source PDF | Text | Filename only (e.g. DRN-1135035.pdf) |

### Controlled Vocabulary

**Dispute Type**
- `Coverage Dispute`
- `Handling / Reinstatement Dispute`
- `Endorsement / Exclusion Challenge`
- `Pre-Inception Damage Dispute`
- `Peril Classification Dispute`
- `Claim Recording / Administrative Dispute`
- `Broker Conduct Dispute`

**Coverage Decision**
- `Declined — Full`
- `Declined — Partial`
- `Accepted`
- `Accepted — Disputed Settlement`
- `Not Applicable`

**Outcome Category**
- `Upheld`
- `Upheld in Part`
- `Not Upheld`
- `Compensation Only` — coverage decline upheld; compensation for handling failure only

**Is Core Case**
- `Yes`
- `No — Administrative`
- `No — Handling Dispute`
- `No — Commercial`
- `No — Broker Dispute`

---

## Cases Processed: EOW-001 to EOW-057 (complete)

| Case ID | FOS ID | Insurer | Outcome Category | Is Core Case |
|---|---|---|---|---|
| EOW-001 | DRN-2088302 | Fairmead Insurance Limited | Upheld in Part | Yes |
| EOW-002 | DRN-3258984 | U K Insurance Limited (UKI) | Upheld | Yes |
| EOW-003 | DRN-4132443 | QIC Europe Ltd | Upheld | Yes |
| EOW-004 | DRN4969124 | UK Insurance Limited (UKI) | Upheld | No — Administrative |
| EOW-005 | DRN-4334761 | Royal & Sun Alliance Insurance Ltd (RSA) | Upheld | Yes |
| EOW-006 | DRN-1135035 | Covea Insurance plc | Not Upheld | Yes |
| EOW-007 | DRN-1525346 | AXA Insurance UK Plc | Not Upheld | Yes |
| EOW-008 | DRN-1623377 | Insurers at Lloyd's (Society of Lloyd's) | Not Upheld | Yes |
| EOW-009 | DRN-2340952 | U K Insurance Limited (UKI) | Not Upheld | No — Handling Dispute |
| EOW-010 | DRN2512341 | Zurich Insurance PLC | Not Upheld | Yes |
| EOW-011 | DRN-2540997 | Fairmead Insurance Limited | Upheld | Yes |
| EOW-012 | DRN-2572816 | Admiral Insurance (Gibraltar) Limited | Upheld | Yes |
| EOW-013 | DRN-2749125 | Covea Insurance Plc | Compensation Only | Yes |
| EOW-014 | DRN-2771304 | AXA Insurance UK Plc | Upheld | No — Commercial |
| EOW-015 | DRN-2806638 | Endsleigh Insurance Services Ltd | Not Upheld | No — Broker Dispute |
| EOW-016 | DRN-3022853 | Admiral Insurance (Gibraltar) Limited | Not Upheld | Yes |
| EOW-017 | DRN-3053156 | Aviva Insurance Limited | Upheld | Yes |
| EOW-018 | DRN-3078337 | QIC Europe Ltd | Upheld | Yes |
| EOW-019 | DRN-3121008 | AXA Insurance UK Plc | Upheld | Yes |
| EOW-020 | DRN-3517894 | AA Underwriting Insurance Company Limited | Upheld | Yes |
| EOW-021 | DRN-3606995 | Ageas Insurance Limited | Not Upheld | Yes |
| EOW-022 | DRN-3860121 | Tesco Underwriting Limited | Upheld | Yes |
| EOW-023 | DRN-4205492 | QIC Europe Limited | Upheld | Yes |
| EOW-024 | DRN-4223988 | HDI Global Speciality SE | Upheld | Yes |
| EOW-025 | DRN-4227214 | Saga Services Limited | Not Upheld | Yes |
| EOW-026 | DRN3376494 | Covea Insurance plc | Not Upheld | Yes |
| EOW-027 | DRN3405029 | Aviva Insurance Limited | Upheld | Yes |
| EOW-028 | DRN4208888 | Millennium Insurance Company Limited | Upheld | Yes |
| EOW-029 | DRN-4307523 | QIC Europe Ltd | Upheld | Yes |
| EOW-030 | DRN-4368751 | AXA Insurance Limited | Not Upheld | Yes |
| EOW-031 | DRN4464315 | Fairmead Insurance Limited | Not Upheld | Yes |
| EOW-032 | DRN-4521660 | esure Insurance Limited | Not Upheld | Yes |
| EOW-033 | DRN-4704763 | Covea Insurance plc | Not Upheld | No — Commercial |
| EOW-034 | DRN-4744346 | AXIS Specialty Europe SE | Not Upheld | No — Commercial |
| EOW-035 | DRN-4749282 | Aviva Insurance Limited | Not Upheld | Yes |
| EOW-036 | DRN-5088221 | Accelerant Insurance Europe SA/NV UK Branch | Upheld in Part | No — Commercial |
| EOW-037 | DRN-5193042 | Allied World Assurance Company (Europe) dac | Not Upheld | Yes |
| EOW-038 | DRN-5198749 | Accredited Insurance (Europe) Ltd | Upheld | Yes |
| EOW-039 | DRN-5199107 | Admiral Insurance (Gibraltar) Limited | Not Upheld | No — Handling Dispute |
| EOW-040 | DRN-5396824 | Aviva Insurance Limited | Upheld in Part | Yes |
| EOW-041 | DRN5611706 | CIS General Insurance Limited | Upheld | Yes |
| EOW-042 | DRN-5649220 | Aviva Insurance Limited | Not Upheld | Yes |
| EOW-043 | DRN5670903 | Insurers at Lloyd's (Society of Lloyd's) | Not Upheld | Yes |
| EOW-044 | DRN-5805040 | Saga Services Limited | Not Upheld | Yes |
| EOW-045 | DRN5927839 | Admiral Insurance (Gibraltar) Limited | Not Upheld | Yes |
| EOW-046 | DRN-5979060 | Ecclesiastical Insurance Office Plc | Not Upheld | Yes |
| EOW-047 | DRN-5982848 | Protector Insurance UK | Not Upheld | No — Commercial |
| EOW-048 | DRN-6004362 | Aviva Insurance Limited | Upheld in Part | No — Handling Dispute |
| EOW-049 | DRN-6046762 | INTACT INSURANCE UK LIMITED | Not Upheld | No — Handling Dispute |
| EOW-050 | DRN6187392 | St Andrew's Insurance Plc | Upheld in Part | Yes |
| EOW-051 | DRN-6263959 | Tesco Underwriting Limited | Upheld | Yes |
| EOW-052 | DRN6739737 | Royal & Sun Alliance Insurance Plc | Not Upheld | Yes |
| EOW-053 | DRN7090448 | UK Insurance Limited | Upheld | Yes |
| EOW-054 | DRN7112734 | Hiscox Insurance Company Limited | Not Upheld | No — Administrative |
| EOW-055 | DRN7147115 | Legal & General Insurance Limited | Upheld | Yes |
| EOW-056 | DRN7411842 | Ageas Insurance Limited | Not Upheld | Yes |
| EOW-057 | DRN9891691 | UK Insurance Limited | Not Upheld | Yes |

---

## Database: Storm_Case_Database.xlsx

**Path:** `knowledge/case-databases/Storm_Case_Database.xlsx`
**Schema version:** v1 (21 columns — same schema as EOW v2)
**Total cases processed:** 15 (STORM-001 to STORM-015)
**Total PDFs in repository:** 38 (`knowledge/raw-cases/`)
**Remaining to process:** 23 (STORM-016 to STORM-038)

### Cases Processed: STORM-001 to STORM-015

| Case ID | FOS ID | Insurer | Outcome Category | Is Core Case |
|---|---|---|---|---|
| STORM-001 | DRN-1207086 | Aviva Insurance Limited | Not Upheld | Yes |
| STORM-002 | DRN-1223113 | Ageas Insurance Limited | Upheld | Yes |
| STORM-003 | DRN-1586732 | esure Insurance Limited | Not Upheld | Yes |
| STORM-004 | DRN-2053943 | Lloyds Bank General Insurance Limited | Not Upheld | Yes |
| STORM-005 | DRN-2556262 | Lloyds Bank General Insurance Limited | Not Upheld | Yes |
| STORM-006 | DRN-2560515 | Royal & Sun Alliance Insurance Plc | Not Upheld | Yes |
| STORM-007 | DRN-2737383 | Liverpool Victoria Insurance Company Limited | Upheld | Yes |
| STORM-008 | DRN-2788212 | Lloyds Bank General Insurance Limited | Not Upheld | Yes |
| STORM-009 | DRN-2877529 | Lloyds Bank General Insurance Limited | Not Upheld | Yes |
| STORM-010 | DRN-2926734 | Fairmead Insurance Limited | Upheld | Yes |
| STORM-011 | DRN-2926772 | National Farmers' Union Mutual Insurance Society Limited | Not Upheld | Yes |
| STORM-012 | DRN-3173328 | Aviva Insurance Limited | Not Upheld | Yes |
| STORM-013 | DRN-3211590 | Aviva Insurance Limited | Not Upheld | Yes |
| STORM-014 | DRN-3295758 | Aviva Insurance Limited | Not Upheld | Yes |
| STORM-015 | DRN-3574617 | Lloyds Bank General Insurance Limited | Not Upheld | Yes |

### Next Batch

**Batch 3 — STORM-016 to STORM-020:**

| STORM-ID | PDF |
|---|---|
| STORM-016 | DRN-3638410.pdf |
| STORM-017 | DRN-3643634.pdf |
| STORM-018 | DRN-3819182.pdf |
| STORM-019 | DRN-3829618.pdf |
| STORM-020 | DRN-4293834.pdf |

---

## EOW Next Batch

**None — all 57 PDFs have been processed.** The EOW case database is complete.
EOW playbook authoring may begin.

---

## Scripts

| Script | Status | Purpose |
|---|---|---|
| `scripts/append_eow_v2.py` | **Active** | Standard append for all future EOW batches (schema v2, 21 columns, controlled-vocab validation) |
| `scripts/append_storm_v1.py` | **Active — use this for Storm** | Standard append for all Storm batches (schema v1, 21 columns, controlled-vocab validation) |
| `scripts/create_storm_case_db.py` | Superseded — do not re-run | Created the empty Storm_Case_Database.xlsx header row; re-running will overwrite data |
| `scripts/migrate_schema_v2.py` | Reference only | One-time migration that added 7 columns and backfilled EOW-001–015; do not re-run |
| `scripts/append_eow_006_015.py` | Superseded | 14-column schema; historical record of first batch only |
| `scripts/create_eow_case_db.py` | Superseded | Original 14-column seeder; do not re-run (will overwrite with old schema) |

---

## Non-Core Case Handling Rules

Cases flagged `Is Core Case = No` are retained in the database as reference records
but are **excluded from core playbook rule derivation**. Rules:

| Flag | Definition | Cases so far |
|---|---|---|
| `No — Administrative` | Claim recording or admin dispute; no coverage analysis performed | EOW-004 |
| `No — Handling Dispute` | Claim was accepted; entire FOS dispute concerns reinstatement quality or settlement quantum; no coverage principle established | EOW-009 |
| `No — Commercial` | Commercial or all-risks policy; peril classification or excess rules may not apply to standard residential policies | EOW-014 |
| `No — Broker Dispute` | Broker conduct / renewal disclosure complaint; no claim assessment principle; no coverage decision made | EOW-015 |

When processing future PDFs, apply `Is Core Case = No` if **any** of the following are true:
- The FOS decision concerns only how the claim was administered or recorded
- The claim was fully accepted and the only dispute is about reinstatement quality or cash settlement amount
- The policy is commercial, all-risks, or a property owners policy for a managed block
- The complaint names a broker and concerns advice, disclosure, or renewal documents rather than the insurer's claim decision

---

## Recent Commits

| Hash | Message |
|---|---|
| `3ca74df` | Add STORM-011 to STORM-015 to Storm Case Database |
| `fdb38e6` | Add STORM-006 to STORM-010 to Storm Case Database |
| `2492d77` | Add STORM-001 to STORM-005 to Storm Case Database |
| `03333d8` | Add EOW-036 to EOW-057 to Escape of Water Case Database |
| `2161a62` | Add EOW-026 to EOW-035 to Escape of Water Case Database |
| `e8bd4f7` | Add EOW-016 through EOW-025 to Escape of Water Case Database |
| `19e90a8` | Add schema v2 append script for Escape of Water cases |
| `5bd8477` | Add append_eow_v2.py — standard append script for schema v2 (21 columns) |
| `0d350d5` | Schema v2: add 7 new columns and backfill EOW-001 to EOW-015 |

---

## Overall Project Roadmap

See `docs/roadmap/mvp-roadmap.md` for the full 6-phase plan.

Current phase: **Knowledge base construction** (pre-Phase 1).
EOW case database must reach sufficient coverage before EOW playbook
authoring begins in Phase 6 (Multi-Playbook Abstraction).

Target case count before playbook authoring: **all 57 PDFs processed**.
