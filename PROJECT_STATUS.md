# Insurance Claims AI — Project Status

**Last updated:** 2026-06-22 (Subsidence Batch 2 complete — SUBS-011 to SUBS-020; 20 cases processed)
**Branch:** `master`
**Repository:** https://github.com/shaniv914/insurance-claims-ai

---

## Current Focus

EOW, Storm, and Flood case databases are complete. Subsidence Batch 2 complete (SUBS-001 to SUBS-020). Next: Subsidence Batch 3 (SUBS-021 to SUBS-030).

- EOW: 57 cases (EOW-001 to EOW-057) — complete
- Storm: 38 cases (STORM-001 to STORM-038) — complete
- Flood: 55 database rows (56 PDFs reviewed; FLOOD-031 excluded) — complete
- Subsidence: 20 cases processed (SUBS-001 to SUBS-020) — 12 remaining

Flood Playbook note written: `knowledge/playbooks/flood/flood-source-of-water-interpretation.md` (FLOOD-035 vs FLOOD-038 "such as" language analysis).

---

## Database: Flood_Case_Database.xlsx

**Path:** `knowledge/case-databases/Flood_Case_Database.xlsx`
**Schema version:** v1 (21 columns — same schema as EOW v2 / Storm v1)
**Total flood PDFs:** 56 (`knowledge/raw-cases/flood/`)
**Total flood cases processed:** 55 (FLOOD-031 excluded — 56 PDFs reviewed)
**Remaining to process:** 0 — all PDFs processed

### Processing Schedule (batch size = 10)

| Batch | Case IDs | PDFs | Status |
|---|---|---|---|
| 1 | FLOOD-001 – FLOOD-010 | DRN0070249 → DRN-2339494 | **Complete** |
| 2 | FLOOD-011 – FLOOD-020 | DRN2341674 → DRN-2882609 | **Complete** |
| 3 | FLOOD-021 – FLOOD-030 | DRN-2922192 → DRN-3370157 | **Complete** |
| 4 | FLOOD-031 – FLOOD-040 | DRN-3709658 → DRN-5186962 | **Complete** (9 appended; FLOOD-031 excluded) |
| 5 | FLOOD-041 – FLOOD-050 | DRN-5267073 → DRN6619758 | **Complete** |
| 6 | FLOOD-051 – FLOOD-056 | DRN7363961 → DRN9771710 | **Complete** |

### Next Batch

**None — all 56 PDFs have been processed.** The Flood case database is complete (55 database rows; FLOOD-031 excluded).
Flood playbook authoring may begin.

### Full PDF → Case ID Assignment

| Case ID | Source PDF | Batch |
|---|---|---|
| FLOOD-001 | DRN0070249.pdf | 1 |
| FLOOD-002 | DRN0420936.pdf | 1 |
| FLOOD-003 | DRN1043965.pdf | 1 |
| FLOOD-004 | DRN-1611818.pdf | 1 |
| FLOOD-005 | DRN-1846883.pdf | 1 |
| FLOOD-006 | DRN-2024904.pdf | 1 |
| FLOOD-007 | DRN-2075105.pdf | 1 |
| FLOOD-008 | DRN-2101511.pdf | 1 |
| FLOOD-009 | DRN2337599.pdf | 1 |
| FLOOD-010 | DRN-2339494.pdf | 1 |
| FLOOD-011 | DRN2341674.pdf | 2 |
| FLOOD-012 | DRN2449131.pdf | 2 |
| FLOOD-013 | DRN-2482936.pdf | 2 |
| FLOOD-014 | DRN2512341.pdf | 2 |
| FLOOD-015 | DRN-2541699.pdf | 2 |
| FLOOD-016 | DRN-2632381.pdf | 2 |
| FLOOD-017 | DRN-2785967.pdf | 2 |
| FLOOD-018 | DRN-2787998.pdf | 2 |
| FLOOD-019 | DRN-2800821.pdf | 2 |
| FLOOD-020 | DRN-2882609.pdf | 2 |
| FLOOD-021 | DRN-2922192.pdf | 3 |
| FLOOD-022 | DRN-2928961.pdf | 3 |
| FLOOD-023 | DRN2955063.pdf | 3 |
| FLOOD-024 | DRN-2965648.pdf | 3 |
| FLOOD-025 | DRN-3121807.pdf | 3 |
| FLOOD-026 | DRN-3219788.pdf | 3 |
| FLOOD-027 | DRN3290959.pdf | 3 |
| FLOOD-028 | DRN-3295916.pdf | 3 |
| FLOOD-029 | DRN3348419.pdf | 3 |
| FLOOD-030 | DRN-3370157.pdf | 3 |
| FLOOD-031 | DRN-3709658.pdf | 4 |
| FLOOD-032 | DRN-3710798.pdf | 4 |
| FLOOD-033 | DRN4280012.pdf | 4 |
| FLOOD-034 | DRN4396587.pdf | 4 |
| FLOOD-035 | DRN-4415847.pdf | 4 |
| FLOOD-036 | DRN-4895575.pdf | 4 |
| FLOOD-037 | DRN-4901901.pdf | 4 |
| FLOOD-038 | DRN-4948332.pdf | 4 |
| FLOOD-039 | DRN-5057225.pdf | 4 |
| FLOOD-040 | DRN-5186962.pdf | 4 |
| FLOOD-041 | DRN-5267073.pdf | 5 |
| FLOOD-042 | DRN-5285327.pdf | 5 |
| FLOOD-043 | DRN-5349922.pdf | 5 |
| FLOOD-044 | DRN-5387216.pdf | 5 |
| FLOOD-045 | DRN-5601561.pdf | 5 |
| FLOOD-046 | DRN5640983.pdf | 5 |
| FLOOD-047 | DRN-5827621.pdf | 5 |
| FLOOD-048 | DRN-6051732.pdf | 5 |
| FLOOD-049 | DRN6137899.pdf | 5 |
| FLOOD-050 | DRN6619758.pdf | 5 |
| FLOOD-051 | DRN7363961.pdf | 6 |
| FLOOD-052 | DRN7937294.pdf | 6 |
| FLOOD-053 | DRN7939869.pdf | 6 |
| FLOOD-054 | DRN8469463.pdf | 6 |
| FLOOD-055 | DRN9152389.pdf | 6 |
| FLOOD-056 | DRN9771710.pdf | 6 |

### Cases Processed: FLOOD-001 to FLOOD-056 (Batches 1–6 complete — 55 database rows; FLOOD-031 excluded)

| Case ID | FOS ID | Insurer | Outcome Category | Is Core Case |
|---|---|---|---|---|
| FLOOD-001 | DRN0070249 | Royal & Sun Alliance Insurance Plc | Upheld | Yes |
| FLOOD-002 | DRN0420936 | U K Insurance Limited | Upheld | Yes |
| FLOOD-003 | DRN1043965 | AXA Insurance UK Plc | Upheld | Yes |
| FLOOD-004 | DRN-1611818 | Fairmead Insurance Limited | Not Upheld | Yes |
| FLOOD-005 | DRN-1846883 | U K Insurance Limited | Not Upheld | Yes |
| FLOOD-006 | DRN-2024904 | Ageas Insurance Limited | Not Upheld | Yes |
| FLOOD-007 | DRN-2075105 | The Salvation Army General Insurance Corporation Ltd | Not Upheld | No — Administrative |
| FLOOD-008 | DRN-2101511 | Accredited Insurance (Europe) Ltd | Not Upheld | Yes |
| FLOOD-009 | DRN2337599 | UK Insurance Limited | Upheld in Part | Yes |
| FLOOD-010 | DRN-2339494 | Aviva Insurance Limited | Upheld | Yes |
| FLOOD-011 | DRN2341674 | Society of Lloyd's | Not Upheld | Yes |
| FLOOD-012 | DRN2449131 | Woodland Insurance Services Ltd | Upheld | No — Broker Dispute |
| FLOOD-013 | DRN-2482936 | AXA Insurance UK Plc | Not Upheld | Yes |
| FLOOD-014 | DRN2512341 | Zurich Insurance PLC | Not Upheld | Yes |
| FLOOD-015 | DRN-2541699 | The National Farmers' Union Mutual Insurance Society Limited | Upheld | Yes |
| FLOOD-016 | DRN-2632381 | Royal & Sun Alliance Insurance Plc | Not Upheld | Yes |
| FLOOD-017 | DRN-2785967 | Ageas Insurance Limited | Not Upheld | Yes |
| FLOOD-018 | DRN-2787998 | Ocaso SA, Compania de Seguros y Reaseguros | Not Upheld | Yes |
| FLOOD-019 | DRN-2800821 | Aviva Insurance Limited | Upheld in Part | No — Administrative |
| FLOOD-020 | DRN-2882609 | Aviva Insurance Limited | Upheld | Yes |
| FLOOD-021 | DRN-2922192 | St Andrew's Insurance Plc | Not Upheld | No — Administrative |
| FLOOD-022 | DRN-2928961 | AXA Insurance UK Plc | Not Upheld | Yes |
| FLOOD-023 | DRN2955063 | Millennium Insurance Company Limited | Upheld | Yes |
| FLOOD-024 | DRN-2965648 | Ocaso SA, Compania de Seguros y Reaseguros | Not Upheld | Yes |
| FLOOD-025 | DRN-3121807 | Ocaso SA, Compania de Seguros y Reaseguros | Not Upheld | Yes |
| FLOOD-026 | DRN-3219788 | QIC Europe Limited | Not Upheld | Yes |
| FLOOD-027 | DRN3290959 | Ageas Insurance Limited | Not Upheld | Yes |
| FLOOD-028 | DRN-3295916 | QIC Europe Ltd | Upheld | Yes |
| FLOOD-029 | DRN3348419 | Royal & Sun Alliance Insurance Plc | Upheld in Part | Yes |
| FLOOD-030 | DRN-3370157 | Fairmead Insurance Limited | Not Upheld | No — Administrative |
| FLOOD-031 | DRN-3709658 | — EXCLUDED — | — | — (home emergency contractor liability, not flood insurance) |
| FLOOD-032 | DRN-3710798 | Society of Lloyd's | Upheld in Part | No — Handling Dispute |
| FLOOD-033 | DRN4280012 | Liverpool Victoria Insurance Company Limited | Not Upheld | Yes |
| FLOOD-034 | DRN4396587 | Legal & General Insurance Limited | Upheld in Part | Yes |
| FLOOD-035 | DRN-4415847 | Wakam | Upheld | Yes |
| FLOOD-036 | DRN-4895575 | AXA Insurance UK Plc | Upheld | No — Handling Dispute |
| FLOOD-037 | DRN-4901901 | AXA Insurance UK Plc | Upheld | No — Handling Dispute |
| FLOOD-038 | DRN-4948332 | Aviva Insurance Limited | Not Upheld | Yes |
| FLOOD-039 | DRN-5057225 | Royal & Sun Alliance Insurance Limited | Upheld | No — Handling Dispute |
| FLOOD-040 | DRN-5186962 | Society of Lloyd's | Upheld | Yes |
| FLOOD-041 | DRN-5267073 | The National Farmers' Union Mutual Insurance Society Limited | Not Upheld | Yes |
| FLOOD-042 | DRN-5285327 | U K Insurance Limited | Upheld | Yes |
| FLOOD-043 | DRN-5349922 | Covea Insurance plc | Upheld in Part | No — Administrative |
| FLOOD-044 | DRN-5387216 | Covea Insurance plc | Upheld | Yes |
| FLOOD-045 | DRN-5601561 | Advantage Insurance Company Limited | Not Upheld | Yes |
| FLOOD-046 | DRN5640983 | Royal & Sun Alliance Insurance Plc | Upheld | Yes |
| FLOOD-047 | DRN-5827621 | Aviva Insurance Limited | Upheld | No — Handling Dispute |
| FLOOD-048 | DRN-6051732 | Allianz Global Corporate & Specialty SE | Compensation Only | No — Commercial |
| FLOOD-049 | DRN6137899 | UK General Insurance (Ireland) Limited | Upheld in Part | No — Commercial |
| FLOOD-050 | DRN6619758 | International Insurance Company of Hannover SE | Not Upheld | Yes |
| FLOOD-051 | DRN7363961 | Lloyds Bank General Insurance Limited | Upheld | No — Administrative |
| FLOOD-052 | DRN7937294 | UK General Insurance (Ireland) Limited | Compensation Only | Yes |
| FLOOD-053 | DRN7939869 | Royal & Sun Alliance Insurance Plc | Upheld in Part | No — Handling Dispute |
| FLOOD-054 | DRN8469463 | TBO Services Limited | Upheld | No — Broker Dispute |
| FLOOD-055 | DRN9152389 | WR Berkley Insurance (Europe) Limited | Not Upheld | Yes |
| FLOOD-056 | DRN9771710 | Society of Lloyd's | Not Upheld | Yes |

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
**Total cases processed:** 38 (STORM-001 to STORM-038)
**Total PDFs in repository:** 38 (`knowledge/raw-cases/`)
**Remaining to process:** 0 — all PDFs processed

### Cases Processed: STORM-001 to STORM-038 (complete)

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
| STORM-016 | DRN-3638410 | AA Underwriting Insurance Company Limited | Upheld | Yes |
| STORM-017 | DRN-3643634 | UK Insurance Limited | Not Upheld | No — Handling Dispute |
| STORM-018 | DRN-3819182 | Covea Insurance plc | Not Upheld | Yes |
| STORM-019 | DRN-3829618 | Royal & Sun Alliance Insurance Plc | Not Upheld | Yes |
| STORM-020 | DRN-4293834 | AXA Insurance UK Plc | Not Upheld | Yes |
| STORM-021 | DRN-4517146 | AXA XL Insurance Company UK Limited | Not Upheld | Yes |
| STORM-022 | DRN-4757581 | Accredited Insurance (Europe) Ltd | Not Upheld | Yes |
| STORM-023 | DRN-4899211 | U K Insurance Limited | Not Upheld | Yes |
| STORM-024 | DRN-5647934 | AXIS Specialty Europe SE | Not Upheld | Yes |
| STORM-025 | DRN-6075693 | AXA Insurance UK Plc | Not Upheld | Yes |
| STORM-026 | DRN0445901 | Gresham Insurance Company Limited | Not Upheld | Yes |
| STORM-027 | DRN1086734 | U K Insurance Limited | Not Upheld | Yes |
| STORM-028 | DRN1681509 | Kwik-Fit Insurance Services Ltd | Not Upheld | No — Broker Dispute |
| STORM-029 | DRN2201217 | Elite Insurance Company Limited | Not Upheld | Yes |
| STORM-030 | DRN2738252 | U K Insurance Limited | Not Upheld | Yes |
| STORM-031 | DRN3019884 | Royal & Sun Alliance Insurance Plc | Not Upheld | Yes |
| STORM-032 | DRN5013915 | Ageas Insurance Limited | Not Upheld | Yes |
| STORM-033 | DRN5397298 | U K Insurance Limited | Upheld | Yes |
| STORM-034 | DRN7021460 | Liverpool Victoria Insurance Company Limited | Not Upheld | Yes |
| STORM-035 | DRN7244667 | Liverpool Victoria Insurance Company Limited | Not Upheld | Yes |
| STORM-036 | DRN8247030 | Royal & Sun Alliance Insurance Plc | Not Upheld | Yes |
| STORM-037 | DRN8636254 | Aviva Insurance Limited | Not Upheld | Yes |
| STORM-038 | DRN8660161 | Zurich Insurance PLC | Not Upheld | Yes |

### Next Batch

**None — all 38 PDFs have been processed.** The Storm case database is complete.
Storm playbook authoring may begin.

---

## EOW Next Batch

**None — all 57 PDFs have been processed.** The EOW case database is complete.
EOW playbook authoring may begin.

---

## Database: Subsidence_Case_Database.xlsx

**Path:** `knowledge/case-databases/Subsidence_Case_Database.xlsx`
**Schema version:** v1 (21 columns — same as EOW v2 / Storm v1 / Flood v1; column 6 renamed "Movement Cause")
**Total subsidence PDFs:** 32 (`knowledge/raw-cases/subsidence/`)
**Total subsidence cases processed:** 20 (SUBS-001 to SUBS-020)
**Remaining to process:** 12

### Processing Schedule (batch size = 10)

| Batch | Case IDs | PDFs | Status |
|---|---|---|---|
| 1 | SUBS-001 – SUBS-010 | DRN0001741 → DRN-2807339 | **Complete** |
| 2 | SUBS-011 – SUBS-020 | DRN2951368 → DRN-4883553 | **Complete** |
| 3 | SUBS-021 – SUBS-030 | DRN-4950435 → DRN-6019596 | Pending |
| 4 | SUBS-031 – SUBS-032 | DRN8130715 → DRN8561608 | Pending |

### Full PDF List (32 files — sorted by DRN)

| # | Source PDF |
|---|---|
| 1 | DRN0001741.pdf |
| 2 | DRN0017653.pdf |
| 3 | DRN0618226.pdf |
| 4 | DRN1210158.pdf |
| 5 | DRN1933952.pdf |
| 6 | DRN2093738.pdf |
| 7 | DRN-2213774.pdf |
| 8 | DRN2337317.pdf |
| 9 | DRN2707923.pdf |
| 10 | DRN-2807339.pdf |
| 11 | DRN2951368.pdf |
| 12 | DRN-3258437.pdf |
| 13 | DRN-3387540.pdf |
| 14 | DRN-3427348.pdf |
| 15 | DRN-3581769.pdf |
| 16 | DRN-3682901.pdf |
| 17 | DRN-3929594.pdf |
| 18 | DRN-4190935.pdf |
| 19 | DRN-4813489.pdf |
| 20 | DRN-4883553.pdf |
| 21 | DRN-4950435.pdf |
| 22 | DRN5217766.pdf |
| 23 | DRN-5220010.pdf |
| 24 | DRN-5315100.pdf |
| 25 | DRN-5375880.pdf |
| 26 | DRN-5643066.pdf |
| 27 | DRN-5656370.pdf |
| 28 | DRN-5718419.pdf |
| 29 | DRN-5755602.pdf |
| 30 | DRN-6019596.pdf |
| 31 | DRN8130715.pdf |
| 32 | DRN8561608.pdf |

### Cases Processed: SUBS-001 to SUBS-020 (Batches 1–2 complete)

| Case ID | FOS ID | Insurer | Outcome Category | Is Core Case |
|---|---|---|---|---|
| SUBS-001 | DRN0001741 | UK Insurance Limited | Upheld | Yes |
| SUBS-002 | DRN0017653 | Aviva Insurance Limited | Upheld | No — Handling Dispute |
| SUBS-003 | DRN0618226 | AXA Insurance UK Plc | Upheld | Yes |
| SUBS-004 | DRN1210158 | Royal & Sun Alliance Insurance Plc | Upheld | Yes |
| SUBS-005 | DRN1933952 | Royal & Sun Alliance Insurance Plc | Upheld | Yes |
| SUBS-006 | DRN2093738 | Society of Lloyd's | Upheld | Yes |
| SUBS-007 | DRN-2213774 | Advantage Insurance Company Limited | Upheld | Yes |
| SUBS-008 | DRN2337317 | Royal & Sun Alliance Insurance Plc | Upheld | Yes |
| SUBS-009 | DRN2707923 | AXA Insurance UK Plc | Upheld | Yes |
| SUBS-010 | DRN-2807339 | Aviva Insurance Limited | Upheld | Yes |
| SUBS-011 | DRN2951368 | Covea Insurance Plc | Upheld | Yes |
| SUBS-012 | DRN-3258437 | Red Sands Insurance Company (Europe) Limited | Upheld | Yes |
| SUBS-013 | DRN-3387540 | Accredited Insurance (Europe) Limited | Upheld | Yes |
| SUBS-014 | DRN-3427348 | Amtrust Europe Limited | Upheld in Part | Yes |
| SUBS-015 | DRN-3581769 | Fairmead Insurance Limited | Upheld in Part | Yes |
| SUBS-016 | DRN-3682901 | Society of Lloyd's | Upheld | Yes |
| SUBS-017 | DRN-3929594 | AXA Insurance UK Plc | Upheld | Yes |
| SUBS-018 | DRN-4190935 | Society of Lloyd's | Upheld | Yes |
| SUBS-019 | DRN-4813489 | Kennett Insurance Brokers Limited | Upheld | No — Broker Dispute |
| SUBS-020 | DRN-4883553 | AXA Insurance UK Plc | Upheld | Yes |

### Next Batch

**Batch 3 — SUBS-021 to SUBS-030**

PDFs: DRN-4950435.pdf → DRN-6019596.pdf (PDFs #21–30 from the full list below).

---

## Scripts

| Script | Status | Purpose |
|---|---|---|
| `scripts/append_subsidence_v1.py` | **Active — use for Subsidence Batch 2+** | Batch 1 append script (SUBS-001 to SUBS-010); reuse for future batches by updating NEW_CASES |
| `scripts/create_subsidence_case_db.py` | Superseded — do not re-run | Created the empty Subsidence_Case_Database.xlsx header row; re-running will overwrite data |
| `scripts/append_eow_v2.py` | **Active** | Standard append for all future EOW batches (schema v2, 21 columns, controlled-vocab validation) |
| `scripts/append_storm_v1.py` | **Active — use this for Storm** | Standard append for all Storm batches (schema v1, 21 columns, controlled-vocab validation) |
| `scripts/append_flood_v6.py` | Final batch complete — do not reuse | Batch 6 append script (FLOOD-051 to FLOOD-056); Flood database complete |
| `scripts/append_flood_v5.py` | Superseded by v6 — do not reuse | Batch 5 append script (FLOOD-041 to FLOOD-050); historical record only |
| `scripts/append_flood_v4.py` | Superseded by v5 — do not reuse | Batch 4 append script (FLOOD-032 to FLOOD-040); historical record only |
| `scripts/append_flood_v3.py` | Superseded by v4 — do not reuse | Batch 3 append script (FLOOD-021 to FLOOD-030); historical record only |
| `scripts/append_flood_v2.py` | Superseded by v3 — do not reuse | Batch 2 append script (FLOOD-011 to FLOOD-020); historical record only |
| `scripts/append_flood_v1.py` | Superseded by v2 — do not reuse | Batch 1 append script (FLOOD-001 to FLOOD-010); historical record only |
| `scripts/create_flood_case_db.py` | Superseded — do not re-run | Created the empty Flood_Case_Database.xlsx header row; re-running will overwrite data |
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
| *(latest)* | Add SUBS-011 to SUBS-020 to Subsidence Case Database — Batch 2 complete |
| `790486e` | Add SUBS-001 to SUBS-010 to Subsidence Case Database — Batch 1 complete |
| `a0c11c0` | Update PROJECT_STATUS.md — Subsidence peril added, PDF library committed |
| `e2c18a5` | Add Subsidence raw PDF library — 32 PDFs (knowledge/raw-cases/subsidence/) |
| `2902a59` | Add FLOOD-051 to FLOOD-056 to Flood Case Database — Batch 6 complete (Flood database final) |
| `aa506db` | Add FLOOD-041 to FLOOD-050 to Flood Case Database - Batch 5 complete |
| `f7d5ba0` | Add FLOOD-032 to FLOOD-040 to Flood Case Database - Batch 4 complete |
| `b7537ff` | Add FLOOD-021 to FLOOD-030 to Flood Case Database - Batch 3 complete |
| `0772c95` | Add FLOOD-011 to FLOOD-020 to Flood Case Database - Batch 2 complete |
| `56277c5` | Add FLOOD-001 to FLOOD-010 to Flood Case Database - Batch 1 complete |
| `1c9acfc` | Add STORM-036 to STORM-038 to Storm Case Database — complete |
| `6b8340b` | Add STORM-031 to STORM-035 to Storm Case Database |
| `62934d4` | Add STORM-026 to STORM-030 to Storm Case Database |
| `ef9a197` | Add STORM-021 to STORM-025 to Storm Case Database |
| `420499f` | Add STORM-016 to STORM-020 to Storm Case Database |
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
