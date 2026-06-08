# Storm Damage Playbook v1
## UK Home Insurance Storm Claims Assessment — Knowledge Base & Rules Engine

**Version:** 1.0  
**Derived from:** Storm Case Database (38 cases) + Storm Damage Playbook framework  
**Purpose:** Knowledge base and deterministic rules engine for AI-assisted storm claims assessment  
**Jurisdiction:** United Kingdom (FOS / Ombudsman framework)

---

## Table of Contents

1. [Storm Damage Definition](#section-1--storm-damage-definition)
2. [Supported Claim Types](#section-2--supported-claim-types)
3. [Rejection Taxonomy](#section-3--rejection-taxonomy)
4. [Evidence Matrix](#section-4--evidence-matrix)
5. [Claim Strength Scoring System](#section-5--claim-strength-scoring-system)
6. [Missing Evidence Engine](#section-6--missing-evidence-engine)
7. [Emergency Evidence Preservation Checklist](#section-7--emergency-evidence-preservation-checklist)
8. [Ombudsman Reasoning Patterns](#section-8--ombudsman-reasoning-patterns)
9. [AI Workflow Rules — Rules Engine](#section-9--ai-workflow-rules--rules-engine)
10. [Required Inputs](#section-10--required-inputs)
11. [Required Outputs](#section-11--required-outputs)
12. [JSON Schema](#section-12--json-schema)
13. [Case Database Reference](#section-13--case-database-reference-38-cases)

---

## Section 1 — Storm Damage Definition

### The Ombudsman Three-Question Framework

All storm claims must be assessed against all three of the following questions. This is the primary decision framework used by the Financial Ombudsman Service (FOS) in the UK.

| Question | Test | If NO |
|---|---|---|
| **Q1** | Did storm conditions occur at or near the property? | Claim fails immediately |
| **Q2** | Is the damage consistent with storm damage? | Claim fails — high rejection risk |
| **Q3** | Was the storm the dominant and effective cause of the damage? | Claim fails — most common failure point |

**If any answer is NO:**
- Claim Strength = **Weak**
- Rejection Risk = **High**

### What Counts as Storm Conditions?

There is no single universal wind-speed threshold in UK insurance law. The assessment is contextual:

- Some policies define a threshold (commonly 55 mph / Beaufort Scale Force 10)
- Where no definition exists, the Ombudsman applies a reasonableness test
- Local geography matters — an exposed coastal or elevated location may experience storm conditions at lower recorded wind speeds
- Weather data must be validated against the property postcode and nearby weather stations
- Internal insurer storm definitions are **not automatically decisive** if conditions were reasonably storm-like (see STORM-011)

### What Counts as Damage Consistent with Storm Damage?

**Storm-consistent damage (positive indicators):**
- Missing, displaced, or lifted roof tiles
- Collapsed boundary or garden walls
- Displaced flashing or ridge tiles
- Damaged or collapsed chimney stacks
- Storm-blown debris causing structural impact
- Structural movement caused by wind force

**Not storm-consistent damage (negative indicators):**
- Gradual staining or internal damp
- Old leak tracks or historic water marks
- Deteriorated mortar or aged pointing
- Sagging or worn felt / battens
- Cracked or failed sealant / flashings from age
- Leaks that developed over weeks or months

### The Dominant Cause Test

The storm must be the **dominant and effective cause** of the damage. This is the most frequently litigated element:

- A storm that **reveals** an existing defect does **not** cause the damage
- A storm that **triggers** pre-existing deterioration does **not** cause the damage
- A storm that **exacerbates** gradual wear and tear does **not** cause the damage
- The customer must show the storm was the **primary driver**, not merely a contributing factor

> **Most Important Finding (38 cases):** The biggest predictor of claim success is not the severity of the storm. It is whether the customer can prove the storm physically caused the damage, rather than merely revealing an existing defect.

---

## Section 2 — Supported Claim Types

The following damage types are within scope of storm claims assessment:

| Damage Type | Notes |
|---|---|
| Roof tile damage | Most common claim type |
| Roof leaks / water ingress | Must prove storm as source of ingress |
| Chimney stack damage | Damp alone does not prove storm causation |
| Flashing / ridge tile damage | Often displaced by wind |
| Felt / batten damage | Often wear and tear — scrutinise carefully |
| Guttering damage | Distinguish storm damage from blocked/overflowing gutters |
| Boundary wall collapse | Require structural engineer report and maintenance history |
| Garden wall collapse | Check for vegetation damage, age, construction quality |
| Structural storm damage | Any structural damage caused by direct wind force or debris |
| Porch / flat roof damage | Flat roofs have higher deterioration risk — scrutinise carefully |
| Bay window structure | High risk of pre-existing structural weakness claim |
| Caravan roof damage | Design fault / failed seals frequently defeats these claims |

**Out of Scope / Automatic Exclusion Check Required:**
- Boundary fences (frequently excluded by policy — **always check policy wording first**)
- Gradual / progressive damage presented as storm damage
- Pure subsidence or heave

---

## Section 3 — Rejection Taxonomy

### Rejection Categories — Ranked by Frequency (38 cases)

```json
{
  "wear_and_tear_deterioration": 15,
  "failure_to_prove_dominant_cause": 13,
  "gradual_deterioration": 11,
  "storm_revealed_existing_defect": 8,
  "no_storm_conditions_proven": 7,
  "no_physical_storm_damage_visible": 6,
  "poor_workmanship": 6,
  "design_defect_defective_construction": 4,
  "lack_of_evidence_evidence_preservation_failure": 2,
  "specific_policy_exclusion": 1,
  "failure_to_prove_source_of_ingress": 1,
  "failure_to_prove_latest_storm_caused_damage": 1
}
```

### Key Insight

> **Actual absence of a storm is less common than assumed.** Most claims fail not because no storm occurred, but because the roof already deteriorated, maintenance issues existed, design defects existed, or the storm merely exposed an underlying weakness.

### Most Common Reasons Claims Fail (ranked)

1. Wear and tear / deterioration already present
2. Existing deterioration already present
3. Storm revealed a pre-existing defect rather than causing damage
4. No convincing evidence storm was the dominant cause
5. No visible storm-related physical damage

### Winning Factors (policyholder)

- Property was shown to be well-maintained before the storm
- Contractor or expert evidence directly linked damage to the storm
- Insurer relied on assumptions rather than actual evidence
- Insurer could not prove an exclusion applied
- Weather data confirmed qualifying conditions existed at the property
- Photos supported the claimed damage type and causation

### Losing Factors (policyholder)

- Existing deterioration already present at time of storm
- Design or workmanship defects found
- No visible storm-related physical damage
- Customer unable to prove storm was dominant cause
- Storm merely exposed an underlying weakness
- Evidence was destroyed or not preserved before inspection

---

## Section 4 — Evidence Matrix

### Evidence Weight Classification

| Evidence Type | Weight | Notes |
|---|---|---|
| Pre-storm inspection report | Very High | Baseline condition established — highly persuasive |
| Structural engineer report | Very High | Carries enormous weight regardless of which party commissions it |
| Surveyor report | Very High | Insurer surveyor reports are highly persuasive but can be challenged with photos |
| Drone inspection | Very High | Objective aerial evidence — increasingly common and decisive |
| Photos (immediate post-storm) | High | Contemporaneous photos outweigh later assertions |
| Maintenance records | High | Demonstrates property was well maintained |
| Contractor report | Medium | Useful but not sufficient alone — must be supported by other evidence |
| Roofer opinion | Medium | Alone, does not outweigh surveyor or photographic evidence |
| Weather data | Medium | Necessary but not sufficient — location context matters |
| Repair invoices | Medium | Demonstrates prior repairs and maintenance |
| Customer statement | Low | Insufficient alone — must be corroborated |

### Evidence Hierarchy Rules

- Multiple independent experts agreeing > single expert opinion
- Pre-loss inspection records > post-loss assertion about pre-loss condition
- Contemporaneous photographs > retrospective customer description
- Insurer must prove exclusions apply — the burden is on the insurer, not the policyholder
- A surveyor report can be undermined by photographic evidence (see STORM-035)
- A roofer opinion alone rarely overturns a surveyor's report (see STORM-019)

---

## Section 5 — Claim Strength Scoring System

Claims are scored across five dimensions, each worth 0–20 points, giving a total of 0–100.

### Scoring Dimensions

#### Dimension 1 — Storm Verification (0–20)

| Score | Criteria |
|---|---|
| 18–20 | Weather data confirms qualifying conditions; corroborated by multiple sources |
| 13–17 | Weather data supports storm but is from a distant station or is borderline |
| 8–12 | Storm warning existed but property-specific data is absent |
| 3–7 | No weather data; customer assertion only |
| 0–2 | Weather data contradicts storm conditions |

#### Dimension 2 — Physical Damage Evidence (0–20)

| Score | Criteria |
|---|---|
| 18–20 | Clear physical storm indicators (missing tiles, collapsed wall, displaced flashing) with photographs |
| 13–17 | Physical damage present but photographs incomplete |
| 8–12 | Internal damage (leak, ceiling staining) without clear external storm indicators |
| 3–7 | Customer description only; no photographs or physical evidence |
| 0–2 | No visible storm-related physical damage at all |

#### Dimension 3 — Expert Evidence (0–20)

| Score | Criteria |
|---|---|
| 18–20 | Structural engineer or surveyor directly attributes damage to storm |
| 13–17 | Contractor report attributes damage to storm; no contradicting expert |
| 8–12 | Contractor report supportive but contradicted by insurer expert |
| 3–7 | No expert evidence; customer assertion only |
| 0–2 | Multiple independent experts attribute damage to deterioration or defect |

#### Dimension 4 — Maintenance History (0–20)

| Score | Criteria |
|---|---|
| 18–20 | Documented maintenance records show property in good condition pre-storm |
| 13–17 | Contractor confirms good pre-storm condition, no records |
| 8–12 | No maintenance records; no evidence of neglect either |
| 3–7 | Evidence of deferred maintenance or previous unrepaired damage |
| 0–2 | Pre-storm inspection or photos show deterioration existed |

#### Dimension 5 — Causation Strength (0–20)

| Score | Criteria |
|---|---|
| 18–20 | Expert directly links storm as dominant cause; no competing explanation |
| 13–17 | Storm damage is probable dominant cause; minor competing factors |
| 8–12 | Causation contested; competing explanations credible |
| 3–7 | Storm likely revealed rather than caused damage |
| 0–2 | Clear evidence damage predates storm or is from another cause |

### Total Score Interpretation

| Score Range | Classification | Action |
|---|---|---|
| 81–100 | Very Strong | Likely valid claim; support with available evidence |
| 61–80 | Strong | Viable claim; address any evidence gaps |
| 41–60 | Moderate | Contested claim; evidence gaps materially affect outcome |
| 21–40 | Weak | High rejection risk; significant evidence required |
| 0–20 | Very Weak | Claim is unlikely to succeed without major new evidence |

---

## Section 6 — Missing Evidence Engine

For every claim, run the following evidence gap check. Generate a missing-evidence checklist automatically based on what is absent.

### Evidence Gap Check

| Evidence Item | Present? | If Missing — Risk Level |
|---|---|---|
| Weather data for property postcode on date of loss | ☐ | High — storm verification fails |
| Photographs of damage taken before repairs | ☐ | Very High — claim strength severely impaired |
| Photographs of external storm indicators (tiles, walls) | ☐ | High — causation hard to prove |
| Contractor / roofer report | ☐ | Medium — useful corroboration missing |
| Surveyor or structural engineer report | ☐ | High — expert causation evidence missing |
| Maintenance records | ☐ | Medium–High — deterioration defence harder to rebut |
| Repair invoices (prior repairs) | ☐ | Medium — prior repair history unknown |
| Policy wording | ☐ | Critical — exclusions cannot be checked |
| Rejection letter from insurer | ☐ | Critical — rejection basis unknown |
| Drone inspection report | ☐ | Medium — high-value evidence if obtainable |

### Automatic Checklist Generation Logic

```
IF weather_data IS missing:
  → Add to checklist: "Obtain Met Office / Weather Assured / equivalent data for [postcode] on [date]"

IF photos IS missing OR photos taken after repairs:
  → Add to checklist: "CRITICAL: No pre-repair photographs. Obtain any available photos from customer, neighbours, or contractors."
  → Flag: Evidence preservation failure — material reduction in claim strength

IF expert_report IS missing:
  → Add to checklist: "Commission independent surveyor or structural engineer report"

IF maintenance_records IS missing:
  → Add to checklist: "Obtain maintenance records, service history, or prior contractor invoices"

IF policy_wording IS missing:
  → Add to checklist: "CRITICAL: Obtain current policy schedule and full policy wording before assessment"
```

---

## Section 7 — Emergency Evidence Preservation Checklist

> This checklist addresses the most valuable learning from the case database. STORM-036 failed primarily because repairs were completed before inspection, damaged materials were discarded, and no photographs existed.

### Before Any Repairs Are Carried Out

The following steps must be completed **before any repair work begins**:

- [ ] Photograph all external damage (roof, walls, guttering, flashing) from multiple angles
- [ ] Photograph all internal damage (ceilings, walls, floors) with timestamps
- [ ] Video walkthrough of all affected areas
- [ ] Retain all damaged materials (broken tiles, sections of felt, flashing)
- [ ] Obtain written statement from emergency contractor before they begin work
- [ ] Notify insurer in writing before temporary repairs start
- [ ] Obtain and save local weather data / storm warnings for date of damage
- [ ] Note names and contact details of any witnesses (neighbours, passers-by)
- [ ] Request emergency contractor provides dated written report identifying storm damage

### After Temporary Repairs

- [ ] Photograph temporary repair work completed
- [ ] Retain invoices from emergency / temporary repair contractors
- [ ] Do not carry out permanent repairs until insurer has inspected
- [ ] Request insurer inspection date in writing

---

## Section 8 — Ombudsman Reasoning Patterns

These are the consistent reasoning patterns observed in UK Financial Ombudsman Service storm claim decisions, derived from the 38-case database.

### Pattern 1 — Storm Occurred ≠ Claim Succeeds

A confirmed storm does not guarantee a valid claim. The customer must still prove:
- The storm caused (not merely revealed) the damage
- The storm was the dominant cause
- There was no pre-existing deterioration that was the real cause

### Pattern 2 — Storm Exposed Deterioration ≠ Storm Caused Damage

If a storm exposes, worsens, or accelerates existing deterioration, the claim typically fails. The deterioration — not the storm — is treated as the effective cause.

### Pattern 3 — Photos Frequently Outweigh Unsupported Assertions

Photographic evidence — particularly contemporaneous photos taken immediately after the storm — consistently outweighs retrospective descriptions or unsupported contractor opinions. Insurers' photos have also defeated contractor evidence (see STORM-003, STORM-035).

### Pattern 4 — Well-Maintained Property = Stronger Claim

Evidence of regular maintenance, service history, and good pre-storm condition is one of the strongest supporting factors for a policyholder. It defeats the wear and tear and gradual deterioration defences.

### Pattern 5 — Multiple Independent Experts Carry Significant Weight

Where multiple independent experts reach the same conclusion, that consensus is treated as near-decisive evidence. This applies whether the experts support or defeat the claim.

### Pattern 6 — Evidence Beats Opinion

Objective evidence (photos, weather data, inspection records, drone surveys) consistently outweighs subjective opinions (customer statements, unsupported roofer assertions).

### Pattern 7 — Insurer Must Prove Exclusions

The burden of proof for exclusions lies with the insurer. An insurer cannot simply allege deterioration, design defect, or poor workmanship — it must prove the exclusion applies. Speculation without supporting evidence is insufficient (see STORM-012, STORM-035).

### Pattern 8 — Insurer Must Inspect Actual Damaged Area

An insurer who relies on inspection of the wrong area of the property, or who fails to inspect the area actually claimed for, acts improperly (see STORM-006). Evidence must relate to the specific claimed damage.

### Pattern 9 — Settlement Method ≠ Coverage Decision

A dispute about how much to pay is separate from whether the claim is covered at all. Coverage must be established before settlement method is addressed (see STORM-018).

---

## Section 9 — AI Workflow Rules — Rules Engine

These rules are deterministic logic instructions to be applied by the AI assessment engine. Rules are applied in sequence and are not overridden by user description alone.

---

### Rule 1 — Policy Exclusion Gate

```
IF:
  damage_type matches a known policy exclusion
  (fences, gates, hedges, or other listed exclusions)

THEN:
  flag exclusion immediately
  classify as: EXCLUDED_CLAIM
  halt further assessment unless policyholder disputes exclusion wording
  request: full policy wording for review
```

### Rule 2 — Storm Verification Gate

```
IF:
  weather_data is absent OR
  weather_data shows conditions below qualifying threshold AND
  no exposed-location factors identified

THEN:
  classify as: STORM_NOT_PROVEN
  flag as HIGH_REJECTION_RISK
  add to missing-evidence checklist: postcode weather data, nearby stations
```

### Rule 3 — Physical Damage Indicator Gate

```
IF:
  no visible storm-related physical damage is present
  (no missing tiles, no displaced flashing, no collapsed structure)

THEN:
  flag as HIGH_REJECTION_RISK
  note: internal damage alone (leak, staining) is insufficient
  add to checklist: external inspection report
```

### Rule 4 — Dominant Cause Assessment

```
IF:
  damage appears consistent with storm conditions
  AND
  no obvious pre-existing deterioration

THEN:
  proceed to expert evidence and maintenance history assessment
  assign preliminary causation score

IF:
  damage appears consistent with gradual deterioration OR
  storm appears to have revealed rather than caused damage

THEN:
  classify as: HIGH_DETERIORATION_RISK
  flag for expert review
  note pattern: storm exposing existing defect = rejection risk
```

### Rule 5 — Maintenance History Assessment

```
IF:
  maintenance records provided AND
  records show regular upkeep AND
  no prior unrepaired damage

THEN:
  increase storm_causation_confidence_score
  reduce deterioration_risk_flag

IF:
  prior unrepaired storm damage exists AND
  current damage in same area

THEN:
  classify as: PRIOR_DAMAGE_UNREPAIRED_RISK
  request: evidence that prior damage was repaired before current storm
```

### Rule 6 — Maintenance Records Boost

```
IF:
  customer provides maintenance records AND
  contractor confirms roof/structure was in good condition pre-storm

THEN:
  increase storm_causation_confidence_score
  reduce wear_and_tear_risk_flag
```

### Rule 7 — Workmanship Exclusion Check

```
IF:
  insurer alleges poor workmanship as rejection basis

THEN:
  check whether policy actually contains a workmanship exclusion
  IF no workmanship exclusion in policy:
    flag insurer rejection as potentially improper
    note: insurer must rely on policy terms, not general assertions
```

### Rule 8 — Wall Collapse Protocol

```
IF:
  storm claim relates to wall collapse
  (boundary wall, garden wall, retaining wall)

THEN:
  request all of:
    - photographs of collapsed wall and foundations
    - age of wall (estimated or documented)
    - maintenance history of wall
    - evidence of vegetation damage (roots, ivy)
    - surveyor or structural engineer report

  classify as: STRUCTURAL_EVIDENCE_REQUIRED
```

### Rule 9 — No Physical Damage Flag

```
IF:
  no visible storm-related physical damage is documented

THEN:
  flag: HIGH_REJECTION_RISK
  note: customer assertion alone is insufficient
  add to checklist: external inspection; photographs; expert report
```

### Rule 10 — Fence Exclusion Check

```
IF:
  damage_type = fence OR gate OR hedge

THEN:
  immediately check policy for specific fence exclusion
  before any further analysis
  IF excluded:
    classify as: EXCLUDED_CLAIM
    halt storm damage assessment for this element
```

### Rule 11 — Pre-Storm Evidence Weight Boost

```
IF:
  insurer has photographs taken before the storm OR
  pre-storm inspection report exists

THEN:
  assign very high evidential weight to that evidence
  note: pre-storm baseline is near-decisive for condition disputes
```

### Rule 12 — Previous Repairs Protocol

```
IF:
  previous repairs to the claimed area are known or alleged

THEN:
  request:
    - repair invoices
    - repair dates
    - contractor reports confirming repair completion
  flag: claim strength depends on proof that prior damage was fully remedied
```

### Rule 13 — Roofer Opinion Validation

```
IF:
  customer relies solely on roofer opinion to establish storm causation

THEN:
  compare against:
    - surveyor report (if available)
    - photographs
    - maintenance records
  note: roofer opinion alone is medium-weight evidence
  note: roofer opinion alone does not outweigh surveyor evidence
```

### Rule 14 — Lifted Tiles + Weather Data Combination

```
IF:
  damage_type includes lifted tiles OR displaced tiles
  AND
  water_ingress present
  AND
  weather records show qualifying wind conditions

THEN:
  flag as: POTENTIALLY_VALID_STORM_CLAIM
  request: inspection report to confirm causation
  increase claim_strength_score
```

### Rule 15 — Weather Data Dispute Protocol

```
IF:
  insurer declines based solely on weather records

THEN:
  verify:
    - postcode-specific weather data (not just regional)
    - nearby weather station data
    - exposed location factors (coastal, elevated, urban wind tunnel)
  note: STORM-022 was upheld because insurer incorrectly concluded no storm conditions existed
  note: weather data validation is a critical workflow step
```

### Rule 16 — Drone Report Weight

```
IF:
  drone inspection report exists

THEN:
  assign high evidential weighting
  note: drone inspections provide objective aerial evidence
  prioritise drone findings in condition and causation assessment
```

### Rule 17 — Gutter Overflow Classification

```
IF:
  water ingress mechanism is guttering overflow AND
  no displaced tiles AND
  no wind damage to roof structure

THEN:
  classify as: MAINTENANCE_RISK_CLAIM
  note: water ingress from overflowing gutters is maintenance failure, not storm damage
  flag: HIGH_REJECTION_RISK unless storm directly caused gutter damage
```

### Rule 18 — Multiple Expert Consensus

```
IF:
  multiple independent experts agree on causation conclusion

THEN:
  assign very high evidential weight to that consensus
  note: applies whether experts support or defeat the claim
  note: consensus across multiple experts is near-decisive
```

### Rule 19 — Policy Storm Definition Check

```
IF:
  policy contains explicit storm definition (e.g., 55 mph / Force 10)

THEN:
  compare weather data directly against policy threshold
  note: if conditions met the threshold, insurer cannot deny on storm grounds
  note: if below threshold, internal policy definition may be decisive (see STORM-027)
  note: even below threshold, exposed location may be relevant (see STORM-011)
```

### Rule 20 — Maintenance Indicators Classification

```
IF:
  evidence shows any of:
    - debris accumulation in valleys
    - blocked guttering valleys
    - deteriorated felt
    - worn or split battens

THEN:
  classify as: LIKELY_MAINTENANCE_RELATED_CLAIM
  flag: HIGH_REJECTION_RISK
  note: these are indicators of ongoing maintenance failure, not storm damage
```

### Rule 21 — Object Impact Claims

```
IF:
  customer alleges damage caused by:
    - falling branch
    - wind-blown debris
    - greenhouse glass impact
    - any flying object

THEN:
  require supporting evidence:
    - photographs of impact point
    - photographs of object or debris
    - contractor report identifying impact damage pattern
  note: customer explanation alone is insufficient (see STORM-029)
```

### Rule 22 — Design Defect Classification

```
IF:
  independent assessor confirms defective design OR
  structural engineer confirms defective construction

THEN:
  classify as: HIGH_REJECTION_RISK_STORM_CLAIM
  note: design defects can override otherwise valid storm conditions
  note: two independent expert opinions defeat even well-maintained property arguments
```

### Rule 23 — Pre-Inspection Repair Flag

```
IF:
  repairs were completed before insurer inspection AND
  no photographs of original damage exist

THEN:
  flag: HIGH_EVIDENTIAL_RISK
  flag: EVIDENCE_PRESERVATION_FAILURE
  add to checklist: obtain contractor statement from repairing contractor
  note: STORM-036 failed primarily on this basis
```

### Rule 24 — Missing Photographs Penalty

```
IF:
  no photographs of original damage exist

THEN:
  reduce claim_strength_score materially
  flag: EVIDENCE_GAP_CRITICAL
  add to checklist: seek any available photos (customer, neighbours, emergency contractors)
```

### Rule 25 — Structural Engineer Weight

```
IF:
  structural engineer report exists

THEN:
  assign very high evidential weighting
  note: structural engineer evidence carries enormous weight in wall and structural claims
  note: applies regardless of which party commissioned the report
```

### Rule 26 — Expert vs Photo Conflict

```
IF:
  insurer expert opinion conflicts with photographic evidence

THEN:
  flag conflict for human review
  request: independent expert review
  note: STORM-035 — photographic evidence can undermine insurer expert opinions
  note: photos showing intact wall before storm contradicted surveyor conclusions
```

### Rule 27 — Sudden Appearance of Gradual Damage

```
IF:
  internal damage (ceiling collapse, water damage) appeared suddenly AND
  investigation reveals underlying leak existed for weeks or months prior

THEN:
  classify as: GRADUAL_DAMAGE_RISK
  flag: HIGH_REJECTION_RISK
  note: damage becoming visible suddenly ≠ damage occurring suddenly
  note: STORM-038 — old skylight with gradual ingress defeated by this pattern
```

---

## Section 10 — Required Inputs

### Mandatory Inputs

| Input | Purpose |
|---|---|
| Policy wording | Check exclusions, storm definition, coverage scope |
| Rejection letter | Identify rejection basis and category |
| Date of loss | Weather data lookup anchor |
| Property postcode | Postcode-specific weather data |
| Damage description | Initial claim type and damage type classification |
| Photographs of damage | Physical damage evidence assessment |

### Recommended Inputs

| Input | Purpose |
|---|---|
| Surveyor / contractor report | Expert causation evidence |
| Weather data (date + postcode) | Storm verification |
| Maintenance records | Deterioration defence rebuttal |
| Previous claims history | Prior unrepaired damage check |
| Repair invoices | Maintenance history documentation |

### Optional Inputs

| Input | Purpose |
|---|---|
| Drone inspection report | High-weight objective condition evidence |
| Structural engineer report | Wall collapse or structural claims |
| Neighbour statements | Corroboration of storm conditions |
| Emergency contractor statement | Pre-repair damage confirmation |

---

## Section 11 — Required Outputs

Every assessment must generate the following outputs:

| Output | Description |
|---|---|
| Claim Summary | Concise summary of claim, damage type, and rejection basis |
| Timeline | Chronological timeline of storm event, damage discovery, repair, and correspondence |
| Rejection Classification | Category from rejection taxonomy with frequency context |
| Evidence Gap Analysis | List of missing evidence and priority level |
| Claim Strength Score | 0–100 score with dimension breakdown |
| Rejection Risk Rating | Very Low / Low / Moderate / High / Very High |
| Recommended Actions | Prioritised list of next steps |
| Missing Evidence Checklist | Specific items to obtain, in priority order |
| Complaint Draft | Draft internal complaint letter if applicable |
| FOS Referral Draft | Draft Financial Ombudsman Service submission if applicable |

---

## Section 12 — JSON Schema

### Claim Assessment Object

```json
{
  "claim_id": "",
  "claim_type": "Storm",
  "damage_type": "",
  "date_of_loss": "",
  "property_postcode": "",
  "insurer": "",
  "rejection_reason": "",
  "rejection_category": "",
  "evidence_provided": [],
  "missing_evidence": [],
  "rules_triggered": [],
  "scores": {
    "storm_verification": 0,
    "physical_damage_evidence": 0,
    "expert_evidence": 0,
    "maintenance_history": 0,
    "causation_strength": 0,
    "total": 0
  },
  "claim_strength_band": "",
  "rejection_risk": "",
  "flags": [],
  "recommended_actions": [],
  "outputs": {
    "claim_summary": "",
    "timeline": [],
    "evidence_gap_analysis": [],
    "complaint_draft": "",
    "fos_draft": ""
  }
}
```

### Flags Reference

| Flag | Meaning |
|---|---|
| `EXCLUDED_CLAIM` | Damage type is excluded by policy |
| `STORM_NOT_PROVEN` | Storm conditions not established |
| `HIGH_REJECTION_RISK` | Multiple rejection indicators present |
| `HIGH_DETERIORATION_RISK` | Strong evidence of pre-existing deterioration |
| `MAINTENANCE_RISK_CLAIM` | Damage consistent with maintenance failure |
| `GRADUAL_DAMAGE_RISK` | Damage likely gradual, not sudden storm event |
| `PRIOR_DAMAGE_UNREPAIRED_RISK` | Prior unrepaired damage in same area |
| `STRUCTURAL_EVIDENCE_REQUIRED` | Structural engineer report essential |
| `EVIDENCE_PRESERVATION_FAILURE` | Repairs completed before inspection |
| `EVIDENCE_GAP_CRITICAL` | Missing evidence is claim-determining |
| `POTENTIALLY_VALID_STORM_CLAIM` | Core indicators support validity |
| `EXPERT_CONFLICT` | Expert opinion conflicts with photographic evidence |

---

## Section 13 — Case Database Reference (38 Cases)

### Case Summaries

| Case ID | Damage Type | Rejection Reason | Outcome | Key Workflow Insight |
|---|---|---|---|---|
| STORM-001 | Gutter / cornice / masonry | Gradual damage from successive storms | Rejected | Distinguish sudden event vs cumulative deterioration |
| STORM-002 | Roof leak | No evidence of actual storm damage | Rejected | Consumer assertion is insufficient without physical storm indicators |
| STORM-003 | Porch roof leak | Wear and tear + poor workmanship | Rejected | Photographic evidence often outweighs unsupported contractor opinion |
| STORM-004 | Roof / bedroom ceiling water ingress | Rotting windowsill and wear and tear | Rejected | Must prove source of ingress, not merely existence of water damage |
| STORM-005 | Flat roof leak | Storm revealed existing defect | Rejected | Storm exposing a defect is not storm causing the damage |
| STORM-006 | Roof tiles, felt, flashing | Insurer relied on wrong area of roof | **Upheld** | Insurer must inspect actual damaged area before applying exclusion |
| STORM-007 | Roof deterioration | Previous storm damage may not have been repaired | Rejected | Claimant must prove current damage is from latest storm, not prior unrepaired damage |
| STORM-008 | Flat roof | Failure to prove storm was dominant cause | Rejected | Maintenance history matters enormously in storm claims |
| STORM-009 | Roof leak / guttering | Poor workmanship / design defect | Rejected | Design defects and workmanship frequently defeat storm claims |
| STORM-010 | Chimney stack | Wear and tear / deterioration | Rejected | Damp worsening after a storm does not prove storm causation |
| STORM-011 | Roof tiles | Insurer relied on 55 mph storm definition | **Upheld** | Internal insurer storm definitions may not be decisive if conditions were reasonably storm-like |
| STORM-012 | Garden wall collapse | Faulty design / workmanship alleged | **Upheld** | Insurer must prove exclusions apply, not merely speculate about defects |
| STORM-013 | Caravan roof water ingress | Design fault / failed seals | Rejected | Multiple independent experts carry significant weight |
| STORM-014 | Bay window structure | Pre-existing structural weakness | Rejected | Storm exposing an existing defect is a recurring rejection pattern |
| STORM-015 | Roof leak | No evidence of actual storm damage | Rejected | Storm warnings alone do not establish storm damage |
| STORM-016 | Boundary fence | Specific policy exclusion | Rejected | Check policy exclusions before analysing causation |
| STORM-017 | Roof leak / ceiling damage | No storm + wear and tear + failed repairs | Rejected | Previous repairs and maintenance history are critical evidence points |
| STORM-018 | Roof | Claim accepted but wear and tear also present | Mixed | Separate coverage decision from settlement dispute |
| STORM-019 | Roof / water ingress | Wear and tear and gradual deterioration | Rejected | Independent roofer opinion alone may not overturn survey evidence |
| STORM-020 | Roof tiles / water ingress | Existing deterioration highlighted by storm | Rejected | Pre-loss inspections are highly persuasive evidence |
| STORM-021 | Roof leaks / external roof damage | Damage not consistent with storm damage | Rejected | Damage alone is insufficient; causation evidence is critical |
| STORM-022 | Lifted roof tiles and water ingress | Insurer incorrectly concluded no storm conditions | **Upheld (reconsideration ordered)** | Weather-data validation is a critical workflow step |
| STORM-023 | Roof tiles / water ingress | No storm conditions + wear and tear | Rejected | Exposed location arguments need supporting evidence |
| STORM-024 | Roof leak | Wear and tear and gradual deterioration | Rejected | Drone inspections can be highly persuasive evidence |
| STORM-025 | Roof / guttering / water ingress | No storm + maintenance issue | Rejected | Distinguish storm damage from maintenance failures |
| STORM-026 | Roof leaks | Long-term deterioration, not one-off storm event | Rejected | Multiple expert opinions outweigh unsupported contractor assertions |
| STORM-027 | Roof and internal water damage | No storm conditions + poor roof condition | Rejected | Policy storm definitions can be decisive in some cases |
| STORM-028 | Roof valley leak | Gradual deterioration and maintenance issues | Rejected | Storm may expose existing problems without causing them |
| STORM-029 | Roof damage | Lack of maintenance and gradual deterioration | Rejected | Customer explanations need supporting evidence |
| STORM-030 | Verge tiles / roof structure | Defective design and workmanship | Rejected | Design defects can override otherwise valid storm conditions |
| STORM-031 | Chimney cowl | No storm conditions + wear and tear | Rejected | Both weather data and material deterioration can independently defeat a claim |
| STORM-032 | Roof and ceiling water damage | Long-term water ingress | Rejected | Internal ceiling collapse alone does not prove storm damage |
| STORM-033 | Boundary wall collapse | Gradual deterioration and defective construction | Rejected | Structural engineer evidence carries enormous weight |
| STORM-034 | Roof tiles, felt and battens | Wear and tear to felt and battens | Partial success | Coverage can exist for some damage but not all resulting repairs |
| STORM-035 | Garden wall collapse | Insurer failed to prove wear and tear was dominant cause | **Upheld** | Photographic evidence can undermine insurer expert opinions |
| STORM-036 | Roof leak | Lack of evidence — repairs completed before inspection | Rejected | Evidence preservation is critical in storm claims |
| STORM-037 | Roof and ceiling | No storm conditions + gradual deterioration | Rejected | Severe weather is not necessarily a storm for insurance purposes |
| STORM-038 | Skylight leak and internal water damage | Gradual cause, not insured event | Rejected | Distinguish sudden damage from damage becoming visible suddenly |

### Outcome Summary

| Outcome | Count | Percentage |
|---|---|---|
| Rejected | 31 | 81.6% |
| Upheld (full) | 4 | 10.5% |
| Upheld (partial / reconsideration) | 2 | 5.3% |
| Mixed | 1 | 2.6% |

### Upheld Cases — What Made the Difference

| Case | Why It Was Upheld |
|---|---|
| STORM-006 | Insurer inspected wrong roof area; contractor evidence confirmed actual damaged area was different |
| STORM-011 | Insurer's internal 55 mph definition not decisive; contractor confirmed well-maintained roof; storm-like conditions existed |
| STORM-012 | Insurer could not produce evidence of alleged design fault or vegetation damage |
| STORM-022 | Weather data showed qualifying conditions existed; insurer incorrectly concluded no storm occurred |
| STORM-035 | Photographs contradicted insurer surveyor's conclusions about wall condition |

---

*Playbook v1 — Built from 38 storm cases. Ready for integration as knowledge base and rules engine.*
