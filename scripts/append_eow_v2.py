"""
Standard append script for Escape of Water Case Database — Schema v2 (21 columns).

Usage
-----
1. Read the source PDF(s) and extract the fields listed in NEW_CASES below.
2. Add one dict per case to NEW_CASES, following the extraction rules in the
   FIELD EXTRACTION RULES section.
3. Run from the repo root:
       py scripts/append_eow_v2.py

The script appends NEW_CASES rows to:
    knowledge/case-databases/Escape_of_Water_Case_Database.xlsx

It automatically assigns the next available row number and applies
consistent alternating-row formatting.

===========================================================================
FIELD EXTRACTION RULES
===========================================================================

IDENTIFICATION FIELDS
─────────────────────
Case ID
    Format: EOW-NNN  (zero-padded to 3 digits, e.g. EOW-016)
    Source: assign sequentially; last used = EOW-015

FOS Decision ID
    Format: DRN-XXXXXXX or DRNXXXXXXX (match exactly as printed in the PDF)
    Source: first line of the PDF, or the "Ref:" header

Insurer Name
    Source: opening complaint paragraph, e.g.
        "Mr X complains that [INSURER NAME] declined his claim…"
    Use the formal registered name as it appears in the decision.
    For Lloyd's syndicates use: "Insurers at Lloyd's (Society of Lloyd's)"
    For brokers: use the broker's name (not the underlying insurer).

FOS Decision Date
    Format: DD Mon YYYY  (e.g. "15 Aug 2023")
    Source: final paragraph of the decision:
        "I'm required to ask [party] to accept or reject my decision
         before [DATE]."
    Use the accept-or-reject deadline date as printed.

SOURCE / PHYSICAL EVENT FIELDS
───────────────────────────────
Claim Type
    Free text. Describe the physical incident and the nature of the dispute
    in one sentence, e.g.:
        "Escape of water — burst supply pipe in kitchen causing floor damage"
    Do NOT embed dispute classification here — that belongs in Dispute Type.

Leak Source
    Free text. Describe the physical origin of the water.  Examples:
        "Supply pipe — copper pipe behind kitchen units, elbow joint failure"
        "WC cistern overflow — cistern cracked and leaking"
        "Tap left on by resident in managed flat"
        "Pre-existing — inherited damage from previous owner, source unknown"

Property Type
    Free text, but use consistent terms:
        "Residential home"
        "Residential home (kitchen)"  — if damage confined to one room
        "Unoccupied residential property"
        "Unoccupied residential property (intended for refurbishment)"
        "Residential home (recently purchased)"
        "Leasehold flat"
        "Commercial / Management Company"

DISPUTE CLASSIFICATION FIELDS
──────────────────────────────
Dispute Type
    Controlled vocabulary — use EXACTLY one of:
        "Coverage Dispute"
            Insurer declined coverage and customer disputed that decision.
        "Handling / Reinstatement Dispute"
            Insurer accepted claim but dispute arose over reinstatement
            scope, quality of work, or settlement quantum.
        "Endorsement / Exclusion Challenge"
            Insurer applied a specific endorsement or exclusion to decline;
            customer challenged its validity or applicability.
        "Pre-Inception Damage Dispute"
            Insurer declined on the basis that damage occurred before the
            policy start date; may overlap with gradual cause.
        "Peril Classification Dispute"
            Dispute is not about coverage but about which peril applies
            (affects excess level or policy section).
        "Claim Recording / Administrative Dispute"
            Complaint concerns how the claim was recorded or administered,
            not the coverage decision itself.
        "Broker Conduct Dispute"
            Complaint concerns a broker's conduct (disclosure, advice,
            renewal notification), not the insurer's claim decision.

Coverage Decision
    What the INSURER originally decided on coverage — not the FOS outcome.
    Controlled vocabulary — use EXACTLY one of:
        "Declined — Full"
            Insurer declined the entire claim.
        "Declined — Partial"
            Insurer accepted part of the claim and declined the remainder.
        "Accepted"
            Insurer accepted the claim without substantive dispute.
        "Accepted — Disputed Settlement"
            Insurer accepted the claim but the settlement amount, scope,
            or reinstatement quality is disputed.
        "Not Applicable"
            No coverage decision was made (admin / broker disputes,
            claim recording errors, etc.).

FOS OUTCOME FIELDS
──────────────────
Outcome Category
    What the FOS decided.  Controlled vocabulary — use EXACTLY one of:
        "Upheld"
            Complaint fully upheld; insurer required to accept or extend
            the claim.
        "Upheld in Part"
            Some elements upheld, others not.  Coverage may be partially
            extended or a combination of coverage and compensation awarded.
        "Not Upheld"
            Insurer's position maintained throughout.
        "Compensation Only"
            Insurer's coverage decline was upheld as correct, but
            compensation was awarded for a separate handling failure
            (e.g. avoidable delays in processing).

Outcome
    Free text.  Full description of what the FOS required the insurer to do.
    Include: settlement instructions, compensation amount, interest
    obligations.  Match detail level of existing rows.

Compensation Awarded (£)
    Numeric (integer).  The total compensation awarded by the FOS for
    distress and inconvenience.
    - Do NOT include claim settlement amounts (these are in Outcome).
    - Do NOT include interest.
    - If no compensation: 0
    - Example: £150 compensation → 150

Is Core Case
    Whether the case should drive rules in the core residential claims
    assessment playbook.
    Controlled vocabulary — use EXACTLY one of:
        "Yes"
            Standard residential claim — drives core playbook rules.
        "No — Administrative"
            Claim recording or administrative dispute; no coverage
            analysis; retain as reference only.
        "No — Handling Dispute"
            Claim was accepted; dispute entirely about reinstatement
            quality or settlement quantum; no coverage principle.
        "No — Commercial"
            Commercial or all-risks policy; classification rules may
            not apply to standard residential policies.
        "No — Broker Dispute"
            Broker conduct / renewal disclosure complaint; no claim
            assessment principle applies.

ANALYSIS FIELDS  (free text — follow existing depth/style)
────────────────────────────────────────────────────────────
Key Policy Clause   — specific contractual wording or FOS/FCA principle applied
Missing Evidence    — evidence that was absent and affected the outcome
Ombudsman Reasoning — how the ombudsman weighed the evidence
Workflow Insight    — operational rule for the claims workflow
AI Rule Candidate   — machine-evaluable rule for the rules engine

Source PDF
    Format: DRNXXXXXXX.pdf  (match filename in knowledge/raw-cases/escape-of-water/)
===========================================================================
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Column definition — must match live workbook exactly
# ---------------------------------------------------------------------------
COLUMNS = [
    "Case ID",
    "FOS Decision ID",
    "Insurer Name",
    "FOS Decision Date",
    "Claim Type",
    "Leak Source",
    "Property Type",
    "Dispute Type",
    "Coverage Decision",
    "Rejection Reason",
    "Evidence Dispute",
    "Outcome Category",
    "Outcome",
    "Compensation Awarded (£)",
    "Is Core Case",
    "Key Policy Clause",
    "Missing Evidence",
    "Ombudsman Reasoning",
    "Workflow Insight",
    "AI Rule Candidate",
    "Source PDF",
]

# Controlled-vocabulary fields — validated on write
CONTROLLED_VOCAB = {
    "Dispute Type": {
        "Coverage Dispute",
        "Handling / Reinstatement Dispute",
        "Endorsement / Exclusion Challenge",
        "Pre-Inception Damage Dispute",
        "Peril Classification Dispute",
        "Claim Recording / Administrative Dispute",
        "Broker Conduct Dispute",
    },
    "Coverage Decision": {
        "Declined — Full",
        "Declined — Partial",
        "Accepted",
        "Accepted — Disputed Settlement",
        "Not Applicable",
    },
    "Outcome Category": {
        "Upheld",
        "Upheld in Part",
        "Not Upheld",
        "Compensation Only",
    },
    "Is Core Case": {
        "Yes",
        "No — Administrative",
        "No — Handling Dispute",
        "No — Commercial",
        "No — Broker Dispute",
    },
}

# Columns that get centred alignment (not wrapped)
CENTRED_COLS = {
    "Case ID", "FOS Decision ID", "Insurer Name", "FOS Decision Date",
    "Property Type", "Dispute Type", "Coverage Decision",
    "Outcome Category", "Compensation Awarded (£)", "Is Core Case",
    "Source PDF",
}

# ---------------------------------------------------------------------------
# NEW CASES — add one dict per case following the extraction rules above.
# Leave empty until you are ready to process the next batch of PDFs.
# ---------------------------------------------------------------------------
NEW_CASES: list[dict] = [
    {
        "Case ID":                  "EOW-016",
        "FOS Decision ID":          "DRN-3022853",
        "Insurer Name":             "Admiral Insurance (Gibraltar) Limited",
        "FOS Decision Date":        "6 Jan 2022",
        "Claim Type":               "Burst pipe in bathroom caused water damage to bathroom and kitchen below; dispute over settlement quantum with insurer attributing most claimed items to wear and tear rather than EOW",
        "Leak Source":              "Burst pipe in bathroom",
        "Property Type":            "Residential home",
        "Dispute Type":             "Coverage Dispute",
        "Coverage Decision":        "Accepted — Disputed Settlement",
        "Rejection Reason":         "Independent surveyor assessed only £1,800 of damage attributable to EOW; remaining items in policyholder's £22,000 quote attributed to wear and tear and betterment",
        "Evidence Dispute":         "Policyholder relied on own tradesman quotes (~£22,000) and submitted video footage; insurer relied on detailed independent surveyor's report itemising each element and its cause",
        "Outcome Category":         "Not Upheld",
        "Outcome":                  "FOS did not require Admiral to increase settlement beyond £1,200 offered (£1,800 assessed EOW damage less £600 excess); Admiral's surveyor-based assessment held to be reasonable and properly evidenced",
        "Compensation Awarded (£)": 0,
        "Is Core Case":             "Yes",
        "Key Policy Clause":        "Policy covers damage caused by EOW only; wear and tear and betterment are excluded; insurer entitled to scope attributable damage using independent surveyor; settlement must restore policyholder to pre-loss position without improvement",
        "Missing Evidence":         "No independent counter-survey submitted by policyholder; videos submitted but did not establish additional EOW-attributable damage",
        "Ombudsman Reasoning":      "FOS cannot make independent technical assessments; assessed whether insurer's process was reasonable; commissioning a detailed survey in response to an unusually high quote was proportionate; insurer reviewed its position when new evidence (videos) was submitted — process was fair",
        "Workflow Insight":         "When policyholder quotes significantly exceed insurer's assessment, insurer may commission an independent survey to distinguish EOW damage from wear and tear; a well-evidenced surveyor's report is sufficient basis for a reduced settlement; FOS will not substitute its own technical judgment",
        "AI Rule Candidate":        "IF settlement_disputed AND insurer_has_detailed_surveyor_report AND report_identifies_wear_and_tear_separately THEN surveyor_based_settlement IS defensible AND FOS_likely = Not_Upheld",
        "Source PDF":               "DRN-3022853.pdf",
    },
    {
        "Case ID":                  "EOW-017",
        "FOS Decision ID":          "DRN-3053156",
        "Insurer Name":             "Aviva Insurance Limited",
        "FOS Decision Date":        "18 Oct 2021",
        "Claim Type":               "Leaking asbestos pipe in garage of unoccupied estate property; Aviva initially mishandled claim, then agreed to replace pipe but applied £1,000 EOW excess; dispute over whether EOW excess was triggered when no water damage to property occurred",
        "Leak Source":              "Leaking asbestos pipe in garage",
        "Property Type":            "Unoccupied residential property",
        "Dispute Type":             "Coverage Dispute",
        "Coverage Decision":        "Accepted — Disputed Settlement",
        "Rejection Reason":         "Aviva applied £1,000 EOW excess endorsement asserting the claim was an escape of water claim; policyholder argued standard £175 excess applied as no water damage resulted from the pipe",
        "Evidence Dispute":         "Policyholder relied on policy schedule (standard excess £175) and photographs showing only a small puddle on the concrete garage floor that dried up and was not part of any repair; Aviva relied on the policy's EOW excess endorsement (£1,000 for damage caused by or resulting from EOW)",
        "Outcome Category":         "Upheld",
        "Outcome":                  "FOS required Aviva to reduce excess to £175 (no EOW damage occurred; claim was for specialist asbestos pipe removal and replacement, not water damage); required Aviva to pay £300 compensation for handling failures if not already paid",
        "Compensation Awarded (£)": 300,
        "Is Core Case":             "Yes",
        "Key Policy Clause":        "EOW excess endorsement (£1,000) applies only to 'damage caused by or resulting from Escape of Water from any Tank Apparatus or Pipe'; where the claim is for pipe replacement only and no such damage occurred, the standard excess applies; Trace and Access cover is only triggered where damage from EOW is present or reasonably anticipated",
        "Missing Evidence":         "Aviva's internal records of how the claim was categorised from the outset; full scope of agreed works under the settlement",
        "Ombudsman Reasoning":      "No damage arose from escape of water — claim was for specialist removal of an asbestos pipe that trickled when water was switched on; the property was unoccupied and facilities unused; Aviva's own handling error led it to cover pipe replacement costs it was not obliged to cover; charging the higher EOW excess on top of that was unreasonable; Aviva gave conflicting excess amounts (£175, £250, £1,000) causing further confusion",
        "Workflow Insight":         "EOW excess only triggers when property damage is caused by or results from escape of water; a claim for pipe removal or replacement without associated water damage attracts only the standard excess; insurers must communicate the basis for excess tier clearly and consistently",
        "AI Rule Candidate":        "IF claim_type = pipe_removal_or_replacement AND no_water_damage_to_property THEN eow_excess = NOT applicable AND standard_excess APPLIES",
        "Source PDF":               "DRN-3053156.pdf",
    },
    {
        "Case ID":                  "EOW-018",
        "FOS Decision ID":          "DRN-3078337",
        "Insurer Name":             "QIC Europe Ltd",
        "FOS Decision Date":        "1 Nov 2021",
        "Claim Type":               "Burst mains stopcock pipe in ground floor bathroom; insurer repudiated trace and access claim on grounds the source was already visible and was outside the policy T&A ambit",
        "Leak Source":              "Mains stopcock pipe, ground floor bathroom",
        "Property Type":            "Residential home",
        "Dispute Type":             "Coverage Dispute",
        "Coverage Decision":        "Declined — Full",
        "Rejection Reason":         "QIC argued T&A cover did not apply because: (1) the mains stopcock was outside the policy T&A ambit; (2) the leak source was already visible and did not need to be located; (3) no structural damage had actually occurred",
        "Evidence Dispute":         "Policyholder relied on plumber's invoice (£1,800 T&A + £500 pipe repair) and 12 photographs; QIC relied on in-house surveying team's review and argued only cupboard removal was needed",
        "Outcome Category":         "Upheld",
        "Outcome":                  "FOS directed QIC to pay T&A costs of £1,800 less £750 policy excess, plus 8% simple interest from date Mr C paid invoice to date of settlement; pipe repair costs (£500) excluded as pipe repairs are excluded under policy terms",
        "Compensation Awarded (£)": 0,
        "Is Core Case":             "Yes",
        "Key Policy Clause":        "T&A cover applies to 'reasonable and necessary costs to remove any part of the building to find the source of the damage caused by water escaping from any fixed water or heating installation, apparatus or pipes'; T&A applies even where the source location is approximately known but cannot be physically accessed; T&A applies even where no damage has yet materialised if the policyholder acted reasonably to mitigate; pipe repair costs remain excluded even under a valid T&A claim",
        "Missing Evidence":         "QIC did not appoint its own contractor to undertake T&A; no independent assessment of what T&A work would have cost from an approved contractor",
        "Ombudsman Reasoning":      "It was unfair to decline T&A because the location was known but not accessible — the specific pipe could not be confirmed until physically reached; if Mr C had not acted, significant damage would likely have resulted; QIC encouraged him to proceed but then repudiated; the T&A invoice was proportionate given the extent of access work required (cupboard section, water softener pipes, waste pipe section, concrete floor); since QIC failed to appoint its own contractor, it cannot cap recovery to a lower hypothetical approved-contractor rate",
        "Workflow Insight":         "T&A cover is triggered when a leak exists and physical access to the source is obstructed, regardless of whether the approximate source location is known; insurers who fail to appoint their own T&A contractor cannot subsequently restrict the policyholder's recovery to hypothetical approved contractor rates; pipe repair costs are excluded even within a valid T&A claim",
        "AI Rule Candidate":        "IF eow_confirmed AND leak_source_approximately_known AND physical_access_blocked THEN T&A_cover = applicable; IF insurer_did_not_appoint_own_T&A_contractor THEN insurer_cannot_reduce_recovery_to_approved_contractor_rate",
        "Source PDF":               "DRN-3078337.pdf",
    },
    {
        "Case ID":                  "EOW-019",
        "FOS Decision ID":          "DRN-3121008",
        "Insurer Name":             "AXA Insurance UK Plc",
        "FOS Decision Date":        "19 Nov 2021",
        "Claim Type":               "Leasehold flat; storage tank overflowed twice (16 months apart) from the same component (ball valve and overflow pipe); insurer registered second incident as a separate claim with a second excess; policyholder argued second incident resulted from the insurer's ineffective first repair",
        "Leak Source":              "Storage tank overflow — faulty ball valve and overflow pipe (both incidents)",
        "Property Type":            "Leasehold flat",
        "Dispute Type":             "Coverage Dispute",
        "Coverage Decision":        "Accepted — Disputed Settlement",
        "Rejection Reason":         "AXA asserted the second escape of water had a different cause to the first and should be treated as a new claim attracting a new excess of £1,500",
        "Evidence Dispute":         "AXA relied on its 2020 report claiming a different fault; policyholder noted AXA's 2019 and 2020 reports both reference ball valve and overflow pipe failure; FOS observed: same component failed both times; second incident was the first real test of the repair (Ms K had not been away for any extended period between incidents); gap of 16 months does not demonstrate an effective repair",
        "Outcome Category":         "Upheld",
        "Outcome":                  "AXA must: refund second policy excess of £1,500; remove record of second claim; reimburse cost of second carpet damaged in second incident; pay £250 compensation for inconvenience and upset caused as previously offered if not already paid; reimburse any previously agreed but unpaid works",
        "Compensation Awarded (£)": 250,
        "Is Core Case":             "Yes",
        "Key Policy Clause":        "Insurer that confirms a repair is complete and explicitly reassures the policyholder the same incident will not recur becomes responsible for damage if the same component subsequently fails; second incident should be treated as an extension of the first claim where on balance it results from an ineffective repair; gap between incidents does not demonstrate repair effectiveness if the system was never tested during that period",
        "Missing Evidence":         "No independent engineering report from AXA demonstrating the 2020 fault was technically distinct from and unrelated to the 2019 repair",
        "Ombudsman Reasoning":      "Same component (ball valve and overflow) failed on both occasions; the second incident occurred the next time the system was actually tested (Ms K went away for the first time since the first repair); AXA explicitly reassured Ms K the problem was fixed and she would not have the same issue if she went away — this reassurance transferred responsibility for an ineffective repair to AXA; the 16-month gap cannot be treated as evidence of an effective and lasting repair",
        "Workflow Insight":         "If an insurer confirms a repair is complete and explicitly reassures a policyholder the same event will not recur, the insurer becomes liable for the second incident if the same component fails again; the second incident should be assessed as an extension of the first claim unless the insurer demonstrates a technically distinct and unrelated cause",
        "AI Rule Candidate":        "IF second_eow_incident AND same_component_failed AND insurer_confirmed_first_repair_complete AND policyholder_sought_explicit_reassurance THEN second_incident = extension_of_first_claim AND second_excess = NOT payable",
        "Source PDF":               "DRN-3121008.pdf",
    },
    {
        "Case ID":                  "EOW-020",
        "FOS Decision ID":          "DRN-3517894",
        "Insurer Name":             "AA Underwriting Insurance Company Limited",
        "FOS Decision Date":        "5 Aug 2022",
        "Claim Type":               "Two EOW incidents: original kitchen damage from internal soil stack leak (accepted); second incident — bathroom floor lifting from bath waste pipe — occurred while first claim repairs were being completed; insurer registered second as a new claim; policyholder argued second leak caused by pipe disturbance during first repair",
        "Leak Source":              "First: internal soil stack (kitchen); Second: bath waste pipe (disturbed during first repair)",
        "Property Type":            "Residential home",
        "Dispute Type":             "Coverage Dispute",
        "Coverage Decision":        "Accepted — Disputed Settlement",
        "Rejection Reason":         "AA argued the second leak was from clean water (bath waste) and therefore unrelated to the soil stack claim; asserted the bath was not disturbed during the original repair works",
        "Evidence Dispute":         "Policyholder provided photographs showing toilet and bath share an interconnected pipework system leading to a single outlet pipe; scope of original works confirmed toilet pan, hand basin and vanity unit were removed and refitted; AA argued the soil stack repair was unlikely to displace bath waste pipework",
        "Outcome Category":         "Upheld",
        "Outcome":                  "FOS directed AA Underwriting to: remove record of second claim from all internal and external databases; pay costs to trace and repair bath waste pipe including excess paid under home emergency policy with 8% simple interest; pay repair costs for all EOW damage; pay £200 compensation within 28 days",
        "Compensation Awarded (£)": 200,
        "Is Core Case":             "Yes",
        "Key Policy Clause":        "A second EOW incident occurring during or directly attributable to repair of the first claim should not be recorded as a separate claim; insurer cannot register a secondary leak as a new claim if it is more likely than not caused by disturbance from the repair works; no second excess is payable",
        "Missing Evidence":         "No independent plumbing expert report on whether removal and refitting of toilet and hand basin could have displaced the bath waste pipe",
        "Ombudsman Reasoning":      "Pipework from toilet, bath and hand basin all interconnected and lead into one common pipe; the scope of works for the first claim included removal and refitting of the toilet and hand basin; on balance of probabilities, disturbing this interconnected pipework likely nudged the bath waste pipe out of place; the second leak was more likely than not attributable to the first repair",
        "Workflow Insight":         "When a second EOW incident occurs during or shortly after repair of the first claim, assess whether the repair works disturbed connected pipework before registering a new claim; if pipework is interlinked and the repair scope included adjacent fittings, the second incident should be assessed as part of the first claim",
        "AI Rule Candidate":        "IF second_eow_leak AND first_claim_repair_in_progress AND pipework_interconnected AND repair_scope_included_adjacent_fittings THEN second_leak = likely_attributable_to_first_repair AND new_claim_registration = NOT appropriate",
        "Source PDF":               "DRN-3517894.pdf",
    },
    {
        "Case ID":                  "EOW-021",
        "FOS Decision ID":          "DRN-3606995",
        "Insurer Name":             "Ageas Insurance Limited",
        "FOS Decision Date":        "1 Sep 2022",
        "Claim Type":               "Two separate bathroom EOW incidents (August 2019: bath waste pipe seal failure; February 2020: pop-up plug/waste mechanism failure); disputes over multiple excesses, scope and quantum of repairs, alternative accommodation period, electrical damage, and VAT withholding",
        "Leak Source":              "Bath waste pipe seal failure (first, August 2019); pop-up plug/waste mechanism failure (second, February 2020) — both attributed to old age and wear",
        "Property Type":            "Residential home",
        "Dispute Type":             "Handling / Reinstatement Dispute",
        "Coverage Decision":        "Accepted",
        "Rejection Reason":         "Ageas accepted both claims (£500 cash settlement for bathroom repairs) but refused to treat two EOW incidents as one claim; declined to extend alternative accommodation beyond 9 March 2020; declined electrical repair costs (no expert evidence linking to EOW); withheld VAT on cash settlement until proof of payment; settled only on scope identified by surveyor",
        "Evidence Dispute":         "Policyholder alleged structural instability, electrical damage, and extensive additional damage beyond surveyor's findings; Ageas relied on its own surveyor's May 2020 report (leaks localised and short-lived; no structural damage; electrics unrelated to EOW); policyholder's own loss adjuster friend agreed floor instability was from swelling laminate, not structural weakness",
        "Outcome Category":         "Not Upheld",
        "Outcome":                  "FOS confirmed Ageas did not need to do anything more; £200 compensation already paid by Ageas was sufficient; surveyor-based £500 settlement was fair; alternative accommodation covered only to 9 March 2020; two excesses correctly charged for two separate incidents with different physical causes; electrical repairs excluded (no expert evidence linking to EOW); VAT withholding until proof of payment is standard practice",
        "Compensation Awarded (£)": 0,
        "Is Core Case":             "Yes",
        "Key Policy Clause":        "Each separate EOW incident with a different physical cause attracts its own excess; policy covers repair to pre-loss condition only (like-for-like); VAT not payable on cash settlements until proof of payment to a VAT-registered contractor is provided; alternative accommodation ceases when insurer makes its position clear and property is not demonstrably uninhabitable",
        "Missing Evidence":         "Independent electrician's report linking electrical failure to EOW; itemised invoices separating EOW repair costs from wider renovation work; expert evidence linking floor instability to EOW rather than laminate swelling",
        "Ombudsman Reasoning":      "Two incidents had genuinely different physical causes (seal failure vs pop-up mechanism failure); insurer's surveyor and policyholder's own loss adjuster friend agreed bathroom was not uninhabitable; no expert evidence linked electrical failure to EOW; insurer's expert evidence was strong and consistent at each decision point; £200 compensation was proportionate to communication failures",
        "Workflow Insight":         "Two EOW incidents with different physical causes are separate claims attracting separate excesses; insurer's surveyor report is the primary basis for scoping repairs and assessing habitability; electrical damage requires expert evidence linking failure specifically to EOW; alternative accommodation liability ceases when insurer communicates clearly that the property is habitable and payments will stop",
        "AI Rule Candidate":        "IF two_eow_incidents AND physically_different_causes THEN separate_claims AND separate_excesses = applicable; IF electrical_damage_claimed AND no_expert_report_linking_to_eow THEN electrical_claim = NOT covered",
        "Source PDF":               "DRN-3606995.pdf",
    },
    {
        "Case ID":                  "EOW-022",
        "FOS Decision ID":          "DRN-3860121",
        "Insurer Name":             "Tesco Underwriting Limited",
        "FOS Decision Date":        "12 Jan 2023",
        "Claim Type":               "Three central heating pipe leaks in screed floor (kitchen, hallway, bathroom) discovered during major property renovation; insurer accepted claim but offered only £3,180.10, excluding items it considered renovation-related and reducing costs via unsupported surveyor annotations",
        "Leak Source":              "Central heating pipes embedded in screed floor — kitchen, hallway and bathroom",
        "Property Type":            "Residential home",
        "Dispute Type":             "Handling / Reinstatement Dispute",
        "Coverage Decision":        "Accepted — Disputed Settlement",
        "Rejection Reason":         "Tesco argued many claimed items were renovation-related not EOW-related; applied surveyor's pencil cost annotations based on local knowledge without supporting market evidence; declined alternative accommodation invoice due to address error on invoice",
        "Evidence Dispute":         "Tesco relied on surveyor's unsubstantiated cost annotations; Mr C relied on original surveyor damage report, damage restoration specialist's moisture readings confirming affected rooms, itemised schedule of costs, and receipts/invoices; FOS found Tesco's annotations lacked market evidence",
        "Outcome Category":         "Upheld",
        "Outcome":                  "FOS directed Tesco to pay £11,600.29 total settlement (kitchen/hallway drying and rescreeding £1,533; bathroom/en-suite strip-out, tiling and reinstatement £3,403.61; bedroom joinery, flooring and redecoration £4,765.68; skip hire £348; T&A investigative work £350; alternative accommodation £1,200) less any sum already paid, plus 8% statutory interest",
        "Compensation Awarded (£)": 0,
        "Is Core Case":             "Yes",
        "Key Policy Clause":        "Insurer must handle claims promptly and fairly; surveyor cost reductions require supporting market evidence to be persuasive; only rooms identified in contemporaneous professional reports as EOW-affected are eligible for settlement; investigative work (T&A) is covered; alternative accommodation is covered for the period the property is genuinely uninhabitable due to EOW repairs",
        "Missing Evidence":         "Itemised invoices clearly separating EOW repair costs from renovation costs; building schedule and control documents requested by Tesco but not provided; market cost evidence to support Tesco's reduced settlement figures",
        "Ombudsman Reasoning":      "Tesco's 4-month delay in accepting the claim was unjustified — surveyor's report was available by end of November; Mr C's decision to proceed with repairs during the delay was reasonable given ongoing renovation and builders already on site; surveyor's pencil annotations without documentary support were not persuasive — insurer must evidence that materials and labour could be sourced at the annotated rates; coverage limited to rooms identified in contemporaneous damage assessment reports",
        "Workflow Insight":         "A prolonged delay in accepting a claim forces the policyholder to proceed without authorisation; FOS will hold the insurer responsible for reasonable repair costs incurred during that delay; surveyor cost reductions must be evidenced by market data, not bare assertions; rooms not identified in the contemporaneous damage report are excluded even if costs were incurred there",
        "AI Rule Candidate":        "IF claim_acceptance_delay > 60_days AND policyholder_proceeds_with_repairs THEN insurer_cannot_reject_reasonable_costs_incurred_during_delay; IF surveyor_cost_reduction AND no_market_evidence_provided THEN cost_reduction = NOT accepted",
        "Source PDF":               "DRN-3860121.pdf",
    },
    {
        "Case ID":                  "EOW-023",
        "FOS Decision ID":          "DRN-4205492",
        "Insurer Name":             "QIC Europe Limited",
        "FOS Decision Date":        "8 Sep 2023",
        "Claim Type":               "Two EOW incidents: first was bathroom leak damaging kitchen (accepted and repaired by QIC's contractor); second was a kitchen leak from the area QIC's contractor had been working on, a few weeks after first repair; insurer argued second was an independent frozen pipe event and opened a new claim",
        "Leak Source":              "First: bathroom pipe; Second: isolator valve in kitchen fitted by QIC's contractor (cracked weeks after installation)",
        "Property Type":            "Residential home",
        "Dispute Type":             "Coverage Dispute",
        "Coverage Decision":        "Accepted — Disputed Settlement",
        "Rejection Reason":         "QIC argued the second leak was caused by water freezing in the pipe (outside temperature -3.6°C on the day) and was therefore an independent event warranting a separate claim and separate excess",
        "Evidence Dispute":         "QIC provided outside weather data (-3.6°C) and an industry article on frozen pipe claims; policyholder argued QIC's contractor's valve was defective or incorrectly fitted; FOS noted outside temperature is not the same as enclosed indoor pipe temperature; the freeze-claims article QIC cited referred to dates 10 days after Mr A's claim",
        "Outcome Category":         "Upheld",
        "Outcome":                  "FOS directed QIC to: treat both leaks as one claim including kitchen floor tiles; charge only one excess and refund second if already paid; pay £150 compensation for service shortfalls and avoidable delays",
        "Compensation Awarded (£)": 150,
        "Is Core Case":             "Yes",
        "Key Policy Clause":        "A second leak originating from an area where the insurer's contractor recently worked is prima facie attributable to the contractor's work unless the insurer establishes an independent cause; a valve cracking within weeks of installation suggests manufacturing defect or incorrect fitting; insurer is responsible for its appointed contractor's workmanship",
        "Missing Evidence":         "Engineering inspection of the cracked valve to determine whether failure was from freeze damage or installation defect; indoor temperature data for the property on the day of the leak",
        "Ombudsman Reasoning":      "Valve cracked within weeks of installation in the specific area QIC's contractor had worked on; outside temperature of -3.6°C does not establish that an enclosed kitchen cupboard pipe was at freezing temperatures; freeze-claim surge article cited by QIC referred to a later date; on balance, valve more likely failed due to defect or improper installation; QIC is responsible for its contractor's work",
        "Workflow Insight":         "When a second EOW incident occurs in the area recently worked on by the insurer's contractor, the presumption is the contractor caused it; insurer cannot rebut this by citing weather data alone without establishing indoor pipe temperatures; insurer is liable for its contractor's workmanship failures",
        "AI Rule Candidate":        "IF second_leak_location = area_of_recent_insurer_contractor_work AND time_since_repair < 60_days THEN second_leak = presumptively_attributable_to_contractor AND separate_claim = NOT appropriate UNLESS independent_cause_established",
        "Source PDF":               "DRN-4205492.pdf",
    },
    {
        "Case ID":                  "EOW-024",
        "FOS Decision ID":          "DRN-4223988",
        "Insurer Name":             "HDI Global Speciality SE",
        "FOS Decision Date":        "8 Jan 2024",
        "Claim Type":               "Residential EOW claim for damage from leaking 15mm water supply pipe under floor of rear extension (2019); insurer declined after initial inspection found no evidence of leak; policyholder later produced builder's report confirming pipe found and capped, and water bills showing anomalous usage during the claim period",
        "Leak Source":              "15mm water supply pipe under floor in rear extension",
        "Property Type":            "Residential home",
        "Dispute Type":             "Coverage Dispute",
        "Coverage Decision":        "Declined — Full",
        "Rejection Reason":         "HDI's inspection (June 2019) found no evidence of an escape of water; water mains sound test found no water supply leaks; HDI questioned why the builder's report was not shared until long after the 2019 inspection",
        "Evidence Dispute":         "HDI relied on June 2019 inspection report (no leak found) and water mains sound test; policyholder provided: (1) builder's statement (November 2019) with photographs confirming leaking pipe found and capped, bathroom left inoperable; (2) water bills showing usage spike from ~15-19m3 per period to 63m3 in Feb-Aug 2019, reducing back to normal after pipe was capped in November 2019",
        "Outcome Category":         "Upheld",
        "Outcome":                  "FOS directed HDI to pay Ms L £600 compensation for handling failures; HDI must reconsider the EOW claim based on new evidence; if the claim proceeds to settlement, HDI must settle subject to remaining policy terms and limits (note: this direction does not affect the separate declined subsidence claim)",
        "Compensation Awarded (£)": 600,
        "Is Core Case":             "Yes",
        "Key Policy Clause":        "Insurer must investigate EOW claims thoroughly; where an initial inspection is inconclusive and leaves open the possibility of an undetected leak, the insurer should not issue a final decline without pursuing further investigation; subsequently produced evidence (builder's report, utility bills) must be properly evaluated before a final position is maintained",
        "Missing Evidence":         "Builder's report and water bills were not provided to HDI until after the initial inspection; no independent assessment of whether damp damage was caused by EOW vs subsidence or rising damp (separate subsidence claim exists)",
        "Ombudsman Reasoning":      "Water bills showing usage rising from ~15-19m3 per period to 63m3 during the claimed EOW period, then reducing to normal after the pipe was capped, constitutes strong circumstantial evidence of an ongoing underground leak; builder's report confirms leaking pipe was found and capped; HDI's own inspection acknowledged a remote possibility of an undetected pipe leak; on balance, sufficient evidence that an EOW did take place; £600 compensation is proportionate to HDI's handling failures over several years",
        "Workflow Insight":         "Water utility bill anomalies — a significant usage spike coinciding with the claimed EOW period, returning to normal after the source is capped — are strong circumstantial evidence of a leak even where an initial physical inspection found no visible damage; insurers should consider requesting utility bills as standard evidence in EOW investigations, especially where underground or concealed pipework is involved",
        "AI Rule Candidate":        "IF initial_inspection_inconclusive AND water_bills_show_usage_spike_during_claim_period AND usage_returns_to_normal_after_capping THEN EOW = likely_occurred AND full_decline = NOT defensible",
        "Source PDF":               "DRN-4223988.pdf",
    },
    {
        "Case ID":                  "EOW-025",
        "FOS Decision ID":          "DRN-4227214",
        "Insurer Name":             "Saga Services Limited",
        "FOS Decision Date":        "11 Oct 2023",
        "Claim Type":               "Unoccupied residential property; leak occurred and claim made; insurer declined under endorsement limiting cover to FLEA only (fire, smoke, lightning, aircraft impact); complainants alleged mis-selling and that the EOW exclusion was not made clear during the sales call",
        "Leak Source":              "Not specified (property had a leak)",
        "Property Type":            "Unoccupied residential property",
        "Dispute Type":             "Endorsement / Exclusion Challenge",
        "Coverage Decision":        "Declined — Full",
        "Rejection Reason":         "Policy contained an endorsement restricting cover to FLEA only, explicitly excluding EOW for properties unoccupied for more than 60 consecutive days; endorsement was in place and disclosed at point of sale",
        "Evidence Dispute":         "Complainants alleged sales agent verbally confirmed EOW cover during the call; Saga relied on sales call recordings showing agent stated FLEA-only cover at least twice and read the endorsement in full; IPID stated EOW excluded for properties unoccupied 60+ consecutive days; no recording evidence that EOW was confirmed as covered",
        "Outcome Category":         "Not Upheld",
        "Outcome":                  "FOS did not uphold; Saga adequately disclosed the FLEA-only endorsement during the sales call and in policy documents including the IPID; £250 compensation already paid by Saga for administrative errors was sufficient; FOS could not reasonably ask Saga to do anything more",
        "Compensation Awarded (£)": 0,
        "Is Core Case":             "Yes",
        "Key Policy Clause":        "Unoccupied property endorsement restricting cover to FLEA only (excluding EOW) is a standard and not unusual market practice; insurer adequately discharges its disclosure obligation by reading the endorsement in full during the sales call and providing a clear IPID stating the exclusion; a non-advised sale places the onus on the policyholder to verify the policy meets their needs",
        "Missing Evidence":         "No recording evidence that the sales agent verbally confirmed EOW cover (FOS reviewed recordings — no such statement found); complainants produced no evidence they queried EOW cover during the call",
        "Ombudsman Reasoning":      "Sales call recordings confirmed agent stated FLEA-only cover at least twice and read the endorsement in full; IPID clearly stated EOW excluded for properties unoccupied 60+ consecutive days; endorsement is standard and not unusual for unoccupied properties; as a non-advised sale, onus was on complainants to ensure the policy met their needs; at no point during the call did the agent state EOW was covered",
        "Workflow Insight":         "Where an unoccupied property endorsement excludes EOW and the insurer can demonstrate the endorsement was clearly disclosed during the sales call and in policy documents, the decline is defensible; FOS will not uphold mis-selling claims where the insurer's sales records demonstrate adequate, unambiguous disclosure",
        "AI Rule Candidate":        "IF policy_type = unoccupied AND endorsement_excludes_eow AND insurer_read_endorsement_during_sales_call AND IPID_states_eow_excluded THEN mis_selling_claim = NOT upheld AND coverage_decline = defensible",
        "Source PDF":               "DRN-4227214.pdf",
    },
]


# ---------------------------------------------------------------------------
# Style helpers
# ---------------------------------------------------------------------------
def _row_fill(even: bool) -> PatternFill:
    color = "D6E4F0" if even else "FFFFFF"
    return PatternFill(start_color=color, end_color=color, fill_type="solid")


def _border() -> Border:
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------
def _validate(case: dict) -> None:
    case_id = case.get("Case ID", "?")
    for field, allowed in CONTROLLED_VOCAB.items():
        val = case.get(field, "")
        if val not in allowed:
            raise ValueError(
                f"{case_id} — '{field}' value '{val}' not in controlled vocab.\n"
                f"  Allowed: {sorted(allowed)}"
            )
    if not isinstance(case.get("Compensation Awarded (£)", 0), (int, float)):
        raise TypeError(
            f"{case_id} — 'Compensation Awarded (£)' must be a number."
        )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> None:
    if not NEW_CASES:
        print("NEW_CASES is empty — nothing to append.")
        return

    # validate controlled-vocab fields before touching the file
    for case in NEW_CASES:
        _validate(case)

    repo_root = os.path.normpath(os.path.join(os.path.dirname(__file__), ".."))
    xlsx_path = os.path.join(
        repo_root, "knowledge", "case-databases",
        "Escape_of_Water_Case_Database.xlsx",
    )

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active

    # verify live schema matches COLUMNS
    live_headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    if live_headers != COLUMNS:
        mismatch = [(i+1, live_headers[i] if i < len(live_headers) else "—", COLUMNS[i])
                    for i in range(max(len(live_headers), len(COLUMNS)))
                    if i >= len(live_headers) or i >= len(COLUMNS)
                    or live_headers[i] != COLUMNS[i]]
        raise RuntimeError(
            "Live workbook columns do not match COLUMNS definition.\n"
            "Mismatches (col, live, expected):\n" +
            "\n".join(f"  {c}: '{l}' vs '{e}'" for c, l, e in mismatch)
        )

    cell_font = Font(name="Calibri", size=10)
    border    = _border()
    wrap      = Alignment(wrap_text=True, vertical="top")
    centre    = Alignment(horizontal="center", vertical="top")

    first_new_row = ws.max_row + 1

    for i, case in enumerate(NEW_CASES):
        row_idx = first_new_row + i
        fill = _row_fill(row_idx % 2 == 0)
        for col_idx, col_name in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=case.get(col_name, ""))
            cell.font      = cell_font
            cell.fill      = fill
            cell.border    = border
            cell.alignment = centre if col_name in CENTRED_COLS else wrap
        ws.row_dimensions[row_idx].height = 120

    wb.save(xlsx_path)

    last_row   = ws.max_row
    last_case  = ws.cell(row=last_row, column=1).value
    total_data = last_row - 1

    print(f"Appended {len(NEW_CASES)} case(s) to {xlsx_path}")
    print(f"Total data rows : {total_data}")
    print(f"Last Case ID    : {last_case}")


if __name__ == "__main__":
    main()
