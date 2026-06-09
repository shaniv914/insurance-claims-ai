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
        "Case ID":                  "EOW-026",
        "FOS Decision ID":          "DRN3376494",
        "Insurer Name":             "Covea Insurance plc",
        "FOS Decision Date":        "26 Apr 2018",
        "Claim Type":               "Water from neighbour's property entered under floorboards causing EOW damage; dispute over whether EOW or flood peril applies and which excess is payable",
        "Leak Source":              "Neighbour's escape of water — entered under floorboards from adjacent property's fixed water/drainage installation",
        "Property Type":            "Residential home",
        "Dispute Type":             "Peril Classification Dispute",
        "Coverage Decision":        "Accepted — Disputed Settlement",
        "Rejection Reason":         "Covea insisted the claim must proceed under EOW peril (not flood), attracting the higher £500 excess; Mr M argued the water was an 'ingress' from next door attracting only the standard £250 flood excess",
        "Evidence Dispute":         "Mr M characterised the water as 'ingress' from a neighbour and proposed a flood claim (£250 excess); Covea and FOS confirmed the water originated from the neighbour's fixed installation, qualifying as EOW; FOS investigator applied the dual-peril principle (most beneficial to consumer) but found the flood alternative no more beneficial here",
        "Outcome Category":         "Not Upheld",
        "Outcome":                  "FOS did not require Covea to change the peril classification; EOW excess of £500 properly applied; £50 previously paid by Covea for the initial excess miscalculation was confirmed as reasonable",
        "Compensation Awarded (£)": 0,
        "Is Core Case":             "Yes",
        "Key Policy Clause":        "EOW wording covers loss or damage caused by escape of water from a fixed water or drainage installation — the water does not have to originate from within the insured property; no standalone 'ingress of water' cover exists as a separate peril; where two perils could both apply, the most beneficial to the consumer should be used",
        "Missing Evidence":         "Policy schedule confirming the flood excess vs EOW excess differential; no assessment of whether the neighbour's source installation could independently qualify under the flood peril definition",
        "Ombudsman Reasoning":      "Water clearly came from a neighbour's fixed installation and fell within the EOW peril definition; there is no separate ingress cover under the policy; FOS confirmed the dual-benefit principle but concluded the flood peril would not have produced a more beneficial outcome for Mr M; Covea was entitled to apply the EOW excess",
        "Workflow Insight":         "When an EOW originates from a neighbouring property's fixed installation, the EOW peril applies regardless of the direction of travel into the insured property; the 'ingress of water' characterisation does not create a lower-excess alternative; always check whether an alternative peril (e.g. flood) would produce a lower excess before settling on EOW classification",
        "AI Rule Candidate":        "IF eow_from_neighbours_fixed_installation AND policy_has_no_ingress_cover THEN peril = EOW AND eow_excess APPLIES; IF two_perils_potentially_applicable THEN use_peril_most_beneficial_to_policyholder",
        "Source PDF":               "DRN3376494.pdf",
    },
    {
        "Case ID":                  "EOW-027",
        "FOS Decision ID":          "DRN3405029",
        "Insurer Name":             "Aviva Insurance Limited",
        "FOS Decision Date":        "30 Sep 2018",
        "Claim Type":               "Escape of water from shower valve or shower arm concealed behind tiles causing damp wall in tenanted residential property; insurer declined citing poor workmanship and lack of maintenance",
        "Leak Source":              "Shower valve or shower arm behind bathroom tiles (concealed within wall)",
        "Property Type":            "Residential home (tenanted)",
        "Dispute Type":             "Coverage Dispute",
        "Coverage Decision":        "Declined — Full",
        "Rejection Reason":         "Aviva's contractors attributed damage to poor workmanship (oversized tile hole at shower handle outlet, gap between shower screen and wall, limescale on bath trap, cracked tiles) and lack of maintenance; declined without applying a specific named policy exclusion",
        "Evidence Dispute":         "Aviva relied on contractors' surface-level observations without tracing the leak source through to the wall; Mr B's independent plumber removed the shower plate and found the wall behind it saturated, identifying the concealed shower valve/arm as the specific source; FOS found the independent report more persuasive as it identified a physical source directly linked to the observed damage",
        "Outcome Category":         "Upheld",
        "Outcome":                  "FOS required Aviva Insurance Limited to: settle Mr B's claim including trace and access costs subject to remaining policy terms; pay the cost of Mr B's independent plumber's report; pay £250 compensation to Mr B",
        "Compensation Awarded (£)": 250,
        "Is Core Case":             "Yes",
        "Key Policy Clause":        "Policy covers 'escape of water from any tank, apparatus or pipe' — this includes concealed apparatus and pipes behind tiles; policy has no exclusion for poor maintenance or defective workmanship applicable to the EOW peril; absence of a specific gradual-deterioration exclusion on the EOW contingency means gradual onset alone cannot justify a decline; insurer is responsible for inadequate claim handling by its delegated contractors",
        "Missing Evidence":         "Aviva's contractors did not remove the shower plate to inspect behind it; no contemporaneous photographs of the damp wall area taken by Aviva's contractors at time of inspection to substantiate their findings",
        "Ombudsman Reasoning":      "Mr B's plumber removed the shower plate and found the wall saturated behind it — direct evidence of a concealed source; Aviva's contractor report identified surface issues only without tracing the cause through to the wall; the policy has no workmanship or maintenance exclusion on the EOW peril; insurer is responsible for poor handling by delegated contractors; tenant's children's health exposure was a relevant factor in awarding £250 compensation",
        "Workflow Insight":         "Insurer's contractors must physically trace the leak source before declining on workmanship grounds; surface-level observations (tile gaps, limescale) are insufficient to override a specific independent expert finding of a concealed leak; where a policy lacks a workmanship/maintenance exclusion on the EOW peril, poor maintenance alone cannot justify a full decline; T&A costs are recoverable as part of the claim",
        "AI Rule Candidate":        "IF insurer_declines_on_workmanship_grounds AND policy_has_no_workmanship_exclusion_on_eow THEN decline = NOT valid; IF independent_expert_identifies_concealed_pipe_source AND insurer_expert_did_not_inspect_behind_fixture THEN independent_report = more_persuasive",
        "Source PDF":               "DRN3405029.pdf",
    },
    {
        "Case ID":                  "EOW-028",
        "FOS Decision ID":          "DRN4208888",
        "Insurer Name":             "Millennium Insurance Company Limited",
        "FOS Decision Date":        "11 Aug 2017",
        "Claim Type":               "Shower broke causing single escape of water that flooded the shower cubicle, damaging tiling, base, floor and wall; insurer declined asserting damage was not caused by a single escape of water and excluded as a gradually operating cause",
        "Leak Source":              "Shower fitting failure — shower valve or mechanism broke, causing water to flood the cubicle (single incident)",
        "Property Type":            "Residential home",
        "Dispute Type":             "Coverage Dispute",
        "Coverage Decision":        "Declined — Full",
        "Rejection Reason":         "Millennium's loss adjuster concluded the damage was not caused by a single escape of water; cited wet rot in the shower's timber frame and black mould on the sealant as evidence of gradual deterioration; relied on the gradually operating cause exclusion to decline",
        "Evidence Dispute":         "Millennium relied on the loss adjuster's report citing wet rot on the timber frame and black mould on the sealant as indicators of gradual water ingress over time; Mr and Mrs L provided photographs showing the mould was surface-only and wiped off without penetrating the sealant, and a purchase receipt for shower repair parts confirming a specific repair event; loss adjuster confirmed the wood was dry at the time of his visit",
        "Outcome Category":         "Upheld",
        "Outcome":                  "Millennium Insurance Company Limited must pay Mr and Mrs L's claim subject to policy limits and excesses; if Mr and Mrs L have already paid for any repairs that should have been covered, Millennium must add interest at 8% simple per year from the date of the claim to the date of payment; within 28 days of acceptance",
        "Compensation Awarded (£)": 0,
        "Is Core Case":             "Yes",
        "Key Policy Clause":        "EOW peril covers a single escape of water when a shower breaks; the gradually operating cause exclusion does not apply where the primary claimed damage (tiles, floor, wall) was caused by a single incident even if pre-existing wear exists in a separate unclaimed component (timber frame); wet rot in an adjacent component that was not claimed for does not bring the entire claim under the gradual exclusion; a purchase receipt for shower repair parts is sufficient corroboration of a specific incident",
        "Missing Evidence":         "No photographic record of the shower valve or mechanism that broke; loss adjuster did not physically test the mould (e.g. wipe it) to determine whether it had penetrated the sealant",
        "Ombudsman Reasoning":      "Mr and Mrs L's photographs showed mould wiped off the sealant surface without penetrating — directly contradicts the gradual ingress theory; wet rot in the timber frame was a pre-existing separate issue that Mr and Mrs L were not claiming for; dry wood on inspection is inconsistent with ongoing sealant leakage; a receipt for shower repair parts corroborated the specific incident; Millennium failed to demonstrate the gradually operating cause exclusion applied to the damage being claimed",
        "Workflow Insight":         "Loss adjusters should physically test mould on sealant (e.g. wipe it) before concluding it evidences gradual ingress; wet rot in an unclaimed component does not automatically bring all damage in the same room under the gradual exclusion; a purchase receipt for repair parts is adequate corroborating evidence of a specific incident in the absence of other documentation",
        "AI Rule Candidate":        "IF claimed_damage_is_tiles_floor_wall AND separate_component_has_wet_rot AND policyholder_not_claiming_for_rotten_component THEN wet_rot_does_not_invoke_gradual_exclusion; IF mould_on_sealant AND not_tested_for_penetration THEN mould = insufficient_evidence_of_gradual_ingress",
        "Source PDF":               "DRN4208888.pdf",
    },
    {
        "Case ID":                  "EOW-029",
        "FOS Decision ID":          "DRN-4307523",
        "Insurer Name":             "QIC Europe Ltd",
        "FOS Decision Date":        "6 Oct 2023",
        "Claim Type":               "Burst pipe in loft caused extensive EOW damage in December 2022; insurer accepted claim and produced a scope of works but excluded stair tread damage alleged to have been caused by the EOW and/or the subsequent drying process",
        "Leak Source":              "Burst pipe in loft space (December 2022)",
        "Property Type":            "Residential home",
        "Dispute Type":             "Handling / Reinstatement Dispute",
        "Coverage Decision":        "Accepted — Disputed Settlement",
        "Rejection Reason":         "QIC's surveyor (inspecting in March 2023, three months after the EOW and two months after drying completed) assessed the cracked stair treads as historic, based on paint flecks inside the cracks and the absence of fresh clean timber; argued the stair damage was pre-existing and should be excluded from the scope",
        "Evidence Dispute":         "QIC relied on its surveyor's March 2023 inspection report citing paint flecks and timber appearance; Miss M argued water poured down the stairs from the burst pipe, that a painted plasterboard ceiling had collapsed above the stairs (explaining the paint flecks), and that the delayed inspection affected timber appearance; FOS noted stair damage was entirely absent from QIC's initial December 2022 damage survey",
        "Outcome Category":         "Upheld",
        "Outcome":                  "QIC Europe Ltd must provide a revised scope of works including the stair tread damage with a revised settlement offer at current repair costs; pay Miss M £400 compensation for distress and inconvenience caused over approximately seven months",
        "Compensation Awarded (£)": 400,
        "Is Core Case":             "Yes",
        "Key Policy Clause":        "Where an insurer's initial survey makes no reference to damage in a specific area, it cannot subsequently assert that damage in that area was pre-existing; paint flecks in stair tread cracks explained by debris from a collapsed plasterboard ceiling above are not sufficient proof of historic damage; delayed inspections (months after the EOW and after drying completion) compromise reliance on timber appearance indicators such as fresh timber vs aged cracks",
        "Missing Evidence":         "QIC's initial December 2022 survey makes no reference to the stair treads — its absence is the most significant evidential gap; no independent assessment of whether the drying process itself could crack timber stair treads",
        "Ombudsman Reasoning":      "QIC's own initial survey (December 2022) did not reference any stair damage — if the cracking was truly historic and clearly observable, it should have been documented then; the inspection took place three months post-EOW and two months after drying completed, so timber appearance would have changed; Miss M's explanation of the paint flecks (debris from a collapsed plasterboard ceiling above the stairs) is a plausible alternative to QIC's historic-damage theory; QIC failed to definitively demonstrate the damage was pre-existing",
        "Workflow Insight":         "An initial EOW survey must document all pre-existing damage in all areas inspected — failure to do so prevents the insurer from later asserting damage in those areas is pre-existing; drying-process damage (timber cracking) must be assessed promptly after drying completion; delayed post-drying inspections undermine the reliability of appearance-based pre-existing damage arguments",
        "AI Rule Candidate":        "IF damage_not_noted_in_initial_survey AND insurer_later_asserts_pre_existing THEN pre_existing_claim = weak AND burden_of_proof = high; IF inspection_delayed_post_drying AND insurer_relies_on_timber_appearance THEN appearance_evidence = unreliable",
        "Source PDF":               "DRN-4307523.pdf",
    },
    {
        "Case ID":                  "EOW-030",
        "FOS Decision ID":          "DRN-4368751",
        "Insurer Name":             "AXA Insurance Limited",
        "FOS Decision Date":        "28 Nov 2023",
        "Claim Type":               "Standing water found in cellar; claimant alleged burst pipe caused flooding; insurer declined asserting damage was caused by rising damp from an untanked cellar conversion, not an escape of water",
        "Leak Source":              "Disputed — Mr R alleged a burst or leaking pipe behind wall plasterboard; AXA's surveyor concluded damage was consistent with rising damp due to absence of cellar tanking and poor cellar conversion",
        "Property Type":            "Residential home",
        "Dispute Type":             "Coverage Dispute",
        "Coverage Decision":        "Declined — Full",
        "Rejection Reason":         "AXA's surveyor found no physical evidence of flooding — no tidemarks on walls, no water damage to items on the cellar floor (cardboard boxes, books, bookshelves undamaged), rotten chipboard subfloor, extensive generalised damp; absence of tanking made the observed damp damage inevitable; concluded cause was rising damp and gradually operating deterioration, not an escape of water",
        "Evidence Dispute":         "AXA relied on surveyor B's contemporaneous inspection report (shortly after the claim was reported) and on the physical observations of undamaged floor items; Mr R relied on a contractor's invoice describing 'fixed the leak burst causing flooding' and a separate letter describing near-knee-level flooding requiring 13 hours to clear; FOS preferred B's inspection-based findings over the contractor's later account, which was inconsistent with the physical evidence; a subsequent undated contractor invoice including tanking (£1,500) and damp-proofing slurry (£310) further supported the damp interpretation",
        "Outcome Category":         "Not Upheld",
        "Outcome":                  "FOS did not uphold — Mr R failed to establish that an EOW caused the damage; AXA's surveyor's inspection was more persuasive than the contractor's later account; the gradual operating cause exclusion was fairly applied; AXA required to take no further action",
        "Compensation Awarded (£)": 0,
        "Is Core Case":             "Yes",
        "Key Policy Clause":        "Policyholder bears the initial onus of proving an insured peril (EOW) caused the damage before coverage obligations arise; absence of expected physical signs of flooding (tidemarks, water-damaged floor items) undermines an EOW claim even where a contractor describes flooding; where damage is more consistent with gradual damp penetration (generalised damp, rot, untanked cellar), the gradual operating cause exclusion applies",
        "Missing Evidence":         "Independent expert inspection contemporaneous with the event to establish whether flooding occurred; itemised breakdown of the contractor's £8,700 invoice separating T&A and repair elements; explanation for why cardboard boxes, books and bookshelves on the cellar floor showed no water damage if flooding was near-knee level",
        "Ombudsman Reasoning":      "Surveyor B's report was drawn from an inspection shortly after the claim and is contemporaneous evidence; physical observations (no tidemarks, items on floor undamaged, rotten chipboard, extensive generalised damp) are inconsistent with flooding of the scale described by Mr R; contractor's letter was prepared later and directly contradicted by the physical evidence; subsequent invoice for tanking and damp-proofing slurry supported the damp interpretation; policyholder failed to discharge the initial burden of proving EOW; gradual cause exclusion fairly applied",
        "Workflow Insight":         "A contemporaneous surveyor inspection finding no physical flooding indicators (tidemarks, water-damaged floor items) carries substantially more weight than a later contractor account; a contractor invoice for tanking and damp-proofing is strong corroborating evidence for damp as the cause rather than EOW; policyholders bear the initial burden of establishing an insured peril and cannot simply rely on a contractor's general description of flooding",
        "AI Rule Candidate":        "IF cellar_damage_claimed AND surveyor_finds_no_tidemarks AND floor_items_undamaged AND no_tanking_present THEN eow_peril = NOT established AND gradual_cause_exclusion = applicable; IF contractor_invoice_includes_tanking_works THEN damp_interpretation_supported",
        "Source PDF":               "DRN-4368751.pdf",
    },
    {
        "Case ID":                  "EOW-031",
        "FOS Decision ID":          "DRN4464315",
        "Insurer Name":             "Fairmead Insurance Limited",
        "FOS Decision Date":        "10 Jul 2020",
        "Claim Type":               "Second escape of water at same property in March 2019 (four months after a prior December 2018 EOW claim); claimant alleged the previous emergency plumber damaged the bath causing a new leak damaging bathroom joists and kitchen ceiling; insurer declined for lack of physical evidence of EOW damage",
        "Leak Source":              "Disputed — alleged leak under the bath (bath claimed to have been damaged by the previous home emergency insurer's plumber); loss adjuster found no evidence of an active leak or water damage at the property",
        "Property Type":            "Residential home",
        "Dispute Type":             "Coverage Dispute",
        "Coverage Decision":        "Declined — Full",
        "Rejection Reason":         "Fairmead's loss adjuster attended approximately three weeks after the alleged EOW and found no evidence of water damage to the kitchen ceiling or bathroom joists; joists were concealed under floorboards that showed no moisture damage; bath had already been removed before inspection; adjuster noted no evidence of insured damage and concluded there was nothing to cover under the EOW section",
        "Evidence Dispute":         "Mrs I claimed kitchen ceiling watermarks were still visible and the loss adjuster acknowledged the joists needed repair; Fairmead's adjuster reported Mrs I had told him the kitchen ceiling had already been repaired, and he observed no damage there; joists were hidden under undamaged floorboards with no moisture evidence; FOS could not resolve the conflicting recollections but found the absence of physical evidence supported the decline",
        "Outcome Category":         "Not Upheld",
        "Outcome":                  "FOS did not uphold the EOW element — absence of physical evidence meant Fairmead acted reasonably in declining; £2,000 offered by Fairmead for accidental damage to the bath during the home emergency plumber's visit was fair; Fairmead fulfilled its liability under the policy",
        "Compensation Awarded (£)": 0,
        "Is Core Case":             "Yes",
        "Key Policy Clause":        "Insured must establish that EOW damage occurred during the policy term; where a loss adjuster attends and finds no physical evidence of water damage (no visible ceiling marks, no moisture in exposed floorboards above claimed damaged joists), an insurer is entitled to decline; removal of the source component (the bath) before the loss adjuster's inspection substantially undermines a claim in the absence of other contemporaneous evidence",
        "Missing Evidence":         "Contemporaneous photographs of kitchen ceiling watermarks and any exposed joists taken before the bath was removed; independent assessment of joist condition at the time of the alleged incident; a definitive record of whether Mrs I had or had not already repaired the kitchen ceiling before the adjuster visited",
        "Ombudsman Reasoning":      "Loss adjuster found nothing to repair in the kitchen — Mrs I said she had not carried out the repair but the adjuster saw no damage (discrepancy unexplained but absence of evidence supports the decline); joists were hidden under undamaged floorboards with no moisture evidence; bath had already been removed before inspection, preventing direct assessment of the alleged source; in the absence of contradictory evidence it was reasonable for Fairmead to rely on the adjuster's on-site assessment; £2,000 for accidental bath damage was a fair separate settlement",
        "Workflow Insight":         "Claims for concealed EOW damage (e.g. joists under floorboards) require contemporaneous evidence — once the source component is removed or access is closed before inspection, the policyholder's ability to establish damage is significantly compromised; insurers should advise policyholders not to remove or repair source components before a loss adjuster inspects; undamaged floorboards above claimed damaged joists make a joist damage claim very difficult to sustain",
        "AI Rule Candidate":        "IF loss_adjuster_finds_no_visible_water_damage AND source_component_removed_before_inspection AND claimed_damage_concealed THEN eow_decline = defensible; IF joists_claimed_damaged AND overlying_floorboards_show_no_water_damage THEN joist_damage_claim = weak",
        "Source PDF":               "DRN4464315.pdf",
    },
    {
        "Case ID":                  "EOW-032",
        "FOS Decision ID":          "DRN-4521660",
        "Insurer Name":             "esure Insurance Limited",
        "FOS Decision Date":        "16 Jan 2024",
        "Claim Type":               "Escape of water from cold-water pipe in loft en-suite causing water to pour through ceiling throughout the home; insurer declined under faulty workmanship exclusion, attributing the cause to the installer of the en-suite (Person A)",
        "Leak Source":              "Cold-water pipe in loft en-suite — exact source not established by expert inspection; FOS concluded on balance most likely attributable to faulty workmanship by the en-suite installer (Person A), who attended and resolved the leak by working on the handbasin area then became uncontactable",
        "Property Type":            "Residential home",
        "Dispute Type":             "Coverage Dispute",
        "Coverage Decision":        "Declined — Full",
        "Rejection Reason":         "esure concluded the leak arose from Person A's faulty workmanship in installing the loft en-suite; Person A had recently installed the bathroom, attended as a nearby neighbour within minutes, pulled the handbasin from the wall and did something that stopped the leak, then refused all subsequent contact; esure relied on the faulty workmanship exclusion",
        "Evidence Dispute":         "Mr and Mrs E argued esure's conclusion was purely circumstantial — neither esure's loss adjuster nor a second plumber (Company P) could establish the cause of the leak; FOS agreed Person A's behaviour (attending immediately as a neighbour, stopping the leak by working on his recently installed component, then refusing all contact with both policyholders and esure) was more consistent with awareness of personal liability than with innocence",
        "Outcome Category":         "Not Upheld",
        "Outcome":                  "FOS did not uphold — esure fairly applied the faulty workmanship exclusion on the balance of probabilities; no further compensation directed beyond £100 already paid by esure for complaint handling delays",
        "Compensation Awarded (£)": 0,
        "Is Core Case":             "Yes",
        "Key Policy Clause":        "Once an insured peril is established, the onus shifts to the insurer to demonstrate a valid exclusion applies; faulty workmanship exclusion applies where, on balance of probabilities, a leak is attributable to a tradesperson's work in the same location; where expert inspections cannot identify a source, circumstantial evidence (location of leak, timing relative to installation, installer's post-incident avoidance behaviour) can establish the exclusion; the exclusion need only be established on balance of probabilities, not beyond reasonable doubt",
        "Missing Evidence":         "No direct evidence from Person A as to what caused the leak or what he did to stop it; no independent report confirming whether the specific component Person A worked on was defective or incorrectly installed",
        "Ombudsman Reasoning":      "An EOW clearly occurred — water damage throughout the home was not in dispute; neither esure's loss adjuster nor a second plumber could find the leak source on inspection, consistent with a concealed installation fault; Person A installed the en-suite, attended within 5–10 minutes as a neighbour, worked on his recently installed handbasin area, stopped the leak without explanation, and then refused all contact with both the policyholders and esure — a pattern consistent with awareness of personal liability; FOS upheld the faulty workmanship exclusion on this circumstantial basis",
        "Workflow Insight":         "Where a tradesperson recently installed components in the area of a subsequent EOW, attended quickly to resolve it without explanation, and then became uncontactable, the faulty workmanship exclusion is sustainable even without direct forensic evidence of the defect; installer avoidance of contact after the incident is probative circumstantial evidence; the exclusion need only be established on balance of probabilities",
        "AI Rule Candidate":        "IF eow_confirmed AND installer_recently_completed_work_in_eow_area AND installer_resolved_leak_without_explanation AND installer_subsequently_unavailable THEN workmanship_exclusion = applicable_on_balance_of_probabilities",
        "Source PDF":               "DRN-4521660.pdf",
    },
    {
        "Case ID":                  "EOW-033",
        "FOS Decision ID":          "DRN-4704763",
        "Insurer Name":             "Covea Insurance plc",
        "FOS Decision Date":        "22 Apr 2024",
        "Claim Type":               "Leaseholder's rented flat covered by a residents association block insurance policy; escape of water from a nail-damaged pipe and water ingress from chimney stack base caused a severe dry rot outbreak throughout the subfloor; insurer declined both the EOW coverage claim and trace and access under gradually operating cause exclusion",
        "Leak Source":              "Pipe damaged by a nail (adjacent to living room door); separate water ingress from base of chimney stack; air bricks at rear of property covered over, blocking subfloor ventilation",
        "Property Type":            "Leasehold flat",
        "Dispute Type":             "Coverage Dispute",
        "Coverage Decision":        "Declined — Full",
        "Rejection Reason":         "Covea concluded that while a pipe leak (EOW event) did occur, the primary and dominant damage — a severe dry rot outbreak throughout the timber subfloor causing structural collapse — was caused by a gradually operating process; areas of dry rot had been filled and decorated over, indicating Mr K was or ought to have been aware of the damage; T&A cover was declined as it is contingent on there being non-excluded insurable damage resulting from the EOW",
        "Evidence Dispute":         "Mr K's contractor confirmed the nail-damaged pipe, chimney water ingress, blocked air bricks, and a severe dry rot outbreak with structural subfloor collapse; areas of rot had been filled and decorated over — FOS held this meant Mr K knew or ought to have known of the gradual damage; Mr K argued the filled areas were in completely different locations to the EOW damage area; FOS re-read the contractor's report and maintained that areas of rot had been filled and decorated, supporting constructive knowledge",
        "Outcome Category":         "Not Upheld",
        "Outcome":                  "FOS did not uphold — gradually operating cause exclusion fairly applied to the EOW damage claim; T&A cover not available as it requires underlying non-excluded damage resulting from EOW, which does not exist here; Covea not required to take any further action",
        "Compensation Awarded (£)": 0,
        "Is Core Case":             "No — Commercial",
        "Key Policy Clause":        "Under this block policy, T&A cover is contingent on 'Damage' (defined as accidental loss/destruction/damage unless otherwise excluded) resulting from an escape of water — where all EOW-related damage has been excluded under the gradually operating cause exclusion, T&A cannot apply independently; constructive awareness of pre-existing gradual damage is assessed from physical evidence (filled and decorated areas of rot indicate owner awareness); multiple concurrent water ingress sources increase the likelihood that damage is attributed to gradually operating cause rather than a single EOW event",
        "Missing Evidence":         "No independent assessment of when the dry rot outbreak began relative to the EOW event; no definitive evidence as to whether the filled and decorated rot areas were genuinely in separate locations to the EOW damage zone",
        "Ombudsman Reasoning":      "A pipe leak confirms an EOW insured event occurred, but the primary damage was a severe dry rot outbreak throughout the subfloor consistent with prolonged moisture exposure; areas of rot had been filled and decorated over — Mr K was aware or ought to have been aware; T&A cover requires underlying non-excluded damage — since all damage is excluded under gradually operating cause, T&A cannot respond independently; even if water damage existed separately from dry rot, it would most likely have happened gradually given the severity of the outbreak and multiple concurrent ingress sources",
        "Workflow Insight":         "Block/property owners policy T&A cover is typically contingent on there being insurable (non-excluded) damage — T&A cannot apply independently where the underlying damage is excluded under gradually operating cause; a policyholder who fills and decorates over rot cannot claim ignorance of pre-existing gradual damage; multiple concurrent water ingress sources (pipe, chimney, blocked ventilation) increase the risk of a gradually operating cause finding",
        "AI Rule Candidate":        "IF eow_confirmed AND primary_damage_is_dry_rot AND filled_and_decorated_over_rot_areas_present THEN gradual_cause_exclusion = applicable; IF damage_excluded_under_gradual_cause AND ta_cover_requires_non_excluded_damage THEN ta_cover = NOT applicable",
        "Source PDF":               "DRN-4704763.pdf",
    },
    {
        "Case ID":                  "EOW-034",
        "FOS Decision ID":          "DRN-4744346",
        "Insurer Name":             "AXIS Specialty Europe SE",
        "FOS Decision Date":        "15 May 2024",
        "Claim Type":               "Commercial property EOW claim for damage from a leaking shower at a rented flat; insurer avoided the policy back to March 2022 renewal for non-disclosure of material facts (property condition — damp, rot, structural issues); EOW claim also independently declined for gradual cause and policyholder failure to mitigate",
        "Leak Source":              "Disputed — possibly deteriorated sealant around shower tray (gradual cause, as suggested by scope of works) or shower controller unit and pipework (as Mr H's contractor later maintained); exact source unclear due to conflicting contractor evidence",
        "Property Type":            "Residential home (tenanted, commercial landlord policy)",
        "Dispute Type":             "Coverage Dispute",
        "Coverage Decision":        "Declined — Full",
        "Rejection Reason":         "Policy avoided back to March 2022 renewal: Mr H attended a structural engineer's inspection in early March 2022 that identified damp, rot and structural issues; as a professional landlord he ought to have disclosed this material information at the renewal; AXIS would not have offered cover had it known. EOW claim independently declined: policy covers EOW only where caused by freezing — no freeze evidence; gradual deterioration exclusion applied; Mr H failed to maintain the property and delayed remediation by several months, prejudicing AXIS's position",
        "Evidence Dispute":         "Mr H argued he received the engineer's report only after the renewal date (April 2022); AXIS said he was physically present at the March 2022 inspection and ought to have been aware of conditions; conflicting contractor evidence on the EOW source (original estimate listed shower tray works; later contractor email cited the shower controller unit); multiple concurrent water ingress routes to the basement (chimney, external walls, lack of ventilation) alongside the EOW; FOS upheld the avoidance and found the EOW claim independently declined on both multiple-cause and mitigation failure grounds",
        "Outcome Category":         "Not Upheld",
        "Outcome":                  "FOS did not uphold — policy avoidance was valid under Insurance Act 2015 (qualifying breach of the duty of fair presentation, treated as neither deliberate nor reckless; premiums relating to the relevant property returned); EOW claim also independently declined as EOW not established as primary cause and policyholder's delays and insufficient remediation prejudiced AXIS's position; AXIS required to take no further action",
        "Compensation Awarded (£)": 0,
        "Is Core Case":             "No — Commercial",
        "Key Policy Clause":        "Insurance Act 2015 duty of fair presentation — a commercial policyholder must disclose material circumstances they know or ought to know; presence at a property inspection establishes constructive knowledge of the condition witnessed; AXIS's policy covers EOW only where caused by freezing of fixed water tanks, apparatus or pipes; gradual deterioration/wet rot/dry rot exclusion applicable to all sections; policyholder duty to keep buildings in good repair and take all reasonable steps to prevent damage; remedies for qualifying breach (non-deliberate/non-reckless) = avoidance with premium refund",
        "Missing Evidence":         "Engineer's report not received by Mr H until after renewal (though he was present at the inspection); definitive independent expert evidence on the specific EOW source (initial estimate and later contractor email contradicted each other); photographs submitted by Mr H showing no basement rot in 2021 were limited in scope and not persuasive of overall basement condition",
        "Ombudsman Reasoning":      "Mr H was physically present at the March 2022 inspection — as a professional landlord he ought to have appreciated the property's condition from what he observed even before receiving the formal report; AXIS would not have offered cover had it known — qualifying breach established; avoidance with premium refund is the correct remedy for a non-deliberate non-reckless breach under Insurance Act 2015; even if the policy could be read to cover EOW beyond freezing, multiple concurrent water ingress sources (chimney, external walls, poor ventilation) mean EOW cannot be established as the primary cause of the damage; significant delay between EOW identification and remediation, plus insufficiently extensive repairs in March/April 2022, prejudiced AXIS's position",
        "Workflow Insight":         "Commercial landlords have a heightened duty of fair presentation at renewal following property inspections even before the formal engineer's report is received — presence at the inspection is sufficient to establish constructive knowledge; where a commercial property policy restricts EOW cover to freeze-caused escapes, policyholders face a materially higher coverage hurdle; multiple concurrent water ingress sources reduce the probability of EOW being the primary cause and strengthen a gradually operating cause defence",
        "AI Rule Candidate":        "IF commercial_policy AND professional_landlord_attended_inspection_revealing_material_facts AND not_disclosed_at_renewal THEN qualifying_breach = likely AND policy_avoidance = available_under_insurance_act_2015; IF policy_limits_eow_to_freezing AND no_evidence_of_freezing THEN eow_claim = NOT covered",
        "Source PDF":               "DRN-4744346.pdf",
    },
    {
        "Case ID":                  "EOW-035",
        "FOS Decision ID":          "DRN-4749282",
        "Insurer Name":             "Aviva Insurance Limited",
        "FOS Decision Date":        "10 Jul 2024",
        "Claim Type":               "December 2022 EOW accepted by Aviva and reinstated (including asbestos ceiling removal); two subsequent disputes: (1) alleged damage to bathroom caused by Aviva's contractors during asbestos ceiling removal; (2) separate claim for water damage to bathroom wall tiles and substrate attributed to the original EOW, which Aviva declined",
        "Leak Source":              "Escape of water in December 2022 (specific source not identified in this decision — the primary EOW claim was accepted and reinstated; this decision concerns only the subsequent reinstatement and follow-on damage disputes)",
        "Property Type":            "Residential home",
        "Dispute Type":             "Handling / Reinstatement Dispute",
        "Coverage Decision":        "Accepted",
        "Rejection Reason":         "(1) Contractor damage: Aviva's asbestos report confirmed working areas were visually clean post-removal; Aviva found no evidence of specific damage to shower enclosure, tiles or fixtures caused during its contractors' works; (2) Tile/substrate damage: post-repair moisture readings confirmed bathroom was dry; Aviva attributed tile adhesion failure to the quality of the October 2022 bathroom installation, not to the December 2022 EOW",
        "Evidence Dispute":         "Miss M and Miss M's contractor described widespread tile adhesion failure and substrate damage but did not attribute it to the EOW; policyholders theorised freeze-thaw expansion of water infiltrating the tile substrate during the sub-zero temperatures of December 2022 and subsequent winter months; Aviva relied on dry post-repair moisture readings and attributed tile failure to October 2022 installation quality; FOS found no expert evidence linking either the alleged contractor damage or the tile adhesion failure to the EOW",
        "Outcome Category":         "Not Upheld",
        "Outcome":                  "FOS did not uphold either element of the complaint — insufficient evidence that Aviva's contractors caused bathroom damage during reinstatement; insufficient evidence that the December 2022 EOW caused the tile and substrate adhesion failure; policyholders' contractor report did not link the damage to the EOW; no expert opinion establishing EOW causation; Aviva not required to take any further action",
        "Compensation Awarded (£)": 0,
        "Is Core Case":             "Yes",
        "Key Policy Clause":        "Policyholders bear the burden of proving an insured peril caused the claimed damage; the assumption that damage occurred during a contractor's repair visit is insufficient without supporting expert evidence; a contractor's report describing damage without identifying its cause does not discharge the policyholder's evidential burden; a lay theory about a freeze-thaw mechanism carries materially less weight than an expert opinion establishing EOW causation",
        "Missing Evidence":         "Independent expert report linking the tile and substrate adhesion failure to the December 2022 EOW or to subsequent moisture exposure; independent survey of the bathroom condition between the EOW (December 2022) and Aviva's repair completion to identify when tile loosening began; any expert evidence that Aviva's contractors caused the specific damage alleged during asbestos ceiling removal",
        "Ombudsman Reasoning":      "No expert evidence was produced to show Aviva's contractors caused the bathroom damage — photographs before and after asbestos removal did not show the specific areas where damage was later found; Miss M and Miss M's contractor described tile adhesion failure without attributing it to the EOW; without expert opinion linking tile failure to EOW, the policyholder's burden is not discharged; the lay freeze-thaw theory is not equivalent to expert opinion and cannot substitute for it; Aviva's dry moisture readings post-repair were not independently challenged",
        "Workflow Insight":         "Policyholders claiming contractor damage during reinstatement, or EOW-caused follow-on damage to a specific component (tiles, substrate), must obtain an expert report linking the damage to the claimed cause; a contractor's report describing damage without identifying cause is insufficient; where an insurer's post-repair moisture readings show dry conditions and are not independently challenged, they are a strong basis for declining a tile or substrate damage claim",
        "AI Rule Candidate":        "IF tile_damage_claimed_post_eow AND contractor_report_does_not_link_to_eow AND no_expert_opinion_establishing_eow_causation THEN tile_claim = NOT established; IF contractor_damage_alleged_during_repairs AND no_expert_report THEN contractor_damage_claim = NOT established",
        "Source PDF":               "DRN-4749282.pdf",
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
