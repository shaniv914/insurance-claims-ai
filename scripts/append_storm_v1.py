"""
Standard append script for Storm Case Database — Schema v1 (21 columns).

Usage
-----
1. Read the source PDF(s) and extract the fields listed in NEW_CASES below.
2. Add one dict per case to NEW_CASES following the extraction rules.
3. Run from the repo root:
       py scripts/append_storm_v1.py

Appends NEW_CASES rows to:
    knowledge/case-databases/Storm_Case_Database.xlsx

===========================================================================
FIELD EXTRACTION RULES
===========================================================================

Case ID         : Format STORM-NNN (zero-padded to 3 digits)
FOS Decision ID : DRN-XXXXXXX or DRNXXXXXXX as printed in the PDF
Insurer Name    : Formal registered name from the FOS decision
FOS Decision Date : DD Mon YYYY — accept-or-reject deadline in final paragraph
Claim Type      : Physical incident and nature of dispute in one sentence
Leak Source     : Physical storm damage mechanism / source of water ingress
                  e.g. "Wind-lifted roof tiles allowing water ingress"
                       "Storm debris impact on roof structure"
                       "Wind-driven rain through failed pointing"
Property Type   : "Residential home" / "Unoccupied residential property" /
                  "Leasehold flat" / "Commercial" / etc.
Dispute Type    : Controlled vocab (7 values)
Coverage Decision : Controlled vocab (5 values)
Rejection Reason  : Insurer's stated reason for declining
Evidence Dispute  : What evidence each party relied on
Outcome Category  : Controlled vocab (4 values)
Outcome           : Full FOS remedy instructions
Compensation Awarded (£) : Integer — D&I only; 0 if none
Is Core Case      : Controlled vocab (5 values)
Key Policy Clause : Policy wording or FOS/FCA principle applied
Missing Evidence  : Evidence that was absent and affected the outcome
Ombudsman Reasoning : How the ombudsman weighed the evidence
Workflow Insight  : Operational rule for the claims workflow
AI Rule Candidate : Machine-evaluable rule for the rules engine
Source PDF        : Filename only (e.g. DRN-1207086.pdf)
===========================================================================
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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

CENTRED_COLS = {
    "Case ID", "FOS Decision ID", "Insurer Name", "FOS Decision Date",
    "Property Type", "Dispute Type", "Coverage Decision",
    "Outcome Category", "Compensation Awarded (£)", "Is Core Case",
    "Source PDF",
}

# ---------------------------------------------------------------------------
# NEW CASES — populate before running
# ---------------------------------------------------------------------------
NEW_CASES: list[dict] = [
    {
        "Case ID": "STORM-021",
        "FOS Decision ID": "DRN-4517146",
        "Insurer Name": "AXA XL Insurance Company UK Limited",
        "FOS Decision Date": "23 Apr 2024",
        "Claim Type": "Storm damage to balcony tiles on residential property partially declined; insurer paid internal damage but declined external tile replacement and scaffolding costs on grounds of wear and tear and sub-storm wind speeds",
        "Leak Source": "Tiles lifted from balcony (one year old) allowing rainwater ingress into building interior; property in exposed position but no corroborating evidence of locally enhanced wind speeds",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Partial",
        "Rejection Reason": "Weather records showed wind speeds insufficient for storm conditions; expert surveyor report cited wear and tear rather than storm damage with photos from the balcony and exterior; scaffolding cost excluded as consequential to declined external repair",
        "Evidence Dispute": "AXA: weather records showing no storm-level winds, expert surveyor report citing wear and tear with balcony and exterior photographs. Policyholders: one-year-old tiles should not have failed without storm force, property in exposed position with argued higher localised winds, logical wind/rain causation — conceded balcony floor wear and tear but disputed tile damage",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; AXA's decision upheld as fair and in line with policy terms; tile damage and scaffolding decline confirmed; internal damage settlement not disturbed",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Storm peril requires violent winds; burden of proof on policyholder to evidence storm conditions; an unsubstantiated claim of enhanced localised wind speeds due to exposed position requires corroborating evidence; insurer entitled to rely on expert surveyor report where report includes on-site photographs; scaffolding costs excluded where the underlying external repair is itself not storm-caused",
        "Missing Evidence": "Corroborating evidence of enhanced wind speeds at the exposed property location; independent expert assessment showing tile damage pattern inconsistent with wear and tear",
        "Ombudsman Reasoning": "Weather records for the period did not indicate storm-level wind speeds; policyholders could not provide corroborating evidence that localised conditions were more severe than nearby stations; AXA's expert report included balcony and exterior photos so could not be dismissed as inadequate; scaffolding was needed solely for the declined external repairs and therefore also excluded; policyholders themselves accepted the balcony floor wear and tear exclusion",
        "Workflow Insight": "An unsubstantiated exposed-position argument for enhanced localised winds is insufficient to establish storm conditions where weather records show sub-storm speeds; where the insurer's expert obtains photographs from the actual claimed damage location, the report cannot be challenged as inadequate; scaffolding costs follow the fate of the underlying repair — if the external damage is excluded, so is the access cost",
        "AI Rule Candidate": "IF weather_records_show_no_storm AND policyholder_claims_enhanced_local_winds AND no_corroborating_evidence THEN storm_q1 = no AND claim_fails; IF external_damage_declined_as_non_storm THEN consequential_scaffolding_cost_also_declined; IF insurer_expert_report_includes_on_site_photographs THEN report_cannot_be_dismissed_as_inadequate",
        "Source PDF": "DRN-4517146.pdf",
    },
    {
        "Case ID": "STORM-022",
        "FOS Decision ID": "DRN-4757581",
        "Insurer Name": "Accredited Insurance (Europe) Ltd",
        "FOS Decision Date": "4 Jul 2024",
        "Claim Type": "Storm damage claim for water ingress through roof declined; insurer found no storm conditions and pre-existing wear and tear from failed previous repairs; claimant subsequently alleged cricket ball impact as alternative cause; both primary and alternative causes rejected",
        "Leak Source": "Roof tiles previously repaired with subsequently failed repairs; broken tiles allowed rainwater to enter roof void; no storm conditions in weeks prior to damage being reported",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "No storm conditions confirmed in weeks prior to damage; approved surveyor found broken tiles with previously failed prior repairs constituting wear and tear and poor workmanship; cricket ball impact claimed as alternative but unsupported by evidence; gradual damage excluded under policy clause 12 and accidental damage cover not included in policy",
        "Evidence Dispute": "Accredited: weather data confirming no storm conditions, approved surveyor report finding previously repaired tiles with failed repairs. Claimant: initially reported storm damage as he was not a roof expert; after seeing surveyor images alleged cricket ball impact from nearby cricket club; said damage happened over time causing ceiling collapse; no evidence provided for either cause",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; Accredited's decline on storm and wear and tear grounds confirmed; cricket ball alternative also rejected as unsupported by evidence and in any event caught by gradual damage exclusion and absence of accidental damage cover",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Storm Q1 failed — no storm conditions in weeks prior to damage; policy general exclusion clause 12: 'Any gradual or maintenance-related loss or damage' including wear and tear and gradual deterioration whether aware or not; accidental damage cover not included in policy; for a claim to succeed the damage must be caused by an insured peril — no insured peril established under any head",
        "Missing Evidence": "Evidence of storm conditions in the relevant period; evidence of cricket ball impact; evidence of a sudden damage event; evidence of roof condition immediately before the damage",
        "Ombudsman Reasoning": "Weather data and Accredited's evidence both confirmed no storm conditions; surveyor found pre-existing tiles with failed prior repairs; cricket ball alternative unsupported by evidence and even if proven would be gradual damage excluded under clause 12; accidental damage cover not in the policy; no insured peril established under any argument advanced",
        "Workflow Insight": "Where a claimant switches claimed peril after initial decline, assess each alternative independently — switching does not reset the burden of proof; a gradual damage exclusion can defeat both the primary storm claim and any subsequent alternative cause if both involve gradual deterioration; confirm whether accidental damage cover is included in the policy before it can be considered as a fallback",
        "AI Rule Candidate": "IF no_storm_conditions AND surveyor_finds_pre_existing_failed_repairs THEN decline_storm_and_wear_and_tear; IF alternative_cause_claimed_post_decline AND alternative_cause_also_gradual AND gradual_damage_excluded THEN alternative_cause_also_fails; IF no_accidental_damage_cover_in_policy THEN accidental_damage_fallback_unavailable",
        "Source PDF": "DRN-4757581.pdf",
    },
    {
        "Case ID": "STORM-023",
        "FOS Decision ID": "DRN-4899211",
        "Insurer Name": "U K Insurance Limited",
        "FOS Decision Date": "23 Aug 2024",
        "Claim Type": "Storm damage claim for water ingress from roof and guttering overflow declined; insurer found no storm conditions and attributed cause to guttering maintenance or defect; internal damage separately covered under accidental damage section; service failures acknowledged with £200 compensation",
        "Leak Source": "Guttering overflow from volume of rainfall allowing water to seep through leadwork at roof edges; no tiles displaced or structurally damaged; battens and felt replacement attributed to maintenance issue or defective guttering system",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Partial",
        "Rejection Reason": "Local weather station data showed no storm conditions on or around the date of loss; surveyor found no evidence of storm damage — no displaced tiles; guttering overflow from rain volume attributed to maintenance issue or possible defective guttering; battens and felt replacement deemed maintenance; internal damage covered separately under accidental damage section",
        "Evidence Dispute": "Insurer: local weather station data, surveyor report finding no storm damage and guttering overflow mechanism. Claimants: Met Office weather warning covering a large geographic area; challenged surveyor conduct; alleged claim mixed up with unrelated claim by agent; felt coerced into low initial internal cash settlement (later increased after second visit)",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold storm claim; decline confirmed; internal damage under accidental damage section not disturbed; UKI's £200 compensation for service failures endorsed as fair and proportionate; no further action required",
        "Compensation Awarded (£)": 200,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Storm defined as violent winds accompanied by heavy rainfall or snow (ABI definition); geographically broad Met Office weather warning is insufficient to establish localised storm conditions at a specific property — local weather station data takes precedence; guttering overflow from rain volume is a maintenance or defective system issue not storm damage; storm and accidental damage sections are independent — external storm claim can be declined while internal accidental damage is accepted under a separate section",
        "Missing Evidence": "Specific local weather data for the claimants' property showing storm-level wind speeds; evidence of displaced tiles or structural storm damage rather than guttering overflow mechanism",
        "Ombudsman Reasoning": "Weather station data for nearest stations showed no storm conditions; Met Office warning covered a wide area and was not specific to claimants' location; no tiles displaced or damaged — guttering backed up allowing water through leadwork, consistent with maintenance or defect not storm; UKI applied separate sections correctly — internal damage was sudden and unintentional so covered under accidental damage; service failures (mixed-up claim, low initial offer, coercion) real but £200 already offered by UKI endorsed as proportionate",
        "Workflow Insight": "A broad regional weather warning does not substitute for local weather station data showing storm-level winds at the specific property; guttering overflow from heavy rain is not storm damage even when rain accompanies otherwise stormy weather; declining storm damage while accepting internal accidental damage under a separate policy section is independently valid; when service failures are acknowledged and compensated before FOS, the FOS-endorsed figure sets the ceiling unless further harm is shown",
        "AI Rule Candidate": "IF weather_warning_covers_large_area AND local_station_shows_no_storm THEN storm_conditions_not_established; IF damage_mechanism_is_guttering_overflow AND no_tiles_displaced THEN storm_q2 = no AND classify_as_maintenance_or_defect; IF storm_section_declined AND accidental_damage_section_applies_independently THEN internal_damage_may_still_be_covered_under_ad_section",
        "Source PDF": "DRN-4899211.pdf",
    },
    {
        "Case ID": "STORM-024",
        "FOS Decision ID": "DRN-5647934",
        "Insurer Name": "AXIS Specialty Europe SE",
        "FOS Decision Date": "18 Jan 2026",
        "Claim Type": "Storm damage claim for boundary fence blown down declined; insurer applied explicit policy exclusion for storm damage to fences under the storm, flood or weight of snow peril despite fences being within the policy definition of buildings",
        "Leak Source": "Boundary fence blown down by storm; no water ingress — structural storm damage to garden boundary structure only",
        "Property Type": "Residential home",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Policy booklet right-hand 'What is not covered' column alongside the storm peril explicitly excludes 'loss or damage to... gates and fences'; exclusion wording clear and unambiguous; storm conditions and fence damage not disputed; the buildings definition including fences does not override the storm-peril-specific fence exclusion",
        "Evidence Dispute": "Insurer: policy booklet exclusion column c) under storm peril explicitly naming fences; FOS confirmed wording unambiguous. Claimant: fences within buildings definition, storm definition met (>47mph), argued exclusion wording was ambiguous, referenced Clause 4 Storm Exclusion Clause as most relevant text — however Clause 4 was not listed in the policy schedule and therefore inapplicable",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; AXIS's decline confirmed as fair; storm-peril exclusion for fences is clear and unambiguous; Clause 4 endorsement not listed in policy schedule so inapplicable to this policy; no award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "A specific peril-level exclusion for fences listed in the storm section 'What is not covered' column overrides the general buildings definition that includes fences; fences within the buildings definition does not mean storm damage to fences is covered; only endorsements listed in the policy schedule apply to that policy — a clause printed in the booklet but absent from the schedule is inapplicable; storm damage to fences is a standard industry exclusion and need not also appear in the general exclusions section",
        "Missing Evidence": "Not applicable — purely a policy wording dispute; no factual evidence gaps",
        "Ombudsman Reasoning": "Storm conditions and fence damage not in dispute; policy booklet used a box format with 'What is covered' left and 'What is not covered' right; right-hand column alongside storm peril listed three specific exclusions including 'gates and fences'; this is separate from and overrides the general buildings definition; Clause 4 in the booklet headed 'Storm Exclusion Clause' not listed in the policy schedule as an applicable endorsement and therefore does not apply; Mrs R's ambiguity arguments rejected — three exclusions in the right-hand column are separate items each read independently",
        "Workflow Insight": "Check both the storm peril's own exclusion column and the general exclusions separately — a specific peril-level exclusion for fences is valid regardless of the buildings definition; verify which endorsements are listed in the policy schedule before applying or arguing a booklet clause; explicitly cite the exact policy wording of the fence exclusion in decline letters to pre-empt ambiguity arguments",
        "AI Rule Candidate": "IF claimed_item IN ['fence', 'gate'] AND policy_storm_peril_exclusion_includes_fences THEN decline_storm_claim_regardless_of_buildings_definition; IF endorsement_in_booklet AND endorsement_not_listed_in_policy_schedule THEN endorsement_inapplicable; IF buildings_definition_includes_item AND storm_peril_exclusion_also_names_item THEN storm_peril_exclusion_prevails",
        "Source PDF": "DRN-5647934.pdf",
    },
    {
        "Case ID": "STORM-025",
        "FOS Decision ID": "DRN-6075693",
        "Insurer Name": "AXA Insurance UK Plc",
        "FOS Decision Date": "20 Feb 2026",
        "Claim Type": "Storm damage claim for structural damage to flat (wall head gutter and cornice) declined by current insurer AXA; damage established by prior Ombudsman decision and structural engineer report as gradual and cumulative from successive storms in 2021-2022, predating AXA policy start date of 31 December 2022",
        "Leak Source": "Wall head gutter and cornice — cumulative structural deterioration from successive storms from 2021 onwards; storms progressively loosened masonry bond allowing water penetration and further accelerated deterioration; damage gradual not sudden",
        "Property Type": "Leasehold flat",
        "Dispute Type": "Pre-Inception Damage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "AXA policy incepted 31 December 2022; prior Ombudsman final decision and chartered structural engineer report established damage was gradual and cumulative from successive storms in 2021-2022, all predating AXA's policy period; AXA not required to respond to claims for events occurring before it was on risk; claimant's discovery of the cause in January 2023 does not shift liability to AXA",
        "Evidence Dispute": "AXA: prior Ombudsman final decision against previous insurer F; chartered structural engineer report (D) describing 'cumulative effects of damage from successive storms' from 2021 with progressive masonry loosening; claimant's own statements referencing November 2022 storm and January 2023 catastrophic collapse. Claimant: argued damage could have occurred due to a storm in January 2023 within AXA's policy period; disputed the gradual nature; said he did not know he had cause to claim until January 2023",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; AXA's decline confirmed as fair; prior Ombudsman decision and structural engineer report established damage predated AXA policy period; claimant's ignorance of the cause is irrelevant to which insurer is on risk; no award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Insurer on risk not required to consider claims for events occurring before its policy period started; storm cover requires sudden not gradual damage; cumulative effects of damage from successive storms over multiple years is gradual damage not a sudden storm event; policyholder's first awareness of the cause does not determine which insurer is liable; a prior Ombudsman final decision establishing the damage chronology is highly persuasive in subsequent related complaints",
        "Missing Evidence": "Evidence of a distinct sudden storm damaging event within AXA's policy period after 31 December 2022; evidence contradicting the structural engineer's finding of cumulative gradual damage commencing in 2021",
        "Ombudsman Reasoning": "Prior Ombudsman established damage was gradual from 2021 storms under previous insurer F; chartered structural engineer D described 'cumulative effects of damage from successive storms' in 2022 with progressive masonry loosening — not a sudden event; claimant's own statements in earlier complaints attributed damage to November 2022 storm predating AXA's cover; AXA policy only covers events after 31 December 2022; policyholder's ignorance of the cause does not shift risk to AXA; no evidence of a new distinct storm damaging event within AXA's period",
        "Workflow Insight": "Where a prior Ombudsman decision establishes the damage chronology for the same property damage, it is highly persuasive in subsequent related complaints and binds the factual finding; when consecutive policies are in play and damage is cumulative, liability rests with the insurer on risk when the damage originated — not when it became apparent to the policyholder; first-awareness date is irrelevant to which insurer is on risk; cumulative successive-storm damage is not a sudden storm event and falls outside the storm peril",
        "AI Rule Candidate": "IF damage_cause_predates_current_insurer_policy_start_date THEN current_insurer_not_liable REGARDLESS_OF_DISCOVERY_DATE; IF prior_ombudsman_decision_establishes_gradual_damage_and_pre_inception THEN current_insurer_may_rely_on_that_finding; IF damage_is_cumulative_successive_storms AND commenced_before_policy_inception THEN storm_peril_not_triggered_under_current_policy",
        "Source PDF": "DRN-6075693.pdf",
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

    for case in NEW_CASES:
        _validate(case)

    repo_root = os.path.normpath(os.path.join(os.path.dirname(__file__), ".."))
    xlsx_path = os.path.join(
        repo_root, "knowledge", "case-databases", "Storm_Case_Database.xlsx"
    )

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active

    live_headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    if live_headers != COLUMNS:
        mismatch = [
            (i + 1, live_headers[i] if i < len(live_headers) else "—", COLUMNS[i])
            for i in range(max(len(live_headers), len(COLUMNS)))
            if i >= len(live_headers) or i >= len(COLUMNS) or live_headers[i] != COLUMNS[i]
        ]
        raise RuntimeError(
            "Live workbook columns do not match COLUMNS definition.\n"
            "Mismatches (col, live, expected):\n"
            + "\n".join(f"  {c}: '{l}' vs '{e}'" for c, l, e in mismatch)
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

    last_row  = ws.max_row
    last_case = ws.cell(row=last_row, column=1).value
    total_data = last_row - 1

    print(f"Appended {len(NEW_CASES)} case(s) to {xlsx_path}")
    print(f"Total data rows : {total_data}")
    print(f"Last Case ID    : {last_case}")


if __name__ == "__main__":
    main()
