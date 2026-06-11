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
        "Case ID": "STORM-006",
        "FOS Decision ID": "DRN-2560515",
        "Insurer Name": "Royal & Sun Alliance Insurance Plc",
        "FOS Decision Date": "8 Jun 2021",
        "Claim Type": "Caravan roof seal failure claimed as storm damage; three repairers all attributed cause to general weathering or design fault; no visible storm-consistent damage such as lifted or broken roof panels; Q2 failed",
        "Leak Source": "Caravan roof ridge seal and bolt hole failure — seals attributed to general weathering or design fault by all three repairers; no lifted or broken roof panels consistent with storm",
        "Property Type": "Static caravan",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Three repairers assessed the caravan — one said seals failed due to general weathering; RSA's specialist initially noted storm or design fault but later concluded design fault as there was no visible storm damage to the roof; no evidence of lifted or broken roof panels",
        "Evidence Dispute": "RSA relied on three concurrent repairer assessments all attributing cause to weathering or design fault; Mrs L argued that less water ingress would be expected from a design fault, but RSA's file showed the site owner confirmed water damage was getting worse; no independent expert evidence submitted to challenge the repairer findings",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; RSA's decision to decline confirmed as fair and reasonable; RSA not at fault for delays in obtaining quotes as site owner took three months; no evidence RSA pressured its specialist into changing position; no award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "All three storm questions must be answered yes; Q2 failed — no evidence of damage consistent with a storm such as lifted or broken roof panels; three concurrent repairer opinions all attributing cause to weathering or design fault is highly persuasive in the absence of independent counter-evidence; an insurer is entitled to rely on concurrent expert opinions to decline",
        "Missing Evidence": "Independent expert report challenging the three repairer assessments; photographs showing storm-consistent structural damage such as lifted or broken roof panels; evidence ruling out general weathering and design fault as the cause",
        "Ombudsman Reasoning": "Storm conditions present (Q1 yes); no evidence of damage consistent with storm — no lifted or broken roof panels (Q2 no); all three repairers attributed cause to general weathering or design fault; RSA's specialist changed its initial equivocal position to design fault after confirming no visible storm damage; Mrs L's argument that less ingress would be expected from a design fault was contradicted by the site owner confirming damage was worsening; no evidence RSA pressured its contractor to change position; Q2 answered no means Q3 not reached",
        "Workflow Insight": "Where three independent repairers all attribute damage to weathering or design fault, an insurer can rely on this concurrence to decline without storm causation being established; absence of lifted or broken structural panels is a key negative indicator for Q2; a policyholder's subjective expectation about the degree of ingress from a design fault does not override concurrent expert physical assessments",
        "AI Rule Candidate": "IF three_or_more_repairers_attribute_damage_to_weathering_or_design_fault AND no_lifted_or_broken_roof_panels THEN storm_q2 = no AND claim = not_established; IF insurer_has_concurrent_expert_opinions AND policyholder_has_no_independent_expert_evidence THEN insurer_evidence = decisive",
        "Source PDF": "DRN-2560515.pdf",
    },
    {
        "Case ID": "STORM-007",
        "FOS Decision ID": "DRN-2737383",
        "Insurer Name": "Liverpool Victoria Insurance Company Limited",
        "FOS Decision Date": "7 May 2021",
        "Claim Type": "Storm damage to residential roof in February 2020; LV declined external damage (tiles, hips, felt, lead flashing) citing wear and tear but based its exclusion on a different roof area from that being claimed for; repair contractor confirmed good overall roof condition and storm causation; FOS upheld",
        "Leak Source": "Roof tiles, hips and felt damaged with lead flashing lifted — contractor confirmed lifted tiles and ripped felt caused by gale force winds in the claimed area; LV inspected a different roof section and applied exclusion incorrectly to uninspected area",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Partial",
        "Rejection Reason": "LV's contractor found age-related deterioration, historic repairs and lead flashing used as a substitute for hip tiles in the inspected areas; LV relied on these findings and Google Street View images to conclude wear and tear and poor workmanship; declined external damage while accepting internal damage and contents",
        "Evidence Dispute": "LV relied on ground-level contractor report and Google Street View images showing deterioration on inspected areas of the roof; Mr H's contractor confirmed the claimed area showed good overall roof condition (lathes and underfelt very good) with lifted tiles and ripped felt consistent with storm damage from gale force winds; LV's own agent who inspected internal damage confirmed water ingress was fresh and not longstanding; repair invoice included chimney works not connected to storm damage, complicating settlement quantum",
        "Outcome Category": "Upheld",
        "Outcome": "Upheld — LV must deal with Mr H's storm damage claim for the affected roof areas he claimed for and had repaired; add 8% simple interest to any cash settlement from date of payment to date of settlement; pay Mr H a further £100 compensation in addition to the £100 already paid for delays",
        "Compensation Awarded (£)": 100,
        "Is Core Case": "Yes",
        "Key Policy Clause": "An insurer must inspect the specific area of roof being claimed for before applying a wear and tear or poor workmanship exclusion; relying on the condition of a different and unrelated roof section to decline a storm claim for an uninspected area is not fair or reasonable; where the insurer's own agent confirms water ingress was fresh and not longstanding, this undermines a maintenance or wear and tear defence; an insurer that fails to inspect the damaged area cannot demonstrate the exclusion applies to that area",
        "Missing Evidence": "Itemised repair invoice breaking down storm-related costs from non-storm elements (chimney and other non-storm works); physical inspection by LV of the specific roof area where storm damage was claimed",
        "Ombudsman Reasoning": "Storm conditions agreed (Q1 yes); damage type — lifted tiles, ripped felt, lifted lead flashing and resultant internal water ingress — consistent with storm damage (Q2 yes); LV inspected only from ground level and did not view the claimed area, instead relying on condition of a different roof section to apply the exclusion — exclusion not shown to apply to the claimed area (Q3 — LV failed to discharge burden); Mr H's contractor confirmed good overall condition with storm-consistent damage in the claimed area; LV's own agent confirmed ingress was fresh not longstanding, inconsistent with longstanding maintenance failure",
        "Workflow Insight": "Inspections must cover the specific area of damage being claimed for; applying a wear and tear exclusion to an uninspected roof section based on the condition of a different area is not fair and will not be upheld by FOS; where a contractor confirms good overall roof condition and storm-consistent damage in the claimed area, and the insurer's own agent reports fresh ingress, storm causation is established for that area; repair invoices covering mixed storm and non-storm works must be itemised before settlement",
        "AI Rule Candidate": "IF insurer_inspected_different_roof_area AND applied_exclusion_to_uninspected_claimed_area THEN exclusion_not_shown_to_apply AND outright_decline_not_fair; IF contractor_confirms_good_roof_condition_in_claimed_area AND insurers_own_agent_confirms_fresh_ingress THEN storm_causation_established_for_claimed_area; IF repair_invoice_contains_non_storm_works THEN require_itemised_breakdown_before_settlement",
        "Source PDF": "DRN-2737383.pdf",
    },
    {
        "Case ID": "STORM-008",
        "FOS Decision ID": "DRN-2788212",
        "Insurer Name": "Lloyds Bank General Insurance Limited",
        "FOS Decision Date": "11 Jun 2021",
        "Claim Type": "Water ingress through roof during storm; policyholder carried out urgent repairs and builder disposed of damaged materials before insurer could inspect; no photographic evidence of pre-repair damage; builder statement alone insufficient to establish Q2 or Q3; claim declined for failure to substantiate loss",
        "Leak Source": "Roof water ingress — nature of storm damage unknown as repairs were completed and damaged materials disposed of before inspection; storm conditions confirmed but damage pattern could not be assessed",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Repairs completed and damaged materials disposed of before Lloyds could inspect; photos provided only showed completed repairs not pre-repair damage; builder gave conflicting reasons when Lloyds spoke to him; Lloyds could not validate Q2 or Q3 without physical evidence",
        "Evidence Dispute": "Mr and Mrs H provided photos of completed roof repairs and a builder email stating the roof was in good condition and damage was wind-caused; Lloyds argued photos showed only completed work, builder gave conflicting reasons when contacted, and damaged sections had been disposed of; FOS agreed builder's statement alone was insufficient and that Q2 and Q3 could not be answered without physical evidence of the pre-repair damage",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; Lloyds acted fairly and within policy terms in declining; policyholder failed to provide sufficient evidence to establish storm causation; no award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Policyholder must not carry out non-emergency repairs or dispose of damaged items before the insurer has had a chance to inspect; storm conditions alone (Q1) are insufficient without evidence to establish Q2 and Q3; a builder's statement without photographic or physical evidence of pre-repair damage does not prove storm causation; a well-maintained roof should withstand all but the most severe weather and Q2 requires evidence to assess this",
        "Missing Evidence": "Photographs of the damaged roof taken before repair commenced; damaged roofing materials (tiles, felt, etc.) retained for inspection; independent inspection of the roof before repair; explanation of the conflicting reasons given by the builder to Lloyds",
        "Ombudsman Reasoning": "Storm conditions present (Q1 yes); Q2 and Q3 could not be answered as repairs were completed and materials disposed of before Lloyds could inspect — no pre-repair photographs and no physical evidence remained; builder gave conflicting reasons to Lloyds; policy required policyholder to preserve evidence and notify insurer before non-emergency repairs; builder could have photographed the damage before starting work; Lloyds acted fairly in declining on the evidence available",
        "Workflow Insight": "Policyholders who carry out urgent storm repairs must photograph all damage and retain damaged materials before work commences; a builder carrying out repairs without documenting the pre-repair state removes the insurer's ability to validate Q2 and Q3; insurers should communicate evidence-preservation obligations clearly at first notification; a builder's verbal or written statement without corroborating physical evidence is insufficient to establish storm causation",
        "AI Rule Candidate": "IF repairs_completed_before_inspection AND damaged_materials_disposed_of AND no_pre_repair_photographs THEN claim_evidence_insufficient AND Q2_Q3_cannot_be_answered AND decline_reasonable; IF only_evidence_is_builder_statement AND no_physical_evidence_of_storm_damage THEN storm_causation_not_established",
        "Source PDF": "DRN-2788212.pdf",
    },
    {
        "Case ID": "STORM-009",
        "FOS Decision ID": "DRN-2877529",
        "Insurer Name": "Lloyds Bank General Insurance Limited",
        "FOS Decision Date": "21 Jul 2021",
        "Claim Type": "Chimney stack water ingress following storms Clara and Denis in February 2020; surveyor found no displaced tiles or render and attributed damage to natural breakdown of flaunching; contractor estimate made no reference to storm damage; external chimney claim declined, internal bedroom damage accepted; FOS not upheld",
        "Leak Source": "Chimney stack — deterioration of flaunching with no displaced tiles or render; chimney caps undisturbed; water entered through everyday weathering not a one-off storm event",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Partial",
        "Rejection Reason": "Surveyor found no visible storm damage to the chimney — no displaced tiles or render, caps still in place, nothing disturbed by wind; flaunching deterioration attributed to natural breakdown of materials; contractor estimate for chimney repairs made no reference to storm damage",
        "Evidence Dispute": "Lloyds relied on surveyor report and photos showing no storm-consistent physical indicators on the chimney and a contractor estimate not referencing storm damage; Mrs C said the damp in the bedroom wall worsened considerably after storms Clara and Denis and that she had always maintained the chimney; FOS agreed with Lloyds that absence of physical storm indicators was decisive regardless of the timing of observed damp worsening",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; Lloyds correctly declined external chimney damage as not consistent with storm; Lloyds had already accepted internal bedroom damage; no award on the chimney claim",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Q2 requires damage to be consistent with storm-typical damage; where a surveyor finds no displaced tiles or render and chimney caps remain in place and undisturbed, Q2 is answered no; deterioration of flaunching is a maintenance issue not a storm peril; a policyholder observing damp worsen after storms does not establish storm causation where physical evidence shows no storm-consistent structural disturbance; contractor estimates that do not reference storm damage are a significant negative indicator",
        "Missing Evidence": "Independent expert report attributing chimney damage to storm; photographs showing displaced tiles, render failure or chimney cap disturbance; evidence of chimney condition immediately before the storms",
        "Ombudsman Reasoning": "Storm conditions present (Q1 yes); surveyor found no displaced tiles or render, chimney caps still in place and nothing disturbed by wind (Q2 no); contractor estimate gave no reference to storm damage; FOS agreed external damage was not consistent with storm; Mrs C's observation that damp worsened after the storms did not override physical evidence showing no storm indicators; Lloyds correctly applied the three-question framework and declined the external chimney claim; internal bedroom damage already accepted",
        "Workflow Insight": "Worsening damp after storms is not conclusive of storm causation — always check for physical storm indicators such as displaced tiles, disturbed render or lifted chimney caps; where a surveyor finds flaunching deterioration with no structural disturbance, Q2 is answered no; a contractor estimate that does not reference storm damage is a significant negative indicator; chimneys are a common source of gradual water ingress and require clear physical storm evidence to distinguish from ongoing deterioration",
        "AI Rule Candidate": "IF surveyor_finds_no_displaced_tiles_or_render AND chimney_caps_undisturbed THEN storm_q2 = no AND damage = weathering_not_storm; IF contractor_estimate_contains_no_storm_damage_reference AND surveyor_attributes_damage_to_natural_breakdown THEN insurer_decline_of_chimney = reasonable; IF policyholder_reports_damp_worsening_after_storm AND no_physical_storm_indicators THEN temporal_correlation_insufficient_to_establish_causation",
        "Source PDF": "DRN-2877529.pdf",
    },
    {
        "Case ID": "STORM-010",
        "FOS Decision ID": "DRN-2926734",
        "Insurer Name": "Fairmead Insurance Limited",
        "FOS Decision Date": "12 Aug 2021",
        "Claim Type": "Water ingress through roof in October 2020; Fairmead outright declined citing no storm conditions on the precise date of loss; FOS found 49mph gusts on 25 September 2020 (within policy's 47mph definition) shortly before the damage; property in exposed coastal location; insurer directed to reconsider with full inspection",
        "Leak Source": "Roof water ingress — lifted tiles visible in policyholder photographs; dominant cause not determined as Fairmead declined without inspection",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Fairmead said weather conditions on 3 October 2020 (reported date of loss) did not reach storm force as defined in the policy (gusts of at least 47mph); declined outright on this basis without inspecting the property or assessing Q2 or Q3",
        "Evidence Dispute": "Fairmead relied on weather records for the specific date of loss showing sub-47mph winds; Ms B provided photos showing lifted tiles and internal water damage; FOS reviewed broader weather records and found 49mph gusts on 25 September 2020 shortly before the loss meeting the policy storm definition; Fairmead's own claim notes acknowledged the property was in an exposed coastal location which could mean actual conditions exceeded nearest weather station readings",
        "Outcome Category": "Upheld",
        "Outcome": "Upheld — Fairmead directed to reconsider Ms B's storm damage claim in line with remaining policy terms including arranging a physical inspection; pay Ms B £100 compensation for distress caused by incorrect outright decline; if Ms B remains dissatisfied after reconsideration she may bring a new complaint",
        "Compensation Awarded (£)": 100,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Weather records must be checked over the period leading up to the date of loss, not only on the precise reported date; a policy storm definition of 47mph gusts is met by 49mph gusts recorded shortly before the loss; an exposed coastal location means actual conditions at the property may exceed nearest weather station readings; outright declining without physical inspection when storm conditions did occur in the relevant period is not fair or reasonable; an insurer must assess all three questions before declining",
        "Missing Evidence": "Physical inspection of roof and internal damage by Fairmead; independent expert report on whether storm or gradual deterioration was the dominant cause; weather data specific to the exposed coastal location rather than the nearest weather station",
        "Ombudsman Reasoning": "Policy defined storm as 47mph gusts; weather records showed 49mph on 25 September 2020 shortly before the loss — Q1 met on the broader weather record; Ms B's photos showed lifted tiles and internal water damage consistent with storm damage (Q2 provisionally yes); Fairmead declined outright without inspecting so Q3 could not be assessed — dominant cause undetermined; Fairmead incorrectly assessed storm conditions only on the precise date of loss rather than the period leading up to it; exposed coastal location acknowledged by Fairmead itself as a relevant factor; directed to reconsider with full inspection and pay £100 compensation",
        "Workflow Insight": "Weather records must be reviewed over the relevant period before the date of loss, not just the precise date reported — storm conditions occurring days before visible damage can still satisfy Q1; an insurer that declines solely on weather records without inspecting the property cannot assess Q2 or Q3 and risks an incorrect outright decline; an exposed coastal location is a material factor requiring consideration when comparing weather station data to actual property conditions; always inspect before declining when storm conditions occurred in the relevant period",
        "AI Rule Candidate": "IF weather_records_show_storm_conditions_in_period_before_date_of_loss AND insurer_checked_only_exact_date_reported THEN storm_q1_assessment_incomplete AND outright_decline_not_justified; IF insurer_declined_without_inspection AND storm_conditions_occurred_in_relevant_period THEN insurer_must_reconsider_and_inspect; IF property_in_exposed_location AND weather_station_readings_near_threshold THEN actual_conditions_may_exceed_station_readings",
        "Source PDF": "DRN-2926734.pdf",
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
