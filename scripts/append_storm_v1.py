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
        "Case ID": "STORM-001",
        "FOS Decision ID": "DRN-1207086",
        "Insurer Name": "Aviva Insurance Limited",
        "FOS Decision Date": "16 Sep 2020",
        "Claim Type": "Roof leak claimed as storm damage; drone inspection revealed slipped, chipped and displaced tiles consistent with long-term weathering; significant prior repair evidence; no evidence of a one-off storm event",
        "Leak Source": "Roof water ingress — slipped, chipped and displaced tiles with evidence of prior repairs; damage pattern consistent with gradual weathering over time, not a single storm event",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Drone inspection showed slipped, chipped and displaced tiles consistent with general weathering over time; significant evidence of previous repair attempts indicating ongoing leaks; no evidence of a one-off storm event",
        "Evidence Dispute": "Aviva relied on drone inspection evidence showing slipped, chipped and displaced tiles and prior repair evidence; Mr E could not pinpoint an exact date of damage; no independent expert evidence produced to counter the drone findings; FOS accepted Aviva's stance",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; Aviva's decision to decline confirmed as reasonable; no award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "All three storm questions must be answered yes; damage inconsistent with a one-off storm event (Q2 answered no) means the claim fails regardless of whether storm conditions existed; drone inspection evidence showing progressive tile displacement and prior repairs is highly persuasive",
        "Missing Evidence": "Exact date of damage; independent expert report challenging drone inspection findings; evidence ruling out gradual weathering as the cause of tile displacement",
        "Ombudsman Reasoning": "Drone evidence showed slipped, chipped and displaced tiles consistent with general weathering over time, not a one-off storm event; significant prior repair attempts corroborated the ongoing nature of the problem; Q2 answered no means Q3 also no; Aviva acted reasonably",
        "Workflow Insight": "Drone inspections produce highly persuasive objective evidence; where tiles are slipped, chipped and displaced across a roof with evidence of prior repairs, the pattern indicates gradual weathering not storm damage; an insurer arranging a drone inspection in response to a disputed claim demonstrates robust evidential practice",
        "AI Rule Candidate": "IF drone_inspection_shows_slipped_chipped_displaced_tiles AND prior_repair_evidence_present THEN damage_pattern = gradual_weathering AND storm_claim = NOT established; IF no_exact_date_of_damage AND no_independent_expert_evidence THEN insurer_drone_report = decisive",
        "Source PDF": "DRN-1207086.pdf",
    },
    {
        "Case ID": "STORM-002",
        "FOS Decision ID": "DRN-1223113",
        "Insurer Name": "Ageas Insurance Limited",
        "FOS Decision Date": "16 Sep 2020",
        "Claim Type": "Loose tiles on a well-maintained roof; Ageas declined citing its 55mph storm definition not met (51mph recorded); FOS upheld — 55mph threshold unreasonably high; contractor confirmed roof in good condition and high winds as the only logical cause",
        "Leak Source": "Wind-loosened roof tiles on a well-maintained roof; wind speed of 51mph recorded at the time of damage",
        "Property Type": "Residential home",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Ageas relied on its policy definition requiring wind speeds in excess of 55mph; weather records showed only 51mph; Ageas concluded storm conditions not met",
        "Evidence Dispute": "Ageas relied on its 55mph policy definition and recorded 51mph wind speeds; Ms H's contractor stated roof was in good condition and high winds were the only logical reason for tiles to come loose; FOS found 55mph threshold above what the industry and FOS consider fair and reasonable; 51mph treated as qualifying storm conditions",
        "Outcome Category": "Upheld",
        "Outcome": "Ageas Insurance Limited directed to pay Ms H's storm damage repair costs",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "An insurer's internal policy definition of storm (e.g. 55mph) is not automatically decisive; FOS applies a reasonableness test and will override an insurer's threshold if it is above what the service and wider insurance industry consider fair; 51mph wind speeds are sufficient to constitute storm conditions for FOS purposes; where a roof is well maintained and a contractor confirms high winds as the only logical cause, all three storm questions are answered yes",
        "Missing Evidence": "Independent weather data confirmation at the specific property postcode; surveyor report assessing pre-storm roof condition",
        "Ombudsman Reasoning": "55mph threshold is above what FOS and the wider industry consider fair; 51mph qualifies as storm conditions; tiles coming loose on a well-maintained roof is consistent with storm damage (Q2 yes); contractor confirmed high winds as the only logical reason — best evidence on file (Q3 yes); all three questions answered yes; Ageas acted unreasonably",
        "Workflow Insight": "Where an insurer relies on an internal storm definition above approximately 50mph to decline, that threshold is likely to be overridden by FOS; always check whether the property was well maintained — a well-maintained roof with loose tiles in high winds strongly supports storm causation; contractor opinion confirming storm as the only logical cause is persuasive where the roof is documented as in good condition",
        "AI Rule Candidate": "IF insurer_relies_on_55mph_storm_definition AND wind_speed_was_51mph THEN threshold_unreasonable AND storm_conditions = established; IF roof_well_maintained AND contractor_confirms_wind_only_logical_cause THEN all_three_storm_questions = yes AND claim = valid",
        "Source PDF": "DRN-1223113.pdf",
    },
    {
        "Case ID": "STORM-003",
        "FOS Decision ID": "DRN-1586732",
        "Insurer Name": "esure Insurance Limited",
        "FOS Decision Date": "8 Jul 2020",
        "Claim Type": "Water damage to roof and bedroom ceiling following heavy rain; storm conditions acknowledged; esure's surveyor found no missing or dislodged slates; internal mould growth and prior ceiling repairs indicated gradual ingress rather than a single storm event",
        "Leak Source": "Roof water ingress — no missing or dislodged slates identified; surveyor attributed roof damage to unknown defect and internal damage to rainwater ingress over a period of time; mould growth and repainted ceiling area indicated gradual ingress",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Surveyor found no missing or dislodged slates; roof damage attributed to unknown defect not storm; internal damage attributed to rainwater ingress over a period of time; mould growth and evidence of prior ceiling repairs indicated gradual ingress; accidental damage also declined as damage was not a single event",
        "Evidence Dispute": "Esure relied on surveyor's report with photographs and moisture readings finding no storm-consistent external damage and gradual internal ingress; Mr and Mrs D had no independent expert evidence to counter the surveyor; Mrs D's description of ceiling collapse in heavy rain was plausible but could not override the surveyor's objective physical findings",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; esure's decline under both storm and accidental damage sections confirmed as fair; no award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Where a surveyor conducts a thorough inspection with photographs and moisture readings and finds no missing or dislodged slates and attributes internal damage to gradual ingress, that report is highly persuasive in the absence of independent counter-evidence; mould growth and prior ceiling repairs are physical indicators of gradual water ingress not a single storm event; heavy rainfall highlighting an existing roof issue is not the same as storm causing damage",
        "Missing Evidence": "Independent expert report challenging the surveyor's findings; photographs taken immediately after the alleged storm event before any repair work; evidence of roof condition prior to the claim",
        "Ombudsman Reasoning": "Storm conditions were acknowledged (Q1 yes); no missing or dislodged slates meant damage not consistent with storm (Q2 no); surveyor found internal damage was rainwater ingress over time — mould growth and repainted ceiling area corroborated this; heavy rainfall exposed existing issues; no independent evidence to counter the surveyor; accidental damage also inapplicable as policy excludes gradual damage",
        "Workflow Insight": "The absence of missing or dislodged slates is a key negative indicator for storm claims; internal mould growth and areas of prior ceiling repair are strong physical evidence of gradual ingress; a comprehensive surveyor report with photographs and moisture readings is highly persuasive where the policyholder cannot produce independent expert counter-evidence",
        "AI Rule Candidate": "IF surveyor_finds_no_missing_or_dislodged_slates AND internal_mould_growth AND prior_ceiling_repairs_visible THEN damage = gradual_ingress NOT storm; IF no_independent_expert_evidence AND insurer_has_surveyor_report_with_photos_and_moisture_readings THEN surveyor_report = decisive",
        "Source PDF": "DRN-1586732.pdf",
    },
    {
        "Case ID": "STORM-004",
        "FOS Decision ID": "DRN-2053943",
        "Insurer Name": "Lloyds Bank General Insurance Limited",
        "FOS Decision Date": "11 Sep 2020",
        "Claim Type": "Leaking skylight causing ceiling, wallpaper and blind damage; skylight was old and replaced by policyholder as maintenance; damage occurred gradually over months; no storm conditions in records; no accidental damage cover; gradually operating cause exclusion applied",
        "Leak Source": "Old skylight requiring replacement — leaking due to age and wear; damage to ceiling plaster, wallpaper and blind accumulated over months of gradual ingress",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "No storm conditions in records for the area in the preceding months; skylight leaked because it was old and needed replacing (maintenance not an insured event); damage described by Mrs E herself as occurring over months; no accidental damage cover on the policy; gradually operating cause exclusion applied",
        "Evidence Dispute": "Lloyds checked weather records and found no evidence of storm conditions; Mrs E's own description confirmed the skylight had been leaking gradually over months not a one-off incident; policy excluded damage by gradually operating causes; even if accidental damage cover had existed, policy excluded water entering other than by storm or flood",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; Lloyds' decline confirmed as fair; Lloyds' offer to reconsider on receipt of a builder's report confirmed as reasonable; no award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Damage that appears sudden to a policyholder may still have been caused gradually — the physical cause not the moment of discovery determines coverage; where the policyholder's own account confirms gradual damage, the insurer is entitled to rely on the gradually operating cause exclusion; replacing an old skylight is a maintenance obligation not a covered event; no storm conditions in records combined with the policyholder's own description of gradual ingress means the claim is correctly declined",
        "Missing Evidence": "Independent builder's report attributing damage to a specific insured event rather than gradual ingress; weather data for specific dates the policyholder believed storms occurred",
        "Ombudsman Reasoning": "Mrs E herself described the damage as happening over the last couple of months via the leaking skylight; she acknowledged the skylight needed replacing due to age; no storm conditions in records; suddenly becoming aware of damage is not the same as the damage happening suddenly; Lloyds correctly applied the gradually operating cause exclusion; no accidental damage cover existed; even with it, water entry other than by storm or flood is excluded",
        "Workflow Insight": "A policyholder who pays to replace a worn-out component before claiming for consequential damage has effectively acknowledged a maintenance cause; damage appearing suddenly is not the same as damage occurring suddenly — always establish the timeline of the underlying cause; the absence of storm records combined with the policyholder's own account of gradual ingress is a strong dual indicator for decline",
        "AI Rule Candidate": "IF policyholder_self_described_damage_as_gradual AND no_storm_conditions_in_records THEN gradually_operating_cause = applies AND storm_claim = fails; IF policyholder_replaced_component_due_to_age_before_claiming THEN maintenance_cause = established AND coverage = unlikely",
        "Source PDF": "DRN-2053943.pdf",
    },
    {
        "Case ID": "STORM-005",
        "FOS Decision ID": "DRN-2556262",
        "Insurer Name": "Lloyds Bank General Insurance Limited",
        "FOS Decision Date": "23 Mar 2021",
        "Claim Type": "Water ingress into attic following heavy rain; storm conditions acknowledged; flat roof section showing lifted and detached covering; worn material at window-roof junction; insurer's report attributed damage to pre-existing deterioration revealed by heavy rainfall",
        "Leak Source": "Flat roof section with lifting and detached covering and worn, cracked material at window-to-roof junction — pre-existing deterioration revealed by heavy rainfall rather than caused by storm",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Insurer's report found roof materials had deteriorated; flat roof edge covering lifting and coming away; material at window-roof junction worn and cracked; heavy rainfall highlighted an existing problem rather than causing it; builder's opinion insufficient to override insurer's detailed inspection report",
        "Evidence Dispute": "Lloyds relied on its inspection report and photographs showing deteriorated flat roof covering and worn window-junction material; Mrs L's builder said the roof was fine and only a corner was damaged; FOS preferred Lloyds' report — builder's opinion lacked sufficient causation detail and his argument that no prior damp marks meant sudden damage was specifically rejected: water may not have been entering before the storm exposed the weakness",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; Lloyds confirmed as acting reasonably in declining; no premium refund; no award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Storm conditions and damage superficially consistent with storm (Q1 and Q2 yes) are not sufficient if the storm was not the dominant cause (Q3); where an insurer's inspection report with photographs shows lifting flat roof material and deteriorated window-junction sealing, this is compelling evidence the storm revealed rather than caused the damage; a builder's opinion without detailed causation analysis does not override an insurer's detailed inspection report; the absence of prior water ingress does not prove storm causation where pre-existing deterioration exists",
        "Missing Evidence": "Independent expert report on the roof's pre-storm condition; detailed causation analysis from Mrs L's builder; evidence that the flat roof and window junction were in good condition before the heavy rain",
        "Ombudsman Reasoning": "Storm conditions present (Q1 yes); water ingress from roof consistent with storm damage (Q2 yes); dominant cause question failed (Q3 no) because flat roof covering was already lifting and detaching and window-junction material was already worn and cracked — heavy rainfall exposed rather than caused these weaknesses; builder's opinion lacked sufficient detail; absence of prior damp marks does not prove storm causation where pre-existing deterioration exists",
        "Workflow Insight": "A flat roof with lifting material or a deteriorated window-to-roof junction is a strong indicator that heavy rainfall revealed a pre-existing weakness; the builder's argument that no prior damp marks means a sudden storm event is not persuasive — water may not have entered before simply because conditions had not yet pushed the weakness to failure; always obtain detailed causation analysis from any contractor rather than repair notes only",
        "AI Rule Candidate": "IF flat_roof_covering_lifting_or_detaching AND window_roof_junction_worn_or_cracked THEN storm_revealed_preexisting_weakness AND dominant_cause_not_storm; IF builder_opinion_lacks_causation_analysis AND insurer_has_inspection_report_with_photos THEN insurer_report = more_persuasive",
        "Source PDF": "DRN-2556262.pdf",
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
