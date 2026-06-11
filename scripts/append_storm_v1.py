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
        "Case ID": "STORM-011",
        "FOS Decision ID": "DRN-2926772",
        "Insurer Name": "National Farmers' Union Mutual Insurance Society Limited",
        "FOS Decision Date": "26 Oct 2021",
        "Claim Type": "Second storm damage claim for porch roof water ingress in August 2020; NFU's loss adjuster found no storm indicators — only chipped tiles, loose mortar in lead valleys and gapping lead flashing attributed to wear and tear and poor workmanship; policyholder's expert report asserted storm damage but contained no photographs and could not displace NFU's photographic evidence; Q2 failed",
        "Leak Source": "Porch roof — chipped tiles, loose mortar sitting in lead valleys and gapping lead flashing where flashing had come away from mortar; loss adjuster found no missing tiles, damaged roofing felt or punctured lead valleys consistent with storm",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "NFU's loss adjuster found no signs of storm damage — photos showed chipped tiles, loose mortar in lead valleys and gapping lead flashing; findings attributed to wear and tear and poor workmanship; damage pattern not consistent with storm-typical indicators such as missing tiles, damaged felt or punctured lead valleys",
        "Evidence Dispute": "NFU relied on its loss adjuster's report and photographic evidence showing wear and tear and poor workmanship; Ms R provided a builder's report asserting storm and rain caused damage including missing tiles and damaged felt, but the report contained no photographs; FOS found the available photos did not show the storm damage Ms R's expert described; minor errors Ms R identified in NFU's report were found not material enough to undermine its overall findings",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; NFU's decision to decline confirmed as fair and reasonable; Ms R's expert report without photographs could not displace NFU's photographic evidence; NFU's handling of the first claim and request for repair estimates also found reasonable; no award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Q2 failed — photos showed no missing tiles, damaged felt or punctured lead valleys consistent with storm; a policyholder's expert report without photographs cannot displace an insurer's expert evidence supported by photos; minor errors in a loss adjuster's report do not undermine its conclusions where photographs support the findings; policy excludes damage happening gradually and damage caused by faulty workmanship",
        "Missing Evidence": "Photographs from Ms R's expert showing missing tiles, damaged roofing felt and punctured lead valleys; independent expert inspection with physical evidence to contradict NFU's photographic findings; evidence ruling out wear and tear and poor workmanship as the cause",
        "Ombudsman Reasoning": "Q1 disputed but FOS addressed Q2 first; Q2 — photos showed chipped tiles, loose mortar in lead valleys and gapping lead flashing consistent with wear and tear and poor workmanship; no missing tiles, damaged felt or punctured lead valleys visible; Ms R's builder's report claimed tiles had come off in high winds and felt was damaged but provided no photographs and could not displace NFU's photographic evidence; minor report errors not material; Q2 answered no so Q3 not reached; NFU's handling of repair estimates and the first claim found reasonable on the call recording evidence",
        "Workflow Insight": "Where an insurer has photographic evidence consistent with its expert's findings of wear and tear, a policyholder's expert report without photographs will not displace that evidence; chipped tiles and loose mortar in lead valleys without missing or dislodged structural elements indicate wear and tear not storm; a loss adjuster report's immaterial errors do not invalidate its conclusions where photos support them; policy exclusions for gradual damage and faulty workmanship apply concurrently with a storm Q2 failure",
        "AI Rule Candidate": "IF insurer_expert_report_supported_by_photographs AND policyholder_expert_report_has_no_photographs THEN insurer_evidence_prevails; IF photos_show_chipped_tiles_and_loose_mortar AND no_missing_tiles_or_damaged_felt_or_punctured_lead_valleys THEN storm_q2 = no AND damage = wear_and_tear; IF loss_adjuster_report_has_minor_errors AND photos_support_overall_conclusions THEN report_errors_not_material_and_conclusions_stand",
        "Source PDF": "DRN-2926772.pdf",
    },
    {
        "Case ID": "STORM-012",
        "FOS Decision ID": "DRN-3173328",
        "Insurer Name": "Aviva Insurance Limited",
        "FOS Decision Date": "4 Jan 2022",
        "Claim Type": "Buildings insurance storm damage claim for water damage to airing cupboard and suspected dropped roof profile; Aviva's surveyor found no storm damage — facias and roof tiles undamaged and in place, chimney flaunching in poor condition; airing cupboard staining attributed to gradual aging not a one-off storm event; £100 delay compensation pre-paid by Aviva confirmed adequate by FOS",
        "Leak Source": "Roof — staining in airing cupboard and suspected dropped roof profile; surveyor found no damage to soffits, facias or roof tiles; all tiles in place; chimney flaunching and pointing in poor condition; staining attributed to aging and gradual deterioration not a storm event",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Aviva's surveyor found no damage to soffits or facias and all roof tiles were in place; staining in the airing cupboard attributed to aging and gradual deterioration rather than a one-off storm event; any problems identified pre-existed the bad weather; independent roofer's photos reviewed by Aviva still showed wear and tear not storm damage",
        "Evidence Dispute": "Aviva relied on its surveyor's report showing no storm indicators and on its review of Miss D's roofer's photos which it still found showed wear and tear not storm damage; Miss D's roofer said a section of tiles was lifting during high winds; FOS agreed Aviva was entitled to rely on its surveyor's finding that problems pre-existed the storm; surveyor found moss growth and chimney flaunching and pointing in poor condition as further indicators of gradual deterioration",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; Aviva's decision to decline confirmed as fair and reasonable; £100 compensation pre-paid by Aviva for delay in communicating findings confirmed as the right amount — the sort of amount FOS would have ordered if not already paid; no further award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Q2 and Q3 not satisfied; where a surveyor finds no damage to soffits, facias or roof tiles and attributes airing cupboard staining to gradual aging rather than a one-off event, an insurer is entitled to rely on that opinion to decline; a roofer's report that tiles were lifting during high winds does not override a surveyor's finding of no storm damage where the insurer finds problems pre-existed the storm; delay in communicating claim decisions warrants compensation independently of the coverage decision",
        "Missing Evidence": "Structural evidence of sudden storm-caused damage rather than gradual deterioration; photographs showing tiles actually missing or structurally displaced rather than merely lifting; evidence distinguishing the property's pre-storm condition from its post-storm condition",
        "Ombudsman Reasoning": "Q1 — storm conditions accepted by Aviva; Q2 and Q3 — surveyor found no damage to soffits or facias, all roof tiles in place, chimney flaunching in poor condition; staining in airing cupboard consistent with gradual aging not a one-off storm event; Miss D's roofer said tiles lifting in high winds but Aviva found problems pre-existed the storm and FOS agreed the insurer was entitled to rely on its surveyor; £100 delay compensation pre-paid by Aviva confirmed as appropriate — FOS would have required the same amount",
        "Workflow Insight": "A surveyor finding all roof tiles in place and no soffit or fascia damage is strong evidence against a storm claim; gradual staining in an interior space is likely deterioration unless a one-off event can be evidenced; delay in communicating claim decisions should be compensated even when the coverage decline is upheld; pre-existing problems identified at inspection defeat storm causation even where storm conditions occurred",
        "AI Rule Candidate": "IF surveyor_finds_no_damage_to_soffits_facias_or_tiles AND interior_staining_attributed_to_gradual_aging THEN storm_q2_q3 = not_satisfied AND decline_reasonable; IF insurer_delayed_communicating_decision AND caused_policyholder_stress THEN compensate_for_delay_regardless_of_coverage_outcome; IF policyholder_roofer_says_tiles_lifting AND insurer_surveyor_finds_problems_preexisted THEN insurer_entitled_to_rely_on_surveyor_and_decline",
        "Source PDF": "DRN-3173328.pdf",
    },
    {
        "Case ID": "STORM-013",
        "FOS Decision ID": "DRN-3211590",
        "Insurer Name": "Aviva Insurance Limited",
        "FOS Decision Date": "26 Jan 2022",
        "Claim Type": "Storm damage to residential roof in December 2020; loss adjuster and independent assessor both found the original construction of the verge tiles was inadequate and this was the dominant cause; storm damage claim declined under faulty workmanship/defective design exclusion; Aviva separately accepted falling debris damage to single storey roof under accidental damage",
        "Leak Source": "Upper roof verge tiles — poor original construction and inadequate fitting identified as dominant cause; previous repointing repairs done by cementing over existing mortar rather than removing it first; debris dislodged from double storey roof damaged the single storey roof below",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Partial",
        "Rejection Reason": "Loss adjuster and independent assessor both concluded the original construction of the verge tiles was inadequate on both the double and single storey roofs; repointing repairs four years prior done incorrectly; damage would not have occurred but for the poor original construction and fitting of verge tiles; policy excludes faulty workmanship, defective design and defective materials",
        "Evidence Dispute": "Aviva relied on its loss adjuster and an independent assessor who both concluded poor original construction of verge tiles was the dominant cause; Miss G argued the assessors focused on previously repaired areas she was not claiming for and that the verges she was claiming were differently constructed; Miss G provided contractor advice that verge tiles are designed to be held by the weight of tiles above with fixings on every 2nd or 3rd row; FOS found the independent assessor's conclusions covered the overall construction of both roofs and were not displaced by contractor statements",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold storm damage claim; Aviva's decision confirmed as reasonable; Aviva had properly accepted falling tiles onto the single storey roof under accidental damage; no additional award on the storm damage claim",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Q3 failed — where two expert assessments (loss adjuster and independent assessor) both conclude poor original construction was the dominant cause, storm is not the main cause even if storm conditions occurred; policy exclusion for faulty workmanship and defective design applies; accepting consequential falling debris under accidental damage while declining the storm claim is a valid partial position; the policyholder must produce evidence of equal weight to contradict expert assessors",
        "Missing Evidence": "Independent evidence of equal weight contradicting both assessors' findings on verge tile construction; physical evidence demonstrating the verge tiles at the specific claimed location were adequately constructed before the storm; evidence ruling out poor original construction as the dominant cause",
        "Ombudsman Reasoning": "Q1 — storm conditions accepted by Aviva; Q2 — possibly consistent with storm (FOS noted this without deciding definitively); Q3 — loss adjuster and independent assessor both found the damage would not have occurred but for poor original construction and fitting of verge tiles on both roofs; Miss G's argument that assessors focused on different roof areas not accepted as the assessor's conclusions covered overall construction; no evidence of equal weight to contradict; Aviva's acceptance of falling debris under accidental damage was a proper and fair partial position",
        "Workflow Insight": "Where two expert assessments both attribute damage to poor original construction, a Q3 failure is well-founded even where the policyholder produces contractor statements to the contrary; accepting consequential damage under accidental damage while declining the main storm claim is a valid and fair approach; policyholders contesting the geographic focus of assessors' reports must produce evidence that the area they are specifically claiming for was differently and adequately constructed",
        "AI Rule Candidate": "IF two_expert_assessments_both_attribute_damage_to_poor_original_construction AND no_contradicting_evidence_of_equal_weight THEN storm_q3 = no AND decline_reasonable; IF policy_excludes_faulty_workmanship_and_defective_design AND experts_conclude_poor_construction_dominant THEN exclusion_applies; IF insurer_accepts_consequential_damage_under_accidental_damage_while_declining_storm THEN partial_acceptance_valid_and_fair",
        "Source PDF": "DRN-3211590.pdf",
    },
    {
        "Case ID": "STORM-014",
        "FOS Decision ID": "DRN-3295758",
        "Insurer Name": "Aviva Insurance Limited",
        "FOS Decision Date": "16 Jun 2022",
        "Claim Type": "Retaining and boundary wall collapse at rear of property attributed to Storm Dennis in February 2020; structural engineer found the retaining wall too thin and inadequate for the three-metre retained height and that the boundary wall had been deteriorating hidden from view; gradual deterioration not storm was the dominant cause; decline confirmed",
        "Leak Source": "Retaining wall (mix of stone, brick and concrete block) collapsed into footpath — structural engineer found wall very thin and inadequate for the three-metre retained height; boundary wall hidden behind it in disrepair and undermined by the newer retaining wall over time; gradual deterioration not a storm event",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Structural engineer's report found the retaining wall very thin and inadequate for the retained height — described as surprising it had lasted as long as it had; remaining wall section bowed and cracked; the boundary wall behind had been undermined by the retaining wall and was in disrepair hidden from view; damage was caused by gradual deterioration not a storm event",
        "Evidence Dispute": "Aviva relied on the structural engineer's report showing the retaining wall was too thin and inadequate, the remaining section bowed and cracked, and the boundary wall in disrepair hidden behind hedging; Mrs V argued the collapse was caused by Storm Dennis and that her garden sloped; FOS agreed the structural evidence showed gradual deterioration and structural inadequacy rather than storm as the dominant cause",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; Aviva's decision to decline confirmed as fair and reasonable; policy exclusion for damage happening gradually applied correctly; other policy sections (flooding, subsidence/landslip) also considered and found inapplicable; no award; question of retaining wall ownership noted as a matter Mrs V could pursue separately",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Q2 and Q3 failed — wall collapse from a structurally inadequate and gradually deteriorating structure is not typical storm damage and storm was not the main cause; policy general exclusion for damage happening gradually over time explicitly lists boundary walls as prone to wear and tear; an inadequate structure described as surprising it lasted so long fails Q3 even where strong storm conditions (Storm Dennis, 70mph peak gust) occurred; subsidence/landslip cover requires the home to be affected — boundary walls alone do not qualify",
        "Missing Evidence": "Evidence that the wall collapse was a sudden one-off event rather than gradual deterioration; independent structural report attributing the collapse primarily to the storm rather than structural inadequacy; evidence that the wall was adequately constructed and maintained before the storm",
        "Ombudsman Reasoning": "Q1 — Storm Dennis 13–19 February 2020, peak gust 70mph, storm conditions confirmed; Q2 — wall collapse could be storm-consistent but structural evidence showed gradual deterioration; Q3 — structural engineer found retaining wall too thin and inadequate for three-metre height and surprising it had lasted so long; remaining section bowed and cracked; boundary wall undermined by retaining wall over time; storm not the main cause; policy exclusion for gradual damage applied; boundary walls listed as wear-and-tear examples; flooding and subsidence/landslip sections considered and found inapplicable",
        "Workflow Insight": "Even where strong storm conditions are confirmed (Storm Dennis 70mph), a wall with structural inadequacy and pre-existing gradual deterioration fails Q3; boundary walls are explicitly identified in many policies as prone to wear and tear — always check the general exclusion language; where two walls are involved (boundary and retaining), both must be assessed individually against the storm questions; always check all relevant policy sections when declining to ensure no alternative peril applies",
        "AI Rule Candidate": "IF structural_engineer_finds_wall_too_thin_and_inadequate AND remaining_section_bowed_and_cracked THEN storm_q3 = no AND damage = gradual_deterioration; IF policy_lists_boundary_walls_as_wear_and_tear_example AND gradual_deterioration_confirmed THEN exclusion_applies; IF storm_conditions_confirmed AND structure_was_pre_existing_inadequate AND surprising_it_lasted_so_long THEN storm_not_main_cause AND decline_reasonable",
        "Source PDF": "DRN-3295758.pdf",
    },
    {
        "Case ID": "STORM-015",
        "FOS Decision ID": "DRN-3574617",
        "Insurer Name": "Lloyds Bank General Insurance Limited",
        "FOS Decision Date": "13 Sep 2022",
        "Claim Type": "Storm damage claim for roof and bay window; Lloyds initially assessed remotely and incorrectly applied subsidence; Mr P complained; roof storm damage subsequently accepted and paid with compensation; bay window declined as pre-existing structural distortion not realigned after historic underpinning — storm revealed but did not cause the bay damage; FOS not upheld on bay window",
        "Leak Source": "Bay window structure — historic distortion not realigned after underpinning works many years earlier; Lloyds' expert identified movement away from house as an ongoing connection issue with evidence of historic distortion; storm revealed the structural weakness but did not cause the underlying damage",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Partial",
        "Rejection Reason": "Lloyds' expert confirmed the bay window damage was an ongoing structural issue with evidence of historic distortion not related to any insured peril; distortion not realigned after underpinning works many years prior; storm identified the weak point but a well-maintained property should withstand all but the most severe conditions; bay window not caused by an insured peril",
        "Evidence Dispute": "Lloyds relied on an expert report (originally commissioned for subsidence) confirming historic distortion with no insured peril as the cause; Mr P's expert acknowledged pre-existing distortion was not realigned after old underpinning but argued the storm placed shock on the building leaving the bay no longer fit for purpose; FOS preferred Lloyds' expert finding that the storm revealed rather than caused the structural weakness; FOS noted a properly commissioned storm survey would not have changed the conclusion since the cause was already identified",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold bay window complaint; roof damage settlement confirmed as fair; £750 compensation already paid by Lloyds for roof and bay handling failures confirmed as adequate; no additional award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Q3 failed for bay window — where an expert identifies pre-existing historic distortion as the ongoing cause and confirms no insured peril caused the damage, storm is not the main cause; storms sometimes reveal existing damage rather than create it; a well-maintained property should withstand all but the most severe conditions; an insurer can rely on a report commissioned under a different peril to identify the actual cause provided the cause is identified; a handling failure in peril assessment does not negate the substantive finding on causation",
        "Missing Evidence": "Evidence that the bay window was in good structural condition before the storm; evidence that the historic underpinning works had successfully realigned the distortion to the required standard; evidence that no structural issues existed before the named storms occurred",
        "Ombudsman Reasoning": "Q1 — three named storms confirmed; Q2 — bay window structure unstable; storm typically dislodges roof not underlying bay structure, but FOS moved to Q3 accepting some storm contribution was possible; Q3 — Lloyds' expert confirmed historic distortion was ongoing and not caused by any insured peril; Mr P's expert acknowledged pre-existing distortion but argued storm caused new structural failure; FOS preferred Lloyds' finding that storm revealed existing weakness; well-maintained property test applied; £750 already paid for handling failures confirmed as adequate; no further award",
        "Workflow Insight": "Where an expert confirms pre-existing historic distortion as the cause and finds no insured peril caused the damage, a storm damage claim fails at Q3 even where some storm contribution occurred; when a storm reveals an existing structural weakness rather than creating damage, the storm is not the main cause; a survey commissioned under the wrong peril can still be used to identify the actual cause of damage; compensation for handling failures does not affect the substantive coverage decision",
        "AI Rule Candidate": "IF expert_confirms_pre_existing_historic_distortion_as_cause AND no_insured_peril_caused_damage THEN storm_q3 = no AND claim_not_established; IF storm_reveals_pre_existing_structural_weakness THEN storm_not_main_cause AND well_maintained_property_test_applies; IF insurer_surveyed_under_wrong_peril AND expert_still_identified_actual_cause THEN coverage_decision_on_correct_cause_stands",
        "Source PDF": "DRN-3574617.pdf",
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
