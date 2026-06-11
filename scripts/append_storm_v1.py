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
        "Case ID": "STORM-016",
        "FOS Decision ID": "DRN-3638410",
        "Insurer Name": "AA Underwriting Insurance Company Limited",
        "FOS Decision Date": "6 Sep 2022",
        "Claim Type": "Garden wall collapsed during storm in November 2021; AA declined citing poor design and faulty workmanship then shifted to vegetation-caused gradual deterioration; no faulty design or workmanship exclusion existed in the policy; vegetation/deterioration argument not evidenced by the surveyor report; FOS confirmed storm conditions and upheld",
        "Leak Source": "Garden wall (residential property boundary structure) — collapsed and found in pieces across the garden during storm; gusts up to 74mph confirmed by FOS weather check; no physical evidence that vegetation had caused gradual mortar breakdown specific to this wall",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Accepted",
        "Rejection Reason": "AA's surveyor concluded poor design and faulty workmanship caused the collapse; AA later asserted vegetation had weakened the wall over time through gradual deterioration; neither reason was supported by a policy exclusion (no faulty design/workmanship exclusion) or by specific evidence of vegetative deterioration in the survey report",
        "Evidence Dispute": "AA relied on surveyor report citing poor design and workmanship and on photos showing vegetation present, arguing vegetation can cause mortar breakdown; Mr and Mrs R argued the wall had stood for years without issue and only came down during this specific storm; FOS checked weather records and found gusts up to 74mph on one tool and up to 60mph on another, both meeting the policy's 55mph storm definition; FOS found AA's vegetation/deterioration argument was a generalisation without specific evidence it applied to this wall",
        "Outcome Category": "Upheld",
        "Outcome": "Upheld — AA Insurance directed to reimburse the cost Mr and Mrs R incurred having their wall rebuilt, subject to the applicable policy excess, plus 8% simple interest from date of invoice payment to date of settlement",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Where a policy contains no exclusion for faulty design or workmanship, an insurer cannot rely on that reason to decline a storm claim; to invoke a gradual deterioration exclusion the insurer must produce specific evidence that applies to the claimed item — a generalisation that vegetation can cause mortar breakdown is insufficient; if a wall has stood without issue for years, this undermines a poor design or deterioration explanation; wind direction at weather station level is not determinative as ground-level conditions can differ due to terrain and surrounding properties",
        "Missing Evidence": "No evidence gaps identified by FOS — all three questions answered yes; the wall's long service history contradicted the poor design theory; AA failed to produce specific evidence that vegetation had caused gradual deterioration of this particular wall",
        "Ombudsman Reasoning": "Q1 — FOS weather check found gusts up to 74mph on one tool and up to 60mph on another, both meeting the policy's 55mph (48 knot/Storm Force 10) storm definition; Q2 — wall found collapsed in pieces across the garden, consistent with storm damage; Q3 — no faulty design or workmanship exclusion in the policy so that basis cannot stand; vegetation/deterioration exclusion not demonstrated as the surveyor report addressed only design and workmanship, not gradual wear; AA's wind direction argument rejected as ground-level conditions can differ from weather station readings; all three questions yes, claim valid",
        "Workflow Insight": "Before declining on faulty design or workmanship, verify the policy actually contains that exclusion; to invoke a gradual deterioration exclusion after the initial surveyor report, the insurer must produce specific evidence applicable to the item — not a general statement about what vegetation can do; if a structure has stood for many years without issue, this is relevant evidence against both poor design and gradual deterioration explanations; where one weather data source shows sub-threshold winds, always check additional sources before concluding storm conditions were absent",
        "AI Rule Candidate": "IF policy_has_no_faulty_design_or_workmanship_exclusion THEN insurer_cannot_decline_on_that_basis; IF insurer_relies_on_gradual_deterioration_exclusion AND evidence_is_only_generalisation_not_specific_to_item THEN exclusion_not_demonstrated AND decline_not_fair; IF multiple_weather_sources_confirm_storm_conditions AND one_source_shows_sub_threshold THEN storm_q1 = yes; IF structure_stood_without_issue_for_years AND insurer_claims_poor_design THEN poor_design_explanation_undermined",
        "Source PDF": "DRN-3638410.pdf",
    },
    {
        "Case ID": "STORM-017",
        "FOS Decision ID": "DRN-3643634",
        "Insurer Name": "UK Insurance Limited",
        "FOS Decision Date": "27 Sep 2022",
        "Claim Type": "Storm damage to residential roof November 2021; claim accepted by UKI; dispute concerns cash settlement quantum (storm portion only, excluding pre-existing wear and tear identified by second contractor) and delays including an incorrect decline letter and file closure; FOS confirmed cash settlement fair and £150 total compensation for delays adequate",
        "Leak Source": "Residential roof — storm damage accepted; second contractor inspection found pre-existing wear and tear in addition to storm damage; UKI limited cash settlement to storm-damaged area only (£1,663.39) as wear and tear area excluded under policy",
        "Property Type": "Residential home",
        "Dispute Type": "Handling / Reinstatement Dispute",
        "Coverage Decision": "Accepted — Disputed Settlement",
        "Rejection Reason": "N/A — storm damage accepted; UKI limited cash settlement to storm-damaged area only, excluding pre-existing wear and tear area identified by second contractor inspection; cash settlement offered at £1,663.39",
        "Evidence Dispute": "UKI relied on second contractor's inspection finding mixed storm damage and pre-existing wear and tear, and on the policy's wear and tear exclusion; Mr W provided three independent quotes all indicating pre-existing issues with one saying a full new roof was needed; Mr W argued delays had worsened the roof condition and that UKI should undertake managed repairs not cash settle; FOS found cash settlement was within policy terms and delays were partly attributable to Mr W's rejection of the first contractor",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; UKI's cash settlement for storm damage only (excluding wear and tear area) confirmed as fair within policy terms; policy permits cash settlement and UKI had offered managed repair which Mr W rejected; £150 total compensation already paid by UKI for delays and incorrect decline letter confirmed as adequate; no additional award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Handling Dispute",
        "Key Policy Clause": "Where an insurer has offered managed repair and the policyholder rejects it, the insurer may switch to a cash settlement at its own-supplier cost rate, which may be less than market rate; when a second inspection reveals pre-existing wear and tear in addition to storm damage, the insurer may limit its cash settlement to the storm portion only in line with the wear and tear exclusion; an insurer is not forced to pay market-rate costs when it has already fulfilled its obligation by offering a repair; delays that are partly attributable to the policyholder's own decisions cannot all be attributed to the insurer",
        "Missing Evidence": "Breakdown of repair costs separating storm damage from wear and tear; evidence of accelerated roof deterioration caused specifically by UKI's delays rather than the pre-existing wear and tear",
        "Ombudsman Reasoning": "Storm damage accepted and a builder appointed within three weeks of notification — no initial delay; Mr W rejected the first contractor (general builders not roofing specialist); second contractor found storm damage and pre-existing wear and tear; UKI incorrectly closed the file causing short Dec–Jan delay; incorrect decline letter sent and later corrected; policy permits cash settlement at own-supplier rate; cash offer of £1,663.39 covers storm damage area only, excluding wear and tear which is policy-excluded; Mr W's three independent quotes all indicated pre-existing issues; £50 for lack of contact and £100 for errors and delay already paid (total £150) confirmed as fair compensation",
        "Workflow Insight": "When a second inspection reveals mixed storm and wear and tear damage, a cash settlement limited to the storm portion only is valid under a wear and tear exclusion; if the policyholder rejects the insurer's initial repair offer, the insurer may offer cash settlement at its own-supplier rate; not all delays in a claim can be attributed to the insurer — some result from the policyholder's own decisions; an incorrect decline letter that is promptly corrected does not require uplift in compensation beyond the insurer's existing offer if that offer is adequate",
        "AI Rule Candidate": "IF insurer_offered_managed_repair AND policyholder_rejected_it AND second_survey_found_mixed_storm_and_wear_and_tear THEN cash_settlement_limited_to_storm_portion_is_fair; IF insurer_has_wear_and_tear_exclusion AND second_inspection_confirms_wear_and_tear_present THEN deduct_wear_and_tear_from_settlement; IF total_compensation_paid_is_adequate_for_delays_caused THEN no_additional_compensation_required",
        "Source PDF": "DRN-3643634.pdf",
    },
    {
        "Case ID": "STORM-018",
        "FOS Decision ID": "DRN-3819182",
        "Insurer Name": "Covea Insurance plc",
        "FOS Decision Date": "9 Jan 2023",
        "Claim Type": "Storm damage to roof during named storm November 2021 causing water ingress; Covea accepted storm conditions but declined citing pre-existing wear and tear highlighted by storm; agent inspection one month before the storm confirmed deterioration with slipped slates, silicon sealing and previous repairs; FOS gave more weight to pre-storm agent evidence than post-repair tradesman statements; internal and accidental damage claims also rejected",
        "Leak Source": "Roof — slipped slates and deterioration with previous repairs confirmed by agent inspection one month before the storm in October 2021; two slates slipped and silicon used to seal slate; storm highlighted existing maintenance issues; internal water damage attributed to pre-existing gradual deterioration not storm",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Covea's agent inspected the roof one month before the storm in October 2021 and reported signs of deterioration, slipped slates, silicon sealing and previous repairs; Covea concluded the storm highlighted pre-existing maintenance issues rather than being the dominant cause; internal damage also declined as resulting from gradual wear and tear; accidental damage rejected as wear and tear is excluded under the general exclusions",
        "Evidence Dispute": "Covea relied on its October 2021 pre-storm agent inspection report showing roof deterioration and slipped slates, and photos taken of the main roof from the front one month before the storm; Mrs B provided reports from a builder and roofer who both said damage was caused by storm not wear and tear; Mrs B argued Covea focused on the rear roof not the storm-damaged front; FOS gave more weight to the pre-storm photographic agent evidence than to post-repair tradesmen's opinions",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; Covea's decision confirmed as fair; agent's October 2021 inspection showing roof deterioration and slipped slates one month before the storm carried more weight than post-repair tradesmen's storm causation opinions; internal damage and accidental damage claims also rejected as wear and tear-driven; no award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Q3 answered no — a pre-storm agent inspection showing active deterioration (slipped slates, silicon sealing, previous repairs) one month before the storm carries more weight than post-repair tradesmen's opinions that the storm caused the damage; a storm that highlights or worsens existing maintenance issues is not the main cause of damage; accidental damage cover defined as unexpected physical damage by unidentifiable external means does not cover damage attributable to gradual wear and tear; the absence of further damage during subsequent storms does not prove the roof was in good condition before the first named storm",
        "Missing Evidence": "Evidence of roof condition at the storm-damaged front specifically, distinct from the October 2021 general inspection; independent expert assessment before repairs distinguishing storm-specific damage from the pre-existing deterioration; evidence contradicting the October 2021 agent report's findings on the main roof",
        "Ombudsman Reasoning": "Q1 — named storm conditions not disputed; Q2 — damage consistent with storm damage; Q3 — October 2021 agent inspection one month before storm found deterioration, slipped slates, silicon sealing and previous repairs; FOS gave this more weight than post-repair tradesmen's statements; storm highlighted existing maintenance issues; Mrs B's argument that Covea focused on rear roof not accepted as agent report covered the main roof from the front; no further damage in subsequent storms not determinative of pre-storm roof condition; accidental damage rejected — policy definition requires unexpected damage by unidentifiable external means, excluding wear and tear under general exclusions",
        "Workflow Insight": "A pre-storm agent inspection showing active deterioration (slipped slates, silicon sealing) is the most powerful Q3 evidence available — reference it specifically in decline reasoning; post-repair tradesmen's storm causation opinions carry less weight than pre-storm photographic inspections; accidental damage cover with a wear and tear exclusion does not provide a fallback route when storm damage fails on a wear and tear basis; the absence of further damage in subsequent storms does not prove the roof was in good condition before the first storm — repairs may have been effective",
        "AI Rule Candidate": "IF pre_storm_inspection_shows_active_deterioration_within_weeks_of_storm AND post_repair_tradesmen_say_storm_damage THEN pre_storm_inspection_carries_more_weight AND storm_q3 = no; IF storm_highlighted_existing_maintenance_issues_confirmed_by_pre_storm_report THEN storm_not_main_cause AND decline_reasonable; IF accidental_damage_cover_excludes_wear_and_tear AND damage_attributable_to_gradual_deterioration THEN accidental_damage_claim_also_fails",
        "Source PDF": "DRN-3819182.pdf",
    },
    {
        "Case ID": "STORM-019",
        "FOS Decision ID": "DRN-3829618",
        "Insurer Name": "Royal & Sun Alliance Insurance Plc",
        "FOS Decision Date": "19 Jan 2023",
        "Claim Type": "Storm damage claim for water ingress to upstairs bedroom ceiling and kitchen; RSA's loss adjuster found no roof ingress and identified rotten window frames with visible daylight through holes as the main water source; ceiling damage described as pre-existing and not a one-off event; Mrs S could not provide evidence of prior roof repairs; Q2 failed on loss adjuster's identification of alternative non-storm ingress source",
        "Leak Source": "Rotten window frames in upstairs bedroom with daylight visible through holes — RSA's loss adjuster identified these as the main ingress source not the roof; bedroom ceiling damage described as pre-existing bowing not consistent with one-off storm event; internal damage described as wear and tear with decayed timber; flat roof inspected and not flagged as a concern",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "RSA's loss adjuster found no water entering through the ceiling and identified rotten window frames with visible daylight through holes as the main ingress source; bedroom ceiling damage described as pre-existing and not a one-off event; internal damage described as wear and tear issues with decayed timber not consistent with storm damage; flat roof inspected and not flagged",
        "Evidence Dispute": "RSA relied on loss adjuster's detailed assessment with photos identifying rotten windowsill as the main ingress source and describing internal damage as wear and tear with decayed timber; Mrs S argued water was entering through the roof not the windowsill and said the roof had been repaired previously; Mrs S could not provide evidence of prior roof repairs; Mrs S's own photos broadly showed the same findings as the survey report; FOS agreed the survey should be relied on as nothing contradicted it",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; RSA's decision confirmed as fair; loss adjuster's identification of rotten window frames as main ingress source not contradicted; ceiling damage pre-existing; internal damage consistent with gradual wear and tear not storm; Mrs S's unsubstantiated claim of prior roof repairs not determinative; no award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Q2 failed — where a loss adjuster identifies a specific non-storm source of ingress (rotten window frames with visible daylight) supported by photos and describes internal damage as pre-existing and wear and tear, the damage is not consistent with storm damage; a policyholder's assertion that water came from a different source requires corroborating evidence to displace the survey finding; unsubstantiated claims of prior roof repairs without evidence do not establish the roof as the ingress source; if the policyholder's own photographs show the same findings as the survey, the survey should be relied on",
        "Missing Evidence": "Evidence of prior roof repairs; photographs or independent expert assessment identifying the roof as the active ingress source; evidence distinguishing roof ingress from window frame ingress; evidence contradicting the loss adjuster's identification of rotten window frames as the main source",
        "Ombudsman Reasoning": "Q1 — storm conditions confirmed; Q2 — loss adjuster found no water entering through ceiling and identified rotten window frames with visible holes as main ingress; bedroom ceiling described as pre-existing bowing and not a one-off event; internal damage described as wear and tear with decayed timber; loss adjuster found water puddles on flat roof but attributed these to recent rainfall not active roof ingress; Q2 answered no; Mrs S's photos broadly showed same findings as survey; prior roof repair claim unsubstantiated; no evidence given to contradict survey; RSA's report relied on",
        "Workflow Insight": "When a loss adjuster identifies a specific alternative non-storm ingress source confirmed by photos, this defeats Q2; a policyholder who asserts a different ingress source but cannot provide corroborating evidence will not displace a survey supported by photos; post-claim photographs from the policyholder that show the same findings as the survey reinforce the insurer's position; prior repairs are a valid basis to support a storm claim only if evidenced — an unsubstantiated assertion of prior repairs does not establish the claimed ingress source",
        "AI Rule Candidate": "IF loss_adjuster_identifies_specific_non_storm_ingress_source AND confirmed_by_photos THEN storm_q2 = no AND decline_reasonable; IF policyholder_asserts_different_ingress_source AND cannot_provide_corroborating_evidence THEN policyholder_assertion_insufficient_to_displace_survey; IF internal_damage_described_as_pre_existing_and_wear_and_tear THEN storm_q2 = no; IF policyholder_photos_show_same_findings_as_survey THEN survey_evidence_reinforced",
        "Source PDF": "DRN-3829618.pdf",
    },
    {
        "Case ID": "STORM-020",
        "FOS Decision ID": "DRN-4293834",
        "Insurer Name": "AXA Insurance UK Plc",
        "FOS Decision Date": "17 Oct 2023",
        "Claim Type": "Storm damage claim for gas flue chimney cowl blown off during winds of 5 January 2023; weather data showed peak winds of 43mph against policy threshold of 55mph (Storm Force 10); surveyor found cowl showed decay and material breakdown not storm damage; Q1 and Q3 both failed; accidental damage claim also failed as wear and tear excluded under that section",
        "Leak Source": "Gas flue chimney cowl — surveyor found signs of decay to the grated section with breakdown in the material identified as the cause; remaining part of cowl still in chimney stack; weather data showed peak winds of 43mph below the 55mph policy storm definition",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Weather data showed peak winds of 43mph below the policy storm threshold of 55mph (48 knots Storm Force 10); surveyor found the chimney cowl had signs of decay to the grated section with breakdown in the material identified as the cause; remaining part of cowl still in chimney stack; accidental damage section also excluded wear and tear",
        "Evidence Dispute": "AXA relied on local weather data showing peak winds of 43mph and surveyor's report with clear photos showing decay and material breakdown; Mr D argued severe winds occurred overnight on 5 January 2023 and that aluminium does not rot or corrode; FOS noted aluminium can corrode and there was no evidence of equal weight to support Mr D's view; FOS agreed Q1 failed on weather data and Q3 failed on surveyor evidence",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; AXA's decision confirmed as fair; Q1 failed — weather data showed peak winds of 43mph below 55mph policy threshold; Q3 also failed — surveyor found material decay and breakdown as cause with photographic evidence; accidental damage claim also failed as wear and tear excluded under that section; no award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Q1 failed — where weather records show peak winds of 43mph against a policy threshold of 55mph (Storm Force 10), storm conditions are not met even if the policyholder believes severe winds occurred overnight; Q3 also failed — surveyor's clear photographic evidence of material decay is sufficient to establish a non-storm cause where the policyholder provides no evidence of equal weight; aluminium can corrode and an insurer is entitled to rely on photographic evidence of decay; wear and tear exclusion in the accidental damage section prevents the policyholder using that section as a fallback when storm damage is declined on wear and tear grounds",
        "Missing Evidence": "Weather data showing winds meeting the 55mph policy threshold on 5 January 2023; independent expert assessment contradicting the surveyor's finding of material decay; evidence that the cowl was in good condition before the claim; evidence that the specific aluminium cowl had not corroded",
        "Ombudsman Reasoning": "Q1 — local weather data showed peak winds of 43mph; policy requires 55mph (48 knots Storm Force 10); storm conditions not met; Q1 answered no; Q2 — chimney cowl blowing off could be consistent with storm damage; Q3 — surveyor found decay to grated section and remaining part still in stack indicating material breakdown not physical dislodgement by wind; surveyor provided clear photos; Mr D's argument that aluminium does not corrode rejected as aluminium can corrode; no evidence of equal weight to support Mr D; Q1 and Q3 both no; accidental damage also declined as wear and tear exclusion applies to that section",
        "Workflow Insight": "Q1 failure on weather records alone defeats a storm claim — where local weather data clearly shows peak winds below the policy threshold, proceed no further; always check accidental damage section exclusions before suggesting it as a fallback for declined storm claims — wear and tear exclusions often apply there too; a surveyor's photographic evidence of material decay defeats Q3 unless the policyholder produces evidence of equal weight; a policyholder's belief that winds were severe is insufficient to override recorded weather data showing sub-threshold winds",
        "AI Rule Candidate": "IF local_weather_data_shows_peak_winds_below_policy_storm_threshold THEN storm_q1 = no AND claim_fails; IF surveyor_finds_material_decay_with_photographic_evidence AND policyholder_provides_no_evidence_of_equal_weight THEN storm_q3 = no AND decline_reasonable; IF storm_damage_declined_and_policyholder_seeks_accidental_damage AND accidental_damage_section_excludes_wear_and_tear AND damage_is_material_decay THEN accidental_damage_claim_also_fails",
        "Source PDF": "DRN-4293834.pdf",
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
