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
        "Case ID": "STORM-026",
        "FOS Decision ID": "DRN0445901",
        "Insurer Name": "Gresham Insurance Company Limited",
        "FOS Decision Date": "30 Nov 2015",
        "Claim Type": "Storm damage to exterior of residential home declined; water leaking into two bedrooms and living room in October 2014; two loss adjusters found wear and tear not storm damage; policyholders' roofer addressed damage scope only, not causation; internal repairs separately paid under accidental damage cover",
        "Leak Source": "Roof — water ingress through roof covering to two bedrooms and living room; two loss adjusters found no storm damage; policyholders' roofer had better roof access but report addressed repair scope not storm causation",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Partial",
        "Rejection Reason": "Two independent loss adjusters both concluded damage was due to wear and tear not storm damage; one explicitly stated 'no storm damage to roof'; policyholders' own roofer's report addressed only the extent of damage and repair costs, not the cause of damage",
        "Evidence Dispute": "Insurer: two concordant loss adjuster reports both finding wear and tear not storm damage, one explicitly stating no storm damage to roof; photographs reviewed. Policyholders: specialist roofer who accessed the roof directly and provided a repair quote; argued roofer's on-roof inspection was more thorough than loss adjusters' ground-level assessments; FOS accepted roofer had better access but found report silent on causation",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; Gresham's external damage decline confirmed as fair; Q2 failed — two concordant expert opinions of no storm damage not displaced by roofer's report which addressed repair scope only; internal repairs and £100 compensation already paid by Gresham pre-FOS not disturbed",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Q2 failed — a policyholders' specialist report that addresses damage extent and repair costs but not storm causation carries no evidential weight on the Q2 question, even if the specialist had superior roof access; two concordant independent expert findings of no storm damage are sufficient to decline the storm peril even where the policyholder's expert had better physical access to the damage",
        "Missing Evidence": "Expert opinion from policyholders' roofer specifically addressing storm causation (the report addressed repair scope only, not whether the damage was caused by a storm); independent assessment contradicting the two loss adjusters' concordant finding of no storm damage",
        "Ombudsman Reasoning": "Q1 — storms confirmed in the area in the ten days before the damage was reported; Q2 — two loss adjusters both found wear and tear not storm damage; one explicitly stated 'no storm damage to roof'; policyholders' roofer had better access to some roof areas but his report said nothing about the cause of damage — only what damage existed and the repair cost; FOS could not find anything in the roofer's report contradicting the two expert opinions; Q2 answered no; Q3 not considered",
        "Workflow Insight": "Two concordant loss adjuster reports finding no storm damage are sufficient to decline Q2 where the policyholder's counter-evidence (a roofer's report) is silent on causation; always ensure decline reasoning explicitly references the lack of causation opinion in the policyholder's counter-report — the report's silence on causation, not its access limitations, is the decisive point; internal accidental damage cover should be assessed independently from the storm peril at first handling to avoid separate complaint exposure",
        "AI Rule Candidate": "IF two_independent_expert_reports_find_no_storm_damage AND policyholder_counter_report_silent_on_causation THEN storm_q2 = no AND decline_reasonable; IF policyholder_report_addresses_repair_scope_only_not_causation THEN report_carries_no_weight_on_storm_q2; IF internal_damage_payable_under_accidental_damage THEN assess_independently_from_storm_peril_at_first_handling",
        "Source PDF": "DRN0445901.pdf",
    },
    {
        "Case ID": "STORM-027",
        "FOS Decision ID": "DRN1086734",
        "Insurer Name": "U K Insurance Limited",
        "FOS Decision Date": "12 Sep 2016",
        "Claim Type": "Storm damage to roof causing water ingress declined; external roof repair declined and internal damage paid; winds reached storm threshold on only two isolated occasions and heavy rain came days later not concurrent; loss adjuster found no storm damage; policyholders did not commission specialist storm damage report despite UKI inviting one",
        "Leak Source": "Roof — rainwater ingress through roof covering; loss adjuster found no evidence of storm damage and attributed cause to general maintenance issues; no expert report identifying storm-specific damage provided by policyholders",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Partial",
        "Rejection Reason": "Winds reached borderline storm level on only two isolated occasions and were not concurrent with the very heavy rain that came days later; loss adjuster found no evidence of storm damage to the roof, attributing cause to general maintenance issues; policyholders did not provide a specialist roofer's report confirming storm damage despite UKI offering to reconsider on that basis",
        "Evidence Dispute": "Insurer: loss adjuster report finding no storm damage and identifying maintenance as the cause; weather records showing borderline winds on only two occasions not concurrent with heavy rain. Policyholders: high-profile storm warnings for their area; argued UKI's loss adjuster was not professional; did not commission a specialist roofer's storm damage report despite UKI's invitation to do so",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; external roof decline confirmed; Q1 borderline and unlikely to have damaged a well-maintained roof; Q2 failed — no expert evidence of storm damage to roof available; UKI's offer to reconsider on receipt of specialist report confirmed as fair and sufficient; no further action required",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Where wind speeds are borderline and reached storm threshold on only isolated occasions without being concurrent with heavy rain, and the only professional evidence available is the loss adjuster's finding of no storm damage, the claim fails; an insurer's conditional offer to reconsider upon receipt of a specialist storm damage report is a fair and proportionate response that FOS will not improve upon; a policyholder who fails to commission an expert report when explicitly invited to do so cannot argue the insurer has not considered their position",
        "Missing Evidence": "Specialist roofer's report confirming storm-specific damage to the roof (UKI offered to reconsider on this basis but no such report was provided)",
        "Ombudsman Reasoning": "Q1 — weather evidence showed consistently strong winds with storm-level gusts on only two isolated occasions; heavy rain came a few days later, not concurrent with the storm winds; FOS thought winds at this level unlikely to have damaged a well-maintained roof; Q2 — loss adjuster report found no evidence of storm damage, attributing cause to maintenance; policyholders provided no expert report contradicting this despite UKI's invitation; Q2 failed; UKI's offer to reconsider on receipt of specialist report confirmed as reasonable",
        "Workflow Insight": "Where winds are borderline and not concurrent with heavy rain, and the loss adjuster finds maintenance issues rather than storm damage, decline the storm peril and offer in writing to reconsider on receipt of a specialist storm damage report — this creates a fair and FOS-endorsed position; a policyholder who fails to take up that offer cannot subsequently argue the insurer made an uninvestigated decision",
        "AI Rule Candidate": "IF winds_borderline_storm_threshold AND not_concurrent_with_heavy_rain AND loss_adjuster_finds_no_storm_damage THEN storm_q1_and_q2_fail AND decline_reasonable; IF insurer_offers_to_reconsider_on_specialist_report AND policyholder_does_not_provide_report THEN insurer_position_confirmed_as_fair; IF no_expert_evidence_of_storm_damage_available THEN storm_q2 = no",
        "Source PDF": "DRN1086734.pdf",
    },
    {
        "Case ID": "STORM-028",
        "FOS Decision ID": "DRN1681509",
        "Insurer Name": "Kwik-Fit Insurance Services Ltd",
        "FOS Decision Date": "22 Jul 2016",
        "Claim Type": "Home insurance policy mis-sold without storm cover despite being told it was included; broker accepted mis-selling and offered £100 compensation; FOS assessed whether storm claim would have succeeded if correct policy held — concluded no, as photos showed no storm damage and damage consistent with gradual deterioration",
        "Leak Source": "Roof — water ingress attributed to gradual deterioration; loss adjuster found no missing tiles or blown-off felt; photos showed no evidence of storm damage; internal damage appeared to have occurred over a long period",
        "Property Type": "Residential home",
        "Dispute Type": "Broker Conduct Dispute",
        "Coverage Decision": "Not Applicable",
        "Rejection Reason": "Policy did not include storm cover (mis-sold); on hypothetical assessment, storm claim would have failed anyway — loss adjuster found no storm damage externally; photos showed no missing tiles or blown felt; internal damage appeared to have occurred over a long period; damage consistent with gradual deterioration over time not a storm event",
        "Evidence Dispute": "Broker/insurer: loss adjuster report finding no storm damage; photographic evidence showing no missing tiles or blown-off felt. Claimant: builder's statement that roof was storm damaged and felt had been blown off; surveyor's letter confirming repair quotation was proportionate and fair (but not confirming storm causation); FOS found photos contradicted builder's storm claim and surveyor's letter addressed only cost proportionality not causation",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; Mr D no worse off from mis-selling — storm claim would have failed on the merits; £100 compensation already paid by Kwik-Fit for mis-selling adequate; no further action required",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Broker Dispute",
        "Key Policy Clause": "Where a policy is mis-sold, FOS assesses whether the policyholder is worse off by hypothetically evaluating whether the correct claim would have succeeded; photos showing no missing tiles, no blown-off felt, and internal damage consistent with long-term water ingress defeat Q2 and Q3; a builder's statement of storm damage and a surveyor's letter confirming only cost proportionality do not constitute expert storm causation evidence sufficient to displace loss adjuster photographic findings",
        "Missing Evidence": "Expert storm causation opinion confirming the roof was damaged by a specific storm event (builder's statement and surveyor's cost letter did not address causation); photographs showing storm-consistent damage such as missing tiles or blown felt",
        "Ombudsman Reasoning": "Kwik-Fit accepted mis-selling; FOS assessed hypothetically whether storm claim would have succeeded if correct policy held; loss adjuster photos showed no missing tiles or blown-off felt; internal damage appeared to have occurred over a long period; on balance photos supported gradual deterioration not storm damage; builder's storm assertion not supported by photos; surveyor's letter confirmed only cost proportionality not storm causation; storm claim would most likely have failed; Mr D no worse off; £100 compensation adequate",
        "Workflow Insight": "When assessing mis-selling complaints involving storm claims, apply the same three-question framework hypothetically — if the underlying claim would have failed, the policyholder has not been materially harmed; photographic evidence of no missing tiles and no blown-off felt is sufficient to defeat Q2 where the policyholder's only counter-evidence is a builder's assertion and a cost-proportionality letter",
        "AI Rule Candidate": "IF policy_mis_sold AND storm_claim_would_fail_on_merits_hypothetically THEN policyholder_not_materially_worse_off AND compensation_for_mis_selling_only; IF photos_show_no_missing_tiles AND no_blown_felt AND internal_damage_gradual THEN storm_q2_and_q3 = no; IF counter_evidence_is_builder_assertion_and_cost_letter_only THEN insufficient_to_displace_photographic_loss_adjuster_findings",
        "Source PDF": "DRN1681509.pdf",
    },
    {
        "Case ID": "STORM-029",
        "FOS Decision ID": "DRN2201217",
        "Insurer Name": "Elite Insurance Company Limited",
        "FOS Decision Date": "8 Jul 2018",
        "Claim Type": "Storm damage to roof and internal damage declined in full; policy defined storm as Force 10 (55-63mph); maximum wind speed in month before claim was 40mph; surveyor found roof in poor condition due to age with loose nails and gaps; internal damage also declined as not accidental — water had leaked gradually over a period of time",
        "Leak Source": "Roof — surveyor found poor condition due to age, with fixing nails working loose leaving gaps for water ingress; water leaked through to extension and kitchen ceiling over a period of time; no storm-force winds recorded",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Policy defined storm as Force 10 winds of 55-63mph or above; weather records showed maximum wind speed of 40mph in the month before the claim — well below the policy threshold; surveyor found roof in poor condition due to age with loose nails creating gaps; internal damage also declined as not accidental — policy required damage to be sudden, unexpected, happening at a specific time, and caused by something external and identifiable; water had leaked over a period of time satisfying none of these conditions",
        "Evidence Dispute": "Insurer: weather records showing maximum 40mph winds; surveyor report describing poor roof condition due to age with loose nails and gaps. Claimant: denied roof was in poor condition; had previously carried out some repairs; FOS accepted it was reasonable to rely on the expert surveyor evidence",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; both storm and accidental damage claims failed; Q1 failed on weather records (40mph against 55-63mph threshold); roof condition confirmed poor by surveyor; internal damage also excluded as water leaked gradually over time — not sudden, unexpected, or happening at a specific time as required by the accidental damage definition; no award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Policy-defined storm threshold of Force 10 (55-63mph); recorded maximum winds of 40mph clearly fail Q1 on numerical grounds; the accidental damage policy definition requiring damage to be 'sudden, unexpected, happening at a specific time, caused by something external and identifiable' excludes water ingress that has leaked gradually over a period of time — all four conditions must be met; prior repairs by a claimant do not displace a surveyor's contemporaneous finding of poor roof condition",
        "Missing Evidence": "Weather data showing winds meeting the Force 10 / 55-63mph policy threshold; independent expert assessment contradicting the surveyor's finding of poor roof condition; evidence that water ingress was sudden and occurred at a specific identifiable time",
        "Ombudsman Reasoning": "Q1 — policy defines storm as Force 10 (55-63mph); weather records showed maximum 40mph in the month before the claim; 40mph is nowhere near sufficient to damage a roof in sound condition; Q1 answered no; surveyor found roof in poor condition due to age with loose fixing nails creating gaps; Ms D's prior repairs noted but expert evidence relied on; internal accidental damage assessed separately — policy required sudden, unexpected, at a specific time, caused by something external and identifiable; surveyor noted water had leaked over a period of time to extension and kitchen ceiling; does not meet any of the suddenness conditions; both claims declined",
        "Workflow Insight": "Where a policy contains an explicit numeric storm threshold (Force 10 / 55mph), cite the exact recorded wind speed and the threshold in the decline reasoning — the numerical gap speaks for itself; always assess accidental damage independently from the storm claim and check whether the policy's suddenness conditions are met — gradual water ingress fails the accidental damage definition even where storm also fails",
        "AI Rule Candidate": "IF policy_storm_threshold_numeric AND recorded_winds_below_threshold THEN storm_q1 = no AND decline; IF accidental_damage_requires_sudden_unexpected_specific_time AND water_ingress_was_gradual_over_period THEN accidental_damage_also_fails; IF surveyor_finds_poor_roof_condition AND policyholder_denies_AND_cites_prior_repairs THEN surveyor_contemporaneous_finding_takes_precedence",
        "Source PDF": "DRN2201217.pdf",
    },
    {
        "Case ID": "STORM-030",
        "FOS Decision ID": "DRN2738252",
        "Insurer Name": "U K Insurance Limited",
        "FOS Decision Date": "20 Jun 2020",
        "Claim Type": "Attic leak and wind-damaged door after stormy weather; UKI assessed from photographs without physical survey; roof repair declined as damage consistent with gradual build-up not storm; door damage paid under accidental damage; policyholders argued physical survey should have been required",
        "Leak Source": "Roof valley between two roof sections — photographs showed accumulated debris and removed tiles in a sheltered location; repairs required clearing the valley and replacing battens and felt, consistent with gradual maintenance build-up; no damage reported to more exposed roof areas",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Partial",
        "Rejection Reason": "Photographs showed damage in sheltered roof valley with accumulated debris characteristic of long-term collection, not storm impact; no damage to more exposed roof areas as would be expected from a storm; repair scope (clearing valley, replacing battens and felt) typical of gradual maintenance build-up not a one-off storm event; physical survey not required where photographs are sufficiently clear",
        "Evidence Dispute": "Insurer: photographs from policyholders showing debris-filled valley, removed tiles in sheltered location, and repair scope typical of gradual build-up; storm conditions in prior two weeks accepted. Policyholders: argued UKI should have conducted a physical survey before declining; said damage was caused by the storm; UKI agreed door damage was accidental",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; Q2 and Q3 failed — valley location with accumulated debris, no damage to exposed areas, and repair scope all inconsistent with storm as the main cause; physical survey not required; door accidental damage payment not disturbed; no further action required",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Q2 failed — damage located in a sheltered roof valley with accumulated debris is not consistent with storm damage; where storm damage to one area would be expected to cause more extensive damage to more exposed areas and none is reported, this undermines the storm causation argument; Q3 failed — repair scope of clearing a valley and replacing battens and felt is characteristic of gradual maintenance build-up not a one-off storm event; a physical survey is not required where photographic evidence is sufficiently clear to assess the claim",
        "Missing Evidence": "Expert assessment establishing storm as the specific cause of the valley damage; evidence of intact roof condition immediately before the storm; evidence of damage to more exposed roof areas consistent with storm impact",
        "Ombudsman Reasoning": "Q1 — storm conditions in the two weeks before the claim accepted by UKI; Q2 — photos showed damage in valley between two roofs, a sheltered location; debris visible consistent with long-term accumulation; no damage to more exposed areas reported, which would be expected if storm had caused the valley damage; if pre-repair photo, did not show typical storm damage; Q2 answered no; Q3 — repair description (clearing valley, replacing battens and felt) typical of gradual build-up not one-off storm; valley may require more frequent maintenance by design; Q3 answered no; physical survey not required — reasonably clear photos sufficient; door damage correctly paid under accidental damage",
        "Workflow Insight": "A sheltered location (roof valley) combined with accumulated debris and no storm damage to more exposed areas is a strong Q2 failure argument — document both points explicitly in decline reasoning; the repair scope (clearing a valley, replacing battens and felt) is a reliable Q3 indicator of gradual build-up rather than storm impact; an insurer can rely on clear photographs in lieu of a physical survey — this is FOS-endorsed and does not give rise to a complaint",
        "AI Rule Candidate": "IF damage_in_sheltered_location AND accumulated_debris_visible AND no_damage_to_exposed_areas THEN storm_q2 = no; IF repair_scope_includes_clearing_valley_and_replacing_battens_and_felt THEN storm_q3 = no_as_gradual_buildup; IF photographic_evidence_sufficiently_clear THEN physical_survey_not_required AND insurer_position_not_undermined",
        "Source PDF": "DRN2738252.pdf",
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
