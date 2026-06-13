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
        "Case ID": "STORM-036",
        "FOS Decision ID": "DRN8247030",
        "Insurer Name": "Royal & Sun Alliance Insurance Plc",
        "FOS Decision Date": "26 Sep 2016",
        "Claim Type": "Storm damage to roof of residential home declined; February 2016 claim; RSA's policy defined storm as exceptionally heavy rainfall usually accompanied by high winds; records showed no qualifying heavy rainfall and high winds a week earlier were not linked to physical roof damage; both RSA contractor inspections found poor design and workmanship at valley-gutter junction excluded under policy; policyholder also alleged defective 2013 RSA contractor repairs caused the leak but lacked independent expert evidence",
        "Leak Source": "Roof valleys and boxed guttering junction — two RSA contractor inspections found a gap between roof and boxed guttering resulting from poor design and/or workmanship; external wall staining indicated gradual ongoing leak not a single storm event; 2013 RSA contractors denied their repairs caused the defect",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "No heavy rainfall meeting policy storm definition around the claim date; high winds about a week before not linked to any physical roof damage; both contractor inspections found poor design and workmanship at valley-gutter junction — excluded under policy clause covering 'any loss damage caused by or resulting from poor or faulty design, workmanship or materials'; external wall staining indicated gradual long-term leak not a single storm event; defective 2013 repairs claim unsupported by independent expert evidence",
        "Evidence Dispute": "RSA: weather records showing no qualifying heavy rainfall; high winds one week before claim but no physical damage linked to them; two contractor inspection reports both finding poor design and workmanship at valley-gutter junction; original 2013 contractor denying liability; external wall staining. Mr B: property built 1998 with no evidence of leaks before 2013 claim; argued 2013 RSA contractors caused the defects; only evidence was contractor's own 2016 inspection report and photographs — no independent expert assessment of 2013 repair quality",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; Q1 failed — no heavy rainfall meeting RSA's policy storm definition; high winds a week before not linked to physical roof damage; Q2 also failed — both contractors found poor design and workmanship excluded under policy; defective repairs claim failed — no independent expert evidence that 2013 repairs were defective; internal damage also not storm-caused as leak was gradual; no award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Storm defined as exceptionally heavy rainfall usually accompanied by high winds — absence of qualifying heavy rainfall around the claim date fails Q1; a policy exclusion for 'any loss damage caused by or resulting from poor or faulty design, workmanship or materials' defeats Q2 and Q3 where both contractor inspections identify poor design or workmanship as the cause; external wall staining indicating a gradual multi-year leak is negative evidence against storm causation; a defective contractor repairs claim requires independent expert evidence — the original contractor's denial of liability is not sufficient",
        "Missing Evidence": "Weather records showing heavy rainfall meeting the policy storm definition around the claim date; independent expert evidence (separate from RSA's contractors) establishing that the 2013 repairs were defective and caused the current defect",
        "Ombudsman Reasoning": "Q1 — RSA's policy required exceptionally heavy rainfall usually accompanied by high winds; records showed no qualifying heavy rainfall around the claim date; very high winds occurred about a week before but no physical damage to roof was linked to them; Q1 failed. Q2 and Q3 — even if Q1 had passed, both RSA contractor inspections found poor design and workmanship at the valley-gutter junction (gap between roof and boxed guttering); policy excluded loss caused by poor or faulty design, workmanship or materials; Q2 also failed. External wall staining showed leak had been ongoing for some time not caused by a single storm. Defective repairs: only expert evidence was the contractor who carried out the 2013 repairs and their 2016 photos; contractor denied liability; no independent expert showed 2013 repairs were defective. Internal damage: no storm established and leak was gradual.",
        "Workflow Insight": "Where a policy defines storm as requiring exceptionally heavy rainfall, document weather records showing rainfall levels specifically — sub-storm rainfall is a standalone Q1 failure independent of wind speed; a faulty design or workmanship exclusion provides a dual basis for declining the claim if the storm analysis also fails; where a policyholder alleges defective insurer-arranged contractor repairs, the insurer should proactively commission an independent expert assessment to avoid an unsupported 'he said/she said' position at FOS",
        "AI Rule Candidate": "IF policy_storm_definition_requires_heavy_rainfall AND no_qualifying_rainfall_recorded THEN storm_q1 = no; IF policy_excludes_poor_design_or_workmanship AND contractor_inspections_find_poor_design_or_workmanship THEN q2_also_fails; IF policyholder_alleges_defective_contractor_repairs AND only_evidence_is_contractor_self_assessment THEN defective_repairs_claim_fails_for_lack_of_independent_expert_evidence",
        "Source PDF": "DRN8247030.pdf",
    },
    {
        "Case ID": "STORM-037",
        "FOS Decision ID": "DRN8636254",
        "Insurer Name": "Aviva Insurance Limited",
        "FOS Decision Date": "21 Nov 2016",
        "Claim Type": "Storm damage to roof of block of flats partially accepted; tiles dislodged and paid by Aviva; felt and battens underneath declined as wear and tear; residents association argued felt was punctured by dislodged tiles during storm; surveyor and FOS found felt and battens had rotted gradually over time — storm highlighted maintenance issue not caused it",
        "Leak Source": "Roof tiles — dislodged by storm (accepted and paid by Aviva); felt and battens underneath — found crumbling and rotted gradually over time by Aviva's surveyor; storm highlighted pre-existing maintenance issue rather than causing felt and batten damage",
        "Property Type": "Block of flats",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Partial",
        "Rejection Reason": "Aviva's surveyor found felt and battens in poor state of repair, having rotted over a period of time; photographs showed felt crumbling and battens in poor condition consistent with gradual wear and tear not storm damage; storm highlighted an existing maintenance issue rather than causing the damage; policy covers one-off storm damage not gradual maintenance deterioration",
        "Evidence Dispute": "Aviva: surveyor report finding felt and battens in poor state of repair, rotted over time, storm highlighted maintenance issue; photographs showing felt crumbling and battens in poor condition. Residents association B: argued felt was directly punctured by the storm-dislodged tiles; contended damage to felt and battens would not have occurred were it not for the storm",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold; Aviva's partial acceptance of tile damage confirmed as correct; felt and battens decline confirmed — Q3 failed; surveyor and photographs showed gradual rot and deterioration not storm-caused damage; Aviva's decision to decline felt and battens confirmed as reasonable; no further award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Q3 — damage to felt and battens that shows crumbling and rot consistent with gradual deterioration over time is not caused by a storm even where tiles above were storm-damaged; policy covers one-off storm incidents not maintenance deterioration that comes to light at the time of a storm; damage to felt and battens is not typical of storm damage (though possible in certain circumstances) — where photographs show rot and crumbling the Q3 answer is no; an insurer can accept storm damage to tiles while simultaneously declining felt and battens on wear and tear grounds where the two damage types have different causation",
        "Missing Evidence": "Expert or photographic evidence showing the felt was intact immediately before the storm and was specifically punctured or torn by the dislodged tiles rather than already rotted",
        "Ombudsman Reasoning": "Q1 — Aviva accepted storm; yes. Q2 — tile damage accepted; damage to felt and battens not typical of storm damage but FOS accepted it could happen in certain circumstances. Q3 — Aviva's surveyor said felt and battens in poor state, rotted over time; storm highlighted maintenance issue, did not cause the damage; FOS reviewed photographs which showed felt crumbling and battens in poor condition — consistent with gradual wear and tear over time not damage from a single storm; Q3 answered no for felt and battens. FOS noted it was not considering the claim for internal damage (which was not pursued as the flat occupier sadly passed away).",
        "Workflow Insight": "When accepting partial storm damage (tiles), assess felt and battens separately under Q3 — their condition may reflect gradual deterioration even where tiles above were storm-damaged; document the surveyor's specific findings on felt and batten condition with photographs to support a partial decline; where photographs show crumbling felt and rotted battens, Q3 for that element fails regardless of what caused tile damage above; a one-off storm does not cause rot — rot is definitionally gradual",
        "AI Rule Candidate": "IF tiles_accepted_as_storm_damaged AND felt_and_battens_show_crumbling_and_rot THEN felt_and_battens_q3 = no AND partial_decline_reasonable; IF surveyor_finds_gradual_rot_in_felt_and_battens THEN storm_not_main_cause_of_felt_and_batten_damage; IF damage_element_shows_gradual_deterioration THEN assess_q3_independently_from_other_accepted_storm_damage_elements",
        "Source PDF": "DRN8636254.pdf",
    },
    {
        "Case ID": "STORM-038",
        "FOS Decision ID": "DRN8660161",
        "Insurer Name": "Zurich Insurance PLC",
        "FOS Decision Date": "29 Feb 2016",
        "Claim Type": "Three-claim complaint: subsidence 2006 declined (no evidence); storm damage 2013 declined (gradual deterioration, no AD cover for internal damage); storm damage 2015 declined (same roofs in poor condition, storm not main cause); policyholder's specific storm event claims unsupported by evidence; surveyor reports showed 60-80 year old main roof in very poor condition",
        "Leak Source": "Main roof and flat roof of residential home — surveyor reports for both 2013 and 2015 claims showed roofs in very poor state of repair; main roof 60-80 years old with very worn mortar and nails, temporary repairs made; flat roof old for its type; 2015 roofer report noted materials had blown off but FOS found storm merely highlighted pre-existing poor condition",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Subsidence 2006: loss adjuster found no evidence of subsidence and no contrary evidence provided. Storm 2013 and 2015: main roof estimated 60-80 years old with very worn mortar and nails; flat roof old for its type; surveyor reports for both claims showed poor condition due to lack of maintenance; policyholder claimed greenhouse glass blown onto roof (2013) and branch blown against chimney (2015) but neither was supported by evidence; storm highlighted poor condition rather than causing damage; gradually-occurring maintenance damage excluded under policy",
        "Evidence Dispute": "Zurich: loss adjuster letter (2006) finding no subsidence evidence; surveyor reports for both 2013 and 2015 claims showing roofs in very poor condition due to lack of maintenance; 2015 roofer report noting materials had blown off (accepted by FOS as possible but not determinative of causation). Miss J: claimed roofs had been maintained (contradicted by photos and reports); claimed 2013 damage caused by storm-blown greenhouse glass onto roof; claimed 2015 chimney damaged by storm-blown branch; claimed roof damage caused by subsidence — no evidence provided for any of these specific events",
        "Outcome Category": "Not Upheld",
        "Outcome": "FOS did not uphold on any head; subsidence claim failed — no evidence provided; both storm claims failed — surveyor reports showed roofs in very poor condition from lack of maintenance; specific storm event claims (greenhouse glass, blown branch) unsupported by evidence; storm highlighted poor condition rather than causing damage; gradually-occurring maintenance damage excluded; 2013 internal damage correctly declined as no AD cover; 2015 internal damage — no evidence of any internal damage; no award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Q3 failed for both storm claims — roofs in very poor condition from lack of maintenance (main roof 60-80 years old, mortar and nails very worn) with storm merely highlighting the pre-existing condition; a policyholder's unsupported assertion of a specific storm event (greenhouse glass, blown branch) is not sufficient to establish storm causation; gradually-occurring damage from lack of maintenance is excluded under policy; absence of AD cover correctly prevents recovery of internal damage under the 2013 policy; where AD cover exists in 2015 but no evidence of internal damage is presented the claim also fails",
        "Missing Evidence": "Evidence supporting the specific 2013 storm event (greenhouse glass blown onto roof) and 2015 event (branch blown against chimney); evidence that roofs were in good condition and maintained before the storms; evidence of internal damage in 2015",
        "Ombudsman Reasoning": "Subsidence 2006: loss adjuster letter found no evidence; Miss J provided nothing to contradict — not upheld. Storm 2013 and 2015: storms occurred around both claim dates; Miss J claimed greenhouse glass blown onto roof (2013) — nothing to support this; claimed branch blown against chimney (2015) — nothing to support this; surveyor reports for both claims showed roofs in very poor condition — main roof 60-80 years old, mortar and nails very worn, temporary repairs made, flat roof old for its type; Miss J claimed roofs were maintained but photos and reports contradicted this; 2015 roofer noted materials had blown off — FOS accepted this may have happened but concluded storm just highlighted poor condition rather than causing damage; Q3 no for both claims; gradually-occurring damage from lack of maintenance excluded; 2013 internal damage: correctly declined as no AD cover; 2015 internal damage: AD cover existed but no evidence of any internal damage presented.",
        "Workflow Insight": "Where a policyholder makes multiple claims over multiple years with roofs consistently found to be in poor condition, the accumulation of surveyor reports showing unmaintained roofs is strong Q3 evidence; a policyholder's assertion of a specific storm trigger event (object blown onto roof, branch impact) must be supported by evidence — unsupported assertions will not displace consistent surveyor findings of general poor condition; when AD cover is absent, decline internal damage at first handling with a clear explanation to avoid a separate complaint",
        "AI Rule Candidate": "IF surveyor_reports_across_multiple_claims_show_poor_roof_condition AND policyholder_claims_specific_storm_trigger_unsupported THEN storm_q3 = no_for_all_claims; IF gradually_occurring_damage_from_lack_of_maintenance THEN excluded_under_policy; IF no_ad_cover AND internal_damage_claimed THEN internal_damage_declined_correctly; IF ad_cover_exists AND no_evidence_of_internal_damage THEN internal_damage_claim_also_fails",
        "Source PDF": "DRN8660161.pdf",
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
