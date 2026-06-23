"""
Standard append script for Subsidence Case Database — Schema v1 (21 columns).
Column 6 is "Movement Cause" (physical cause of ground movement).

Usage
-----
1. Read the source PDF(s) and extract the fields listed in NEW_CASES below.
2. Add one dict per case to NEW_CASES following the extraction rules.
3. Run from the repo root:
       py scripts/append_subsidence_v1.py

Appends NEW_CASES rows to:
    knowledge/case-databases/Subsidence_Case_Database.xlsx

===========================================================================
FIELD EXTRACTION RULES
===========================================================================

Case ID           : Format SUBS-NNN (zero-padded to 3 digits)
FOS Decision ID   : DRN-XXXXXXX or DRNXXXXXXX as printed in the PDF
Insurer Name      : Formal registered name from the FOS decision
FOS Decision Date : DD Mon YYYY — accept-or-reject deadline in final paragraph
Claim Type        : Physical incident and nature of dispute in one sentence
Movement Cause    : Cause of ground movement — physical mechanism
                    e.g. "Tree roots causing clay soil shrinkage"
                         "Escape of water softening ground — water-induced subsidence"
                         "Excessive groundwater moisture from poor surface drainage"
Property Type     : "Residential home" / "Leasehold flats" / "Commercial" / etc.
Dispute Type      : Controlled vocab (7 values)
Coverage Decision : Controlled vocab (5 values)
Rejection Reason  : Insurer's stated reason for declining or disputing
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
Source PDF        : Filename only (e.g. DRN0001741.pdf)
===========================================================================
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

COLUMNS = [
    "Case ID",
    "FOS Decision ID",
    "Insurer Name",
    "FOS Decision Date",
    "Claim Type",
    "Movement Cause",
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
# NEW CASES — Batch 4: SUBS-031 to SUBS-032  (Final batch — database complete)
# ---------------------------------------------------------------------------
NEW_CASES: list[dict] = [
    {
        "Case ID": "SUBS-031",
        "FOS Decision ID": "DRN8130715",
        "Insurer Name": "UK Insurance Limited",
        "FOS Decision Date": "15 Jul 2015",
        "Claim Type": "Home insurance — 2009 subsidence claim recorded against extension crack (clay shrinkage from willow tree); repair cost below £1,000 policy excess so no contractor repairs carried out; UKI failed to explain to Mr A that not proceeding with insurer-managed repairs meant no Certificate of Structural Adequacy would be issued; Mr A unable to sell home or insure elsewhere; dispute about removal of subsidence record and insurer's obligation to provide COSA",
        "Movement Cause": "Clay shrinkage from willow tree (high water demand) causing minor subsidence to extension/main house junction in 2009; willow tree pruned and chemically treated as mitigation; property stabilised by 2013 with no recent or progressive subsidence found at 2013 inspection",
        "Property Type": "Residential home",
        "Dispute Type": "Handling / Reinstatement Dispute",
        "Coverage Decision": "Accepted",
        "Rejection Reason": "UKI: 2009 subsidence correctly recorded (loss adjuster's local knowledge; willow tree; tapering crack; ten years since extension built so not settlement); repair cost fell below £1,000 excess so not carried out; willow tree pruned and chemically treated as mitigation; property since stabilised; 2013 engineer report confirms no current subsidence; Certificate of Structural Adequacy only issued after contractors complete physical repairs — no repairs = no COSA; modified engineer's report provided instead",
        "Evidence Dispute": "Mr A: loss adjuster's 2009 investigation was superficial (no inspection trench, no monitoring); 2013 engineer report found no evidence of recent, historic or ongoing subsidence — proves 2009 recording was wrong; two cracks still same size and shape proving they were never subsidence; UKI told him damage was repaired but no repairs were carried out. UKI: loss adjuster's local knowledge and willow tree justified classification; damage recovered and no further movement proves cause correctly identified; 2013 report confirms no current subsidence; not required to remove valid claim. FOS adjudicator: 2009 investigation should have been more detailed; but subsidence recording reasonable given local knowledge, tapering crack, willow tree. FOS ombudsman: on balance minor subsidence did occur in 2009 — local knowledge, slight tapering, 10 years since extension (settlement less likely), willow tree nearby; UKI investigation poor — no monitoring, no in-depth investigation; cannot direct removal of valid subsidence claim; 2013 report finding no current subsidence does not prove the 2009 cracks were not subsidence; UKI's provided engineer's report was inadequate (cut-and-paste from 2013 report with 'historic' inserted); COSA is appropriate remedy",
        "Outcome Category": "Upheld",
        "Outcome": "UK Insurance Limited to provide Mr A with a Certificate of Structural Adequacy in respect of the 2009 claim; pay £500 compensation for distress and inconvenience caused by poor handling of subsidence claim",
        "Compensation Awarded (£)": 500,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Where a subsidence claim is accepted but repair cost falls below the policy excess and no contractor repairs are carried out, the insurer must explain to the consumer that not proceeding with insurer-managed repairs means no Certificate of Structural Adequacy will be issued; failure to give this warning causes lasting harm (inability to sell, inability to move insurer) for which the insurer is responsible; the appropriate remedy is to provide a genuine COSA or to fund an independent structural engineer's report acceptable to future insurers and purchasers; FOS cannot direct removal of a valid subsidence claim from records even if the original investigation was poor — the remedy is to address the consequences of poor handling, not to re-characterise the claim history; chemically treating and pruning vegetation (even if not physically removed) constitutes remedial work; an inadequate document cut-and-pasted from a prior report does not satisfy the insurer's obligation to provide a COSA-equivalent document",
        "Missing Evidence": "Physical inspection trench or monitoring data from the 2009 claim (loss adjuster did not carry out any in-depth investigation); full engineer's January 2015 report (only conclusion page provided — confirmed to be restatement of 2013 inspection position rather than a new inspection)",
        "Ombudsman Reasoning": "On balance minor subsidence did occur in 2009 — loss adjuster's local knowledge of the area, slight tapering to the crack, ten years since the extension was built (making settlement a less likely cause) and the willow tree nearby; UKI's 2009 investigation was poor (no monitoring, no in-depth investigation, walked away from repairs because cost fell below excess); UKI failed to explain to Mr A that self-arranging the cosmetic repair and not paying the excess meant no COSA would be issued; this left him unable to sell or move insurer; 2013 engineer report finding no current subsidence does not prove the 2009 cracks were not subsidence — the property stabilised after tree treatment; UKI's offered 'structural engineer's report' was a cut-and-paste of the 2013 report with the word 'historic' inserted — inadequate and unhelpful; COSA is the correct remedy; £500 compensation appropriate for the lasting harm caused by poor handling",
        "Workflow Insight": "When a subsidence claim is accepted but repair costs fall below the policy excess, the insurer must actively explain to the consumer that self-managing cosmetic repairs without insurer-appointed contractors means no Certificate of Structural Adequacy will be issued — and must explain the lasting consequences of that (inability to sell; inability to move insurer); failure to give this warning at the time creates a liability that persists until the insurer provides a COSA or funds an independent report; FOS will not direct removal of a valid subsidence claim from records merely because the original investigation was poor — it will instead require the insurer to remedy the consequences; a substitute document that merely copies a prior inspection report does not satisfy the COSA obligation",
        "AI Rule Candidate": "IF subsidence_claim_accepted AND repair_cost_below_excess AND no_contractor_repairs THEN insurer_must_explain_that_no_COSA_will_be_issued_without_insurer_managed_repairs; IF insurer_fails_to_warn_of_COSA_consequence THEN insurer_must_provide_COSA_or_fund_independent_structural_engineers_report; FOS_cannot_direct_removal_of_valid_subsidence_claim_even_if_investigation_was_inadequate; inadequate_cut_and_paste_document_does_not_satisfy_COSA_obligation; chemical_treatment_and_pruning_of_vegetation_constitutes_remedial_work_even_if_vegetation_not_physically_removed",
        "Source PDF": "DRN8130715.pdf",
    },
    {
        "Case ID": "SUBS-032",
        "FOS Decision ID": "DRN8561608",
        "Insurer Name": "Aviva Insurance Limited",
        "FOS Decision Date": "19 May 2016",
        "Claim Type": "Residential Property Owners insurance (block of leasehold flats) — cover lapsed in 2013 when bank switched the policy to another insurer; Aviva was still dealing with a pre-2013 subsidence claim; new policy arranged with Aviva in 2014 but subsidence cover excluded; Aviva argued ABI domestic subsidence continuation guidance did not apply to D's commercially labelled policy; FOS rejected this and directed reinstatement of subsidence cover",
        "Movement Cause": "Subsidence (cause not specified — pre-2013 claim accepted and being dealt with by Aviva; claim not resolved when policy lapsed and new policy was issued without subsidence cover)",
        "Property Type": "Block of leasehold flats (Residential Property Owners insurance; D is a residents association/limited company; individual flat owners are directors/shareholders of D; building described by Aviva as 'occupied — residential property')",
        "Dispute Type": "Handling / Reinstatement Dispute",
        "Coverage Decision": "Declined — Partial",
        "Rejection Reason": "Aviva: not involved in the 2013 decision to switch to another insurer; ABI domestic subsidence continuation guidance applies only to domestic consumers — D's policy is commercial and the guidance does not apply; FOS exceeded its remit by applying domestic guidance to a commercial policyholder; entitled to decide what cover to offer; break in cover between 2013 and 2014 extinguishes any continuation obligation",
        "Evidence Dispute": "D: Aviva was still dealing with a pre-2013 subsidence claim when cover was restructured; new policy should include subsidence cover as Aviva was the handling insurer. Aviva: ABI agreement only applies to domestic consumers; D's policy is commercial; not responsible for the 2013 lapse; entitled to decide what cover to offer. FOS provisional: the 2014 policy documents did not clearly exclude subsidence except in four specified circumstances; even if intended to exclude, the documents did not achieve this; good industry practice requires insurer dealing with subsidence claim to continue offering subsidence cover. FOS final: building is a residential block; Aviva's own schedules describe it as 'Residential Property Owners Insurance' and 'occupied — residential property'; individual flat owners are directors of D; ABI agreement applies to domestic properties owned and occupied in personal capacity — this effectively fits D; FOS also not restricted to ABI agreement — applies good industry practice regardless; Aviva provided no evidence D failed normal underwriting criteria; break in cover caused by third party does not affect the analysis; Aviva continued providing cover for 2014/15, 2015/16 and 2016/17 (even after provisional decision) — consistent with subsidence cover being appropriate",
        "Outcome Category": "Upheld",
        "Outcome": "Aviva Insurance Limited to continue providing cover for D's property including damage caused by subsidence, as long as D and/or its property meet its normal underwriting criteria other than for the risk of subsidence",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "ABI domestic subsidence continuation guidance applies where the building is effectively residential in nature — it is not disapplied merely because the policyholder is a residents association or limited company if the building is owned by individuals and used as private dwellings; where the insurer's own policy schedules describe the property as 'Residential Property Owners Insurance' and 'occupied — residential property', the policy effectively falls within the domestic guidance regardless of how the insurer categorises it commercially; FOS is not restricted to the terms of the ABI agreement — it applies good industry practice standards; where an insurer is dealing with a subsidence claim, good industry practice requires it to continue offering subsidence cover at renewal subject to normal underwriting criteria other than for subsidence risk; if the insurer asserts the policyholder or property fails its normal underwriting criteria, it must produce evidence of this; a break in cover caused by a third party's administrative decision does not extinguish the subsidence continuation obligation",
        "Missing Evidence": "Aviva's normal underwriting criteria evidence demonstrating that D or its property failed to meet those criteria (Aviva was asked to produce this and did not); clarification on whether Aviva settled the pre-2013 claim by undertaking to carry out or arrange repairs (Aviva did not provide this clarification despite being asked)",
        "Ombudsman Reasoning": "ABI agreement applies to domestic properties owned and occupied in personal capacity — building effectively fits this as it is a residential block owned by individual flat owners who are directors of D; Aviva's own schedules describe it as 'Residential Property Owners Insurance' and 'occupied — residential property'; ABI guidance applies; FOS also applies good industry practice beyond the strict ABI agreement; Aviva provided no evidence D failed normal underwriting criteria; break in cover in 2013 not caused by Aviva's actions and does not affect analysis; Aviva's continued provision of cover for 2014/15, 2015/16 and 2016/17 (even after objecting to provisional decision) is consistent with its obligation to include subsidence cover; complaint upheld",
        "Workflow Insight": "Insurer arguing that ABI domestic subsidence continuation guidance does not apply because a policy is labelled 'commercial' must look beyond the label to the nature of the building and its occupants — where the building is a residential block effectively owned and occupied by individuals (even through a corporate vehicle), the guidance applies; the good industry practice standard applied by FOS (continue to offer subsidence cover when dealing with a claim) is not limited to the precise terms of the ABI agreement and applies beyond ABI members and beyond purely domestic policies; insurer asserting a policyholder fails its normal underwriting criteria must produce evidence — bare assertion is insufficient; a break in cover caused by a third party's administrative decision does not extinguish the continuation obligation",
        "AI Rule Candidate": "ABI_continuation_guidance_applies_to_residential_buildings_held_through_corporate_vehicle_where_individual_owners_occupy_as_personal_residence; IF insurer_labels_policy_commercial_BUT_building_is_effectively_residential THEN ABI_continuation_guidance_still_applies; FOS_applies_good_industry_practice_beyond_strict_scope_of_ABI_agreement_and_beyond_ABI_membership; IF insurer_asserts_policyholder_fails_normal_underwriting_criteria THEN insurer_must_produce_evidence_bare_assertion_is_insufficient; break_in_cover_caused_by_third_party_administrative_decision_does_not_extinguish_subsidence_continuation_obligation",
        "Source PDF": "DRN8561608.pdf",
    },
]


# ---------------------------------------------------------------------------
# Style helpers
# ---------------------------------------------------------------------------
def _row_fill(even: bool) -> PatternFill:
    color = "EDD9C8" if even else "FFFFFF"
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
        repo_root, "knowledge", "case-databases", "Subsidence_Case_Database.xlsx"
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
