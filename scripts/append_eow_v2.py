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
    # {
    #     "Case ID":                  "EOW-016",
    #     "FOS Decision ID":          "DRN-XXXXXXX",
    #     "Insurer Name":             "",
    #     "FOS Decision Date":        "DD Mon YYYY",
    #     "Claim Type":               "",
    #     "Leak Source":              "",
    #     "Property Type":            "",
    #     "Dispute Type":             "",          # controlled vocab
    #     "Coverage Decision":        "",          # controlled vocab
    #     "Rejection Reason":         "",
    #     "Evidence Dispute":         "",
    #     "Outcome Category":         "",          # controlled vocab
    #     "Outcome":                  "",
    #     "Compensation Awarded (£)": 0,           # integer
    #     "Is Core Case":             "",          # controlled vocab
    #     "Key Policy Clause":        "",
    #     "Missing Evidence":         "",
    #     "Ombudsman Reasoning":      "",
    #     "Workflow Insight":         "",
    #     "AI Rule Candidate":        "",
    #     "Source PDF":               "DRN-XXXXXXX.pdf",
    # },
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
