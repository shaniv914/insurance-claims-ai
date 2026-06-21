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
# NEW CASES — Batch 1: SUBS-001 to SUBS-010
# ---------------------------------------------------------------------------
NEW_CASES: list[dict] = [
    {
        "Case ID": "SUBS-001",
        "FOS Decision ID": "DRN0001741",
        "Insurer Name": "UK Insurance Limited",
        "FOS Decision Date": "14 Sep 2015",
        "Claim Type": "Home insurance — subsidence claim for cracks in internal and external walls; insurer accepted external cracks (minor subsidence from defective drains, below excess) but declined internal damage as consolidation of fill material under floor slabs (excluded); insurer then refused to renew policy in 2012",
        "Movement Cause": "Possible rotational subsidence from clay shrinkage aggravated by vegetation (consumer's engineers); insurer attributed internal damage to consolidation of fill material below floor slabs; external cracks confirmed as minor subsidence from defective drains",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Partial",
        "Rejection Reason": "UKI: internal damage caused by consolidation of fill material below floor slabs — excluded under policy; external subsidence cracks below £1,000 subsidence excess; non-renewal decision cited three losses in three years",
        "Evidence Dispute": "UKI loss adjuster Z: external cracks = minor subsidence from defective drains (below excess); internal damage = floor slab consolidation (excluded). Consumer engineers Y and P: Y found movement and cracking caused by subsidence (cause unknown); P found rotational subsidence possibly from clay shrinkage, disagreed with Z's floor slab conclusion; neighbours' CCTV found root damage in adjacent pipes. UKI provided two inconsistent policy documents (2009 and 2012) with different floor slab exclusion wording — correct version not established",
        "Outcome Category": "Upheld",
        "Outcome": "UKI to jointly appoint expert (consumer selects from UKI's list of 3 independent names); expert to investigate and opine on valid subsidence claim; if valid claim found, complete schedule of repairs; reimburse consumer's professional fees if valid claim confirmed; pay £1,500 D&I compensation; calculate cost of buildings cover from 2012 — refund difference plus 8% simple interest if consumer's new policy was more expensive; offer new buildings insurance policy at equivalent terms and cost",
        "Compensation Awarded (£)": 1500,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Floor slab exclusion wording differed between 2009 and 2012 policy documents — correct version for year of claim must be identified and applied; where insurer provides conflicting policy documents the correct version must be established before exclusion can be applied; insurer expected to continue offering buildings cover when subsidence claim is pending — non-renewal when claim unresolved is unfair; consumer selects expert from insurer's shortlist of 3 independent professionals, not insurer's own appointee",
        "Missing Evidence": "Correct policy document for 2011 (year of claim) — UKI provided 2009 and 2012 versions but neither confirmed as applicable",
        "Ombudsman Reasoning": "Consumer's two independent engineers supported possible subsidence rather than floor slab consolidation; damage worsening over 4 years unlikely to be caused by compaction alone in a 40-year-old house; both neighbouring properties suffered subsidence; UKI's floor slab exclusion wording inconsistent across two documents — correct version not established; non-renewal while subsidence claim unresolved is contrary to industry practice; consumer impartiality concerns addressed by insurer providing 3 names for consumer to choose from",
        "Workflow Insight": "Insurer must continue to offer buildings cover while a subsidence claim remains unresolved — non-renewal forces consumer to disclose the claim to new insurers making subsidence cover unobtainable elsewhere; expert appointment process must protect consumer impartiality — insurer nominates 3 independent experts and consumer selects; floor slab exclusion must be verified against the exact policy wording in force at the time of claim — providing multiple conflicting versions is insufficient; where policy wording differs between documents the version most favourable to the consumer may apply if the correct version cannot be established",
        "AI Rule Candidate": "IF subsidence_claim_unresolved AND insurer_refuses_renewal THEN insurer_acted_unfairly; IF insurer_provides_conflicting_policy_documents_for_same_claim_year THEN correct_version_must_be_identified_before_exclusion_can_be_applied; IF floor_slab_exclusion_relied_upon THEN policy_wording_for_year_of_claim_must_be_confirmed; expert_appointment_requires_consumer_choice_from_insurer_nominated_shortlist_of_3_independent_names",
        "Source PDF": "DRN0001741.pdf",
    },
    {
        "Case ID": "SUBS-002",
        "FOS Decision ID": "DRN0017653",
        "Insurer Name": "Aviva Insurance Limited",
        "FOS Decision Date": "11 Jan 2016",
        "Claim Type": "Buildings insurance — complaint about suspension of subsidence cover; consumer had previously obtained court order for £80,000 in full and final settlement of original subsidence claim; made second claim alleging Aviva's drain repairs caused new movement; prior FOS ombudsman dismissed second claim; Aviva then suspended subsidence cover because consumer did not carry out repairs despite receiving £80,000",
        "Movement Cause": "Original subsidence (cause not specified); consumer alleged second movement from Aviva's 2011 drain repairs; Aviva's investigation found second movement was continuation of original claim, not new damage",
        "Property Type": "Residential home",
        "Dispute Type": "Handling / Reinstatement Dispute",
        "Coverage Decision": "Not Applicable",
        "Rejection Reason": "Aviva: consumer received £80,000 in full and final court-ordered settlement but did not carry out repairs; subsidence problem continues; unfair to require Aviva to investigate future claims when underlying cause unresolved; offered reinstatement of cover if repairs done, or cover with £80,000 excess",
        "Evidence Dispute": "Consumer: court order did not specify how £80,000 must be spent; new movement caused by Aviva's drain repairs; wants underpinning or rebuild value. Aviva: investigated new movement; found continuation of original settled claim; spent significant time and resource on investigation proving related to settled problem. Prior FOS ombudsman: dismissed second claim — not appropriate for FOS to reconsider matter already considered by courts. FOS (this decision): prior dismissal binding — cannot revisit; only issue is whether suspension justified",
        "Outcome Category": "Upheld",
        "Outcome": "Aviva to reinstate subsidence cover if consumer carries out the subsidence repairs Aviva considers necessary; alternatively consumer may retain cover with an £80,000 excess applied",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Handling Dispute",
        "Key Policy Clause": "General expectation that insurer continues subsidence cover after a claim; but suspension is justified where court-ordered cash settlement has been paid, consumer has not used it for repairs, and subsidence is continuing unresolved; conditional suspension (reinstate on completion of repairs) is proportionate; prior FOS dismissal of related claim on same subject matter is binding — cannot be reopened by later complaint",
        "Missing Evidence": "Specific schedule of repairs Aviva considered necessary (not detailed in decision)",
        "Ombudsman Reasoning": "General FOS expectation of subsidence cover continuation applies; but £80,000 paid in full and final court settlement — reasonable to expect consumer to use funds for repairs; without repairs subsidence continues and Aviva faces unlimited future investigations into an unsettled original problem; prior FOS ombudsman already dismissed the second claim — cannot revisit; conditional suspension pending consumer completing repairs is proportionate and not a total withdrawal",
        "Workflow Insight": "Insurer obligation to continue subsidence cover can be qualified where original claim was settled by court-ordered cash payment, consumer did not use the funds for repairs, and subsidence remains active; in such cases conditional suspension pending completion of repairs is proportionate; prior FOS ombudsman decision dismissing a related claim on the same subject matter is binding — consumer cannot re-open it via a new complaint",
        "AI Rule Candidate": "IF court_ordered_full_final_settlement_paid AND consumer_did_not_carry_out_repairs AND subsidence_continuing THEN conditional_suspension_of_cover_may_be_justified; IF prior_FOS_decision_dismissed_related_claim THEN same_subject_matter_cannot_be_reopened_by_later_complaint; insurer_obligation_to_continue_cover_qualified_where_settlement_paid_and_repairs_not_made",
        "Source PDF": "DRN0017653.pdf",
    },
    {
        "Case ID": "SUBS-003",
        "FOS Decision ID": "DRN0618226",
        "Insurer Name": "AXA Insurance UK Plc",
        "FOS Decision Date": "14 Aug 2020",
        "Claim Type": "Buildings and contents insurance — subsidence claim made 5 months after inception (Aug 2018); AXA declined claim and voided policy for alleged non-disclosure of previous ground movement at inception; dispute under CIDRA 2012 as to whether consumer made a qualifying misrepresentation",
        "Movement Cause": "Ongoing ground movement — property foundations near bottom of firm clay above poor load-bearing ground; single-storey extensions showed signs of previous movement; prior property on same site demolished; local church demolished for alleged subsidence in 2015",
        "Property Type": "Residential home (built in 1997 by consumer's building company)",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "AXA: consumer stated property free from damage/defect and free from subsidence at inception; investigation revealed signs of previous movement; prior property on site demolished for subsidence; local church demolished for subsidence 2015; Mr W is a builder who would know about ground movement; foundations inadequate for site; would not have offered policy if facts disclosed",
        "Evidence Dispute": "Consumer: did not know prior property was demolished for subsidence — understood it was a foundation issue; did not know about church subsidence; extension movement thought to be normal settlement; proposal questions vague. AXA: consumer built property himself and would know about inadequate foundations; church demolition publicised locally; re-cracking to previous repairs; experienced builder would foresee clay/peat foundation risk. FOS: 'free from subsidence in an area' question undefined — no geographic area or timeframe specified; consumer not expected to know reasons for demolition of neighbouring properties he did not own or inhabit; builder's professional knowledge does not import specific knowledge of this site's subsidence; AXA did not substantiate failure to comply with building regulations",
        "Outcome Category": "Upheld",
        "Outcome": "AXA to honour the claim as if the policy had not been voided; fair compensation likely exceeds FOS £160,000 award limit — FOS recommends AXA pay the balance above £160,000 (recommendation only, not binding determination)",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "CIDRA 2012 — consumer must take reasonable care not to misrepresent; vague questions ('free from subsidence in an area' without defining area or timeframe) reduce consumer culpability for incorrect answers; consumer not expected to know reasons for demolition of neighbouring properties they did not own or inhabit; builder's occupation imports professional knowledge of building regulations but not specific area subsidence history; for careless misrepresentation insurer must show it would not have offered the policy at all — not merely on different terms",
        "Missing Evidence": "Specific evidence that prior property was demolished for subsidence (AXA's assertion not independently proven); evidence of failure to comply with building regulations when property was constructed (not provided)",
        "Ombudsman Reasoning": "Proposal questions vague — 'area free from subsidence' defines neither area nor timeframe; consumer's belief that prior property demolished for foundation issues not unreasonable given he did not live there; church demolition — consumer denied knowledge and AXA's information board image was unclear; builder's occupation does not import knowledge of this specific site's subsidence history; AXA did not substantiate building regulation failure; not satisfied qualifying misrepresentation established; policy voiding unfair",
        "Workflow Insight": "Subsidence non-disclosure cases under CIDRA: vague proposal questions without geographic and temporal definitions substantially weaken the insurer's case; consumer's professional background (e.g. builder) only imports knowledge directly relevant to that profession, not general area subsidence history; insurer must provide specific evidence that neighbouring demolitions were for subsidence and that consumer would have known the reason — general assertion insufficient; for careless misrepresentation and policy avoidance the threshold is that insurer would not have offered the policy at all",
        "AI Rule Candidate": "IF subsidence_question_does_not_define_geographic_area_or_timeframe THEN question_is_vague_reducing_consumer_culpability; IF consumer_alleged_to_know_subsidence_from_professional_background THEN insurer_must_show_direct_specific_knowledge_of_this_site; IF prior_property_demolished_and_insurer_relies_on_consumer_knowing_reason THEN specific_evidence_of_consumer_knowledge_required; careless_misrepresentation_requires_insurer_to_show_policy_would_not_have_been_offered_at_all",
        "Source PDF": "DRN0618226.pdf",
    },
    {
        "Case ID": "SUBS-004",
        "FOS Decision ID": "DRN1210158",
        "Insurer Name": "Royal & Sun Alliance Insurance Plc",
        "FOS Decision Date": "20 Jan 2017",
        "Claim Type": "Home insurance (all-risks policy) — cracks in external walls caused by water leaking from underground pipes softening the soil and causing subsidence; RSA discovered prior 2001 claim (handled as subsidence) not disclosed at inception; retrospectively applied subsidence exclusion endorsement; declined new claim on basis it was a subsidence claim excluded by that endorsement",
        "Movement Cause": "Water leakage from cracked clay drain pipes and mains water pipe washing away and softening ground under building; softened ground caused building to subside, cracking external walls",
        "Property Type": "Residential home",
        "Dispute Type": "Peril Classification Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "RSA: prior 2001 claim was a subsidence claim — consumer failed to disclose it at inception; RSA entitled to retrospectively apply subsidence exclusion endorsement; new claim caused by subsidence (not escape of water); subsidence endorsement also excludes escape of water, so claim fails on both bases",
        "Evidence Dispute": "Consumer: 2001 claim also involved water-induced damage and was not treated as subsidence; took reasonable care at inception; new damage is escape of water, not subsidence. RSA: broker and RSA correspondence in 2001 referred to claim as subsidence; loss adjuster's subsidence department handled it; endorsement excludes escape of water from drains along with subsidence. FOS: balance of probabilities — 2001 claim probably handled as subsidence; retrospective exclusion endorsement justified given non-disclosure; BUT chain of events analysis — escape of water caused soil softening which caused subsidence (water first, not subsidence first); proximate cause of current damage is escape of water; endorsement is a subsidence exclusion — its reference to 'escape of water from drains' relates only to loss of rent, alternative accommodation and third-party liability, not to primary damage; on all-risks policy escape of water is not excluded",
        "Outcome Category": "Upheld",
        "Outcome": "RSA to settle claim for damage to home caused by escape of water",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Proximate cause rule in chain-of-events cases: where escape of water caused soil softening which caused subsidence, proximate cause is escape of water not subsidence; if subsidence had cracked pipes first, proximate cause would be subsidence; insurer must correctly identify the causal sequence; retrospective application of subsidence exclusion is justified where prior claim was handled as subsidence and consumer failed to disclose it at inception; subsidence exclusion endorsement referencing 'escape of water from drains' in context relates to ancillary items (rent, accommodation, liability), not to primary structural damage claim",
        "Missing Evidence": "Detailed records from 2001 insurer (could not be obtained due to passage of time)",
        "Ombudsman Reasoning": "2001 claim handled as subsidence on balance of probabilities (correspondence, loss adjuster's subsidence department); consumer carelessly failed to disclose it — retrospective endorsement justified; however endorsement is a subsidence exclusion — its escape of water reference relates to consequential losses (rent, accommodation, liability to third parties), not to primary structural damage; current claim is escape of water that caused subsidence: proximate cause = escape of water; on all-risks policy this is covered; RSA's reliance on endorsement to decline the claim is unreasonable",
        "Workflow Insight": "Chain of causation is critical in combined subsidence/escape of water claims: identify which event occurred first — escape of water softening ground (EOW proximate cause) or subsidence cracking pipes (subsidence proximate cause); insurer must correctly sequence events before applying a subsidence exclusion; subsidence exclusion endorsements containing incidental references to escape of water must be read carefully — such references typically relate only to consequential items (loss of rent, liability) and do not exclude primary structural damage from escape of water; prior claim categorisation by loss adjusters can establish it was a subsidence claim for disclosure purposes",
        "AI Rule Candidate": "IF escape_of_water_caused_soil_softening_which_caused_subsidence THEN proximate_cause_is_escape_of_water_not_subsidence; IF subsidence_caused_cracked_pipes_which_then_caused_escape_of_water THEN proximate_cause_is_subsidence; subsidence_exclusion_endorsement_reference_to_escape_of_water_relates_only_to_ancillary_consequential_items_unless_expressly_stated_otherwise; retrospective_subsidence_exclusion_justified_where_prior_claim_handled_as_subsidence_and_not_disclosed_at_inception",
        "Source PDF": "DRN1210158.pdf",
    },
    {
        "Case ID": "SUBS-005",
        "FOS Decision ID": "DRN1933952",
        "Insurer Name": "Royal & Sun Alliance Insurance Plc",
        "FOS Decision Date": "19 Oct 2018",
        "Claim Type": "Home insurance — block of three leasehold flats; subsidence damage from nearby tree roots damaging drains; RSA declined claim and removed subsidence cover alleging non-disclosure of (1) tree within 7 metres of property at inception, and (2) prior 2006 subsidence claim",
        "Movement Cause": "Tree roots damaging drains causing subsidence to the building",
        "Property Type": "Leasehold flats (three flats in one residential building; standard household buildings policy)",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "RSA: Ms M told broker no trees within 7 metres at inception (incorrect — tree was within 5 metres); also told broker property had not moved in 40 years (incorrect — 2006 subsidence claim existed); if known about tree, would not have offered subsidence cover; subsidence cover removed and database entries updated accordingly",
        "Evidence Dispute": "Consumer: pre-call questionnaire did not ask about trees; statement of fact issued with policy makes no reference to the tree question; tree was over 7 metres from Ms M's own flat (though closer to other two flats); 2006 claim outside 5-year claims question window; RSA's statement of fact question about subsidence was confusingly worded (answering 'No' to 'is your home free from subsidence' means admitting to subsidence signs). RSA: specific oral question about trees in entire building's vicinity asked in phone call; Ms M said no trees at all in vicinity; 2006 claim = misrepresentation about no movement in 40 years. FOS: oral tree question not recorded in statement of fact — statement of fact is the document on which policy is based; three flat owners have equal responsibility for disclosures — all should have been consulted; subsidence question on statement of fact was confusingly phrased; 2006 claim was beyond the 5-year general claims window — not specifically asked about prior subsidence claims; RSA's own acceptance confirmed property had declared subsidence history",
        "Outcome Category": "Upheld",
        "Outcome": "RSA to reinstate subsidence cover and settle the subsidence claim on reinstated policy terms; remove all database entries recording that subsidence cover was withdrawn",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Multi-party leasehold block policies require all named policyholders to be informed of and agree to material questions — a single phone call with one leaseholder is insufficient; if a material question is asked orally but not reflected in the statement of fact, reliance on the oral answer is weakened; subsidence question wording must be clear — a yes/no question about being 'free from subsidence' can produce misleading answers and cannot reliably support avoidance; prior claims outside the specific question window (5 years) need not be disclosed unless a specific question about subsidence history is asked",
        "Missing Evidence": "Pre-call questionnaire (not seen by FOS — Ms M said it did not mention trees); complete recording of the broker phone call",
        "Ombudsman Reasoning": "Tree question asked orally but not on statement of fact — statement of fact said to be the basis of the policy; three flat owners all had equal responsibility for disclosures but only Ms M was consulted on the tree question; subsidence question on statement of fact confusingly phrased — 'No' answer means consumer is admitting to signs of subsidence; 2006 claim — only a general 5-year claims question was asked; no specific question about prior subsidence claims was asked; RSA's own acceptance record reflected declared subsidence history; not satisfied clear misrepresentation established",
        "Workflow Insight": "Leasehold block policies with multiple named policyholders require all policyholders to be informed of and confirm material questions — telephone questioning of only one flat owner is procedurally inadequate; if an oral question about trees/proximity is asked in a phone call but is not reproduced in the statement of fact issued, the insurer's ability to rely on the oral answer is significantly weakened; subsidence question design matters — a separate clearly worded question about prior subsidence history is needed to capture claims outside the standard 5-year window",
        "AI Rule Candidate": "IF multi_party_leasehold_block_policy AND material_question_asked_only_to_one_policyholder THEN all_policyholders_should_have_been_consulted; IF oral_question_about_trees_asked_but_not_recorded_on_statement_of_fact THEN insurer_reliance_on_oral_answer_is_weakened; IF prior_subsidence_claim_outside_5_year_claims_question_window AND no_specific_subsidence_history_question_asked THEN non_disclosure_of_that_claim_is_not_a_qualifying_misrepresentation; ambiguous_subsidence_question_wording_reduces_insurer_ability_to_establish_qualifying_misrepresentation",
        "Source PDF": "DRN1933952.pdf",
    },
    {
        "Case ID": "SUBS-006",
        "FOS Decision ID": "DRN2093738",
        "Insurer Name": "Society of Lloyd's",
        "FOS Decision Date": "14 Jan 2016",
        "Claim Type": "Buildings insurance — insurer removed subsidence cover at renewal after two subsidence claims (2001: underpinning, claim closed 2010; 2011: new shrubs caused cracking, repairs done); property stable at time of renewal; insurer cited claims history, cost and high-risk area data to justify removal",
        "Movement Cause": "First claim (2001): unspecified subsidence, underpinning required; second claim (2011): new shrubs planted post-first-claim caused new subsidence cracking",
        "Property Type": "Residential home",
        "Dispute Type": "Handling / Reinstatement Dispute",
        "Coverage Decision": "Accepted",
        "Rejection Reason": "Lloyd's: claims history (two subsidence claims); risk data showed property in high-risk subsidence area; commercial decision; cost of claims; property currently stable but risk remains high",
        "Evidence Dispute": "Consumer: property stable after both claims; cannot obtain subsidence cover elsewhere because of claims history; proposed independent structural survey to evidence stability. Lloyd's: claims history justifies commercial decision to withdraw; offered structural survey by consumer's choice of engineer to present to another insurer; admitted it was a 'one-off' case and had no internal guidelines for policies with more than one subsidence claim. FOS: Lloyd's could not demonstrate withdrawal aligned with any written acceptance criteria; only provided individual letters; ABI Domestic Subsidence Agreement recommends continuation of cover; consumer had confirmed inability to obtain cover elsewhere",
        "Outcome Category": "Upheld",
        "Outcome": "Lloyd's to reinstate subsidence cover from July 2014 and charge appropriate premium from reinstatement date; pay £350 compensation for trouble and upset caused",
        "Compensation Awarded (£)": 350,
        "Is Core Case": "Yes",
        "Key Policy Clause": "ABI Domestic Subsidence Agreement: insurer should continue providing subsidence cover after a claim because policyholders will have difficulty obtaining cover elsewhere; FOS treats this as good industry practice binding on all insurers regardless of whether they are ABI signatories; insurer's commercial judgement to withdraw cover remains subject to FOS review for fairness; FOS may consider and overturn commercial judgement decisions where doing so is appropriate; insurer citing commercial judgement must still demonstrate alignment with its own acceptance criteria",
        "Missing Evidence": "Lloyd's written acceptance criteria for subsidence cover withdrawal (not provided — Lloyd's admitted it had no internal guidelines for multiple-claim policies)",
        "Ombudsman Reasoning": "ABI Domestic Subsidence Agreement requires continuation of cover — FOS treats as good industry practice regardless of ABI membership; consumer could not obtain subsidence cover elsewhere; Lloyd's could not show withdrawal aligned with written acceptance criteria — only provided individual letters; no internal guidelines existed for this scenario; property was stable at time of renewal; withdrawal likely made property unsaleable; £350 appropriate for the trouble and upset caused",
        "Workflow Insight": "Insurer withdrawing subsidence cover at renewal must: (1) demonstrate alignment with written acceptance criteria; (2) consider the ABI Domestic Subsidence Agreement / FOS continuation standard; (3) show withdrawal is proportionate given property stability and consumer's inability to obtain cover elsewhere; citing claims history and high-risk area without documented criteria is insufficient; FOS has and will exercise discretion to review commercial judgement decisions in subsidence cases where consumer harm is evident",
        "AI Rule Candidate": "IF insurer_withdraws_subsidence_cover_at_renewal THEN must_demonstrate_alignment_with_written_acceptance_criteria; ABI_domestic_subsidence_agreement_continuation_principle_applies_regardless_of_ABI_membership; IF property_stable_AND_consumer_cannot_obtain_cover_elsewhere THEN withdrawal_of_subsidence_cover_likely_unfair; FOS_may_review_commercial_judgement_decisions_in_subsidence_cases_where_consumer_harm_evident",
        "Source PDF": "DRN2093738.pdf",
    },
    {
        "Case ID": "SUBS-007",
        "FOS Decision ID": "DRN-2213774",
        "Insurer Name": "Advantage Insurance Company Limited",
        "FOS Decision Date": "26 Oct 2020",
        "Claim Type": "Home insurance — subsidence claim accepted and settled by Advantage during year 1 (Apr 2018–Apr 2019); broker unable to obtain renewal from Advantage or any panel insurer; Advantage stated refusal was due to new flood risk data showing property in high-risk flood area (not due to subsidence claim); consumer could not obtain subsidence cover elsewhere",
        "Movement Cause": "Subsidence (cause not specified in decision — accepted and settled in year 1)",
        "Property Type": "Residential home",
        "Dispute Type": "Handling / Reinstatement Dispute",
        "Coverage Decision": "Accepted",
        "Rejection Reason": "Advantage: non-renewal reason is new flood risk data purchased during policy term showing property in high-risk flood area — deemed uninsurable under its underwriting criteria; not refusing because of subsidence claim",
        "Evidence Dispute": "Consumer: unable to find any insurer willing to cover subsidence because of prior claim. FOS investigator: Advantage is ABI member — bound by Domestic Subsidence Agreement to continue cover. Advantage: genuine and distinct reason (flood risk) justifies refusal; forcing renewal would require it to cover an unacceptable risk. FOS: flood risk existed at inception — Advantage just lacked the data then; if Advantage had the data in 2018 it would not have offered cover; consumer could then have obtained alternative policy where the accepting insurer would be bound by ABI agreement to continue subsidence cover; consumer should not lose subsidence protection because Advantage accepted a risk it now regrets; flood risk alone not so fundamentally unreasonable as to override ABI obligations",
        "Outcome Category": "Upheld",
        "Outcome": "Advantage to adhere to ABI Domestic Subsidence Agreement and renew policy from April 2021; cover any subsidence claims arising between now and renewal; pay £100 D&I compensation",
        "Compensation Awarded (£)": 100,
        "Is Core Case": "Yes",
        "Key Policy Clause": "ABI Domestic Subsidence Agreement — ABI member insurer that has dealt with a subsidence claim must continue to offer cover on renewal; insurer can only depart from this obligation where the additional risk beyond subsidence is fundamentally unreasonable (e.g. fraud), not merely commercially unwanted (e.g. post-inception discovery of flood risk data); where insurer accepted a risk at inception it later regrets, the consumer should not lose subsidence continuation protection as a consequence",
        "Missing Evidence": "N/A — Advantage's flood risk data accepted as genuine; dispute concerned whether that data was sufficient to override ABI obligations",
        "Ombudsman Reasoning": "Advantage is ABI member — bound by ABI Domestic Subsidence Agreement; flood risk was pre-existing and unchanged — Advantage simply obtained better data after inception; had Advantage had that data in 2018 it would not have offered cover and consumer would have placed policy elsewhere; the alternative insurer would then be bound by ABI agreement; consumer should not lose subsidence protection because of Advantage's underwriting error; flood risk not so fundamental as to justify departure from ABI agreement — only fraud or equivalently egregious risk would justify departure",
        "Workflow Insight": "ABI member insurers cannot escape Domestic Subsidence Agreement renewal obligations by pointing to a separate uninsurable risk discovered post-inception — only a fundamentally unreasonable additional risk (e.g. fraud) justifies departure; the causation test is: where would the consumer have been if the insurer had acted correctly at inception; if consumer could have obtained alternative cover, the alternative insurer would be bound by the ABI agreement and consumer would retain subsidence protection; post-inception discovery of existing risk data does not override ABI obligations",
        "AI Rule Candidate": "IF ABI_member_insurer_settled_subsidence_claim AND refuses_renewal_citing_separate_risk THEN still_bound_by_ABI_domestic_subsidence_agreement; IF insurer_accepted_risk_at_inception_it_later_regrets THEN consumer_should_not_lose_subsidence_continuation_protection; only_fraud_or_fundamentally_unreasonable_risk_justifies_departure_from_ABI_subsidence_renewal_obligation; flood_risk_data_discovered_post_inception_does_not_override_ABI_renewal_obligation",
        "Source PDF": "DRN-2213774.pdf",
    },
    {
        "Case ID": "SUBS-008",
        "FOS Decision ID": "DRN2337317",
        "Insurer Name": "Royal & Sun Alliance Insurance Plc",
        "FOS Decision Date": "8 Feb 2016",
        "Claim Type": "Buildings insurance for a block of five domestic flats managed by management company S Ltd (insured under commercial policy); original subsidence claim dealt with by RSA; further cracking appeared and new claim opened — independent engineers found cracking was thermal movement, not subsidence; RSA then announced withdrawal of subsidence cover at next renewal",
        "Movement Cause": "Original subsidence (cause not detailed); subsequent cracking attributed to thermal movement by two independent engineers — confirmed not subsidence",
        "Property Type": "Block of five domestic flats (insured under commercial policy by management company; FOS treated as domestic property)",
        "Dispute Type": "Handling / Reinstatement Dispute",
        "Coverage Decision": "Accepted",
        "Rejection Reason": "RSA: claims history, recent cracking, property in high-risk subsidence area justify withdrawal; S Ltd holds a commercial policy and RSA can withdraw commercial cover at its discretion; if S does not accept policy without subsidence cover RSA will offer no cover at all",
        "Evidence Dispute": "S Ltd: small group of flat-owning individuals — not a commercial entity; two independent expert engineers confirmed no further subsidence; flat owners face mortgage and saleability concerns without subsidence cover. RSA: commercial policy; commercial decision; entitled to withdraw cover. FOS: domestic flats regardless of policy label — substance over form; RSA could not show withdrawal aligned with its written acceptance criteria — provided only individual letters; both RSA's and independent experts confirmed no current subsidence; RSA's threat to withdraw all cover if S insists on subsidence cover is also unreasonable",
        "Outcome Category": "Upheld",
        "Outcome": "RSA to continue offering subsidence cover as part of the policy; no monetary award",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Block of domestic flats insured under commercial policy is treated as a domestic property for FOS subsidence continuation purposes — substance of insured risk determines applicable rules, not policy label; insurer withdrawing subsidence cover must demonstrate alignment with its own written acceptance criteria — vague letters citing risk area and claims history are insufficient; ABI / FOS good practice on continuation applies regardless of commercial policy label; insurer threatening to withdraw all cover unless consumer accepts removal of subsidence cover is a separate and additional unreasonable act",
        "Missing Evidence": "RSA's written acceptance criteria for withdrawing subsidence cover from this type of policy (not provided)",
        "Ombudsman Reasoning": "Domestic flats — FOS applies domestic subsidence continuation principles regardless of commercial policy label; RSA could not demonstrate withdrawal aligned with written criteria — only individual letters; two experts (RSA's and independent) confirmed property not suffering from subsidence; RSA's threat to offer no cover at all if subsidence cover insisted upon is an additional unreasonable act not shown to be within any reasonable criteria; continuation of cover required",
        "Workflow Insight": "The commercial policy label does not exempt an insurer from FOS domestic subsidence continuation expectations where the substance of the insured risk is residential; insurer must produce written acceptance criteria showing withdrawal was criteria-based — anecdotal letters or verbal assertions do not suffice; the threat to withdraw all cover unless the consumer accepts exclusion of subsidence is a separate unreasonable act that warrants FOS intervention; current confirmed absence of subsidence reinforces that withdrawal is unjustified",
        "AI Rule Candidate": "IF insured_property_is_domestic_flats AND policy_label_is_commercial THEN FOS_applies_domestic_subsidence_continuation_principles; IF insurer_withdraws_subsidence_cover THEN must_provide_written_acceptance_criteria_demonstrating_criteria_based_decision; IF insurer_threatens_total_cover_withdrawal_if_subsidence_cover_insisted_upon THEN additional_unreasonable_conduct_established; confirmed_absence_of_current_subsidence_by_expert_engineers_weighs_strongly_against_cover_withdrawal",
        "Source PDF": "DRN2337317.pdf",
    },
    {
        "Case ID": "SUBS-009",
        "FOS Decision ID": "DRN2707923",
        "Insurer Name": "AXA Insurance UK Plc",
        "FOS Decision Date": "2013 (est.)",
        "Claim Type": "Home insurance — subsidence claim accepted by AXA; dispute is about whether AXA must fund preventative drainage works required to stabilise the ground before structural repairs can be carried out; AXA argued preventative works are not covered by the policy; consumer's surveyor argued further investigation needed before any works",
        "Movement Cause": "Excessive moisture in ground from multiple sources: surface water pooling due to defective/uneven paving; water percolating from neighbouring land; possibly high inherent water table; combination caused soil saturation and subsidence",
        "Property Type": "Residential home",
        "Dispute Type": "Handling / Reinstatement Dispute",
        "Coverage Decision": "Accepted",
        "Rejection Reason": "AXA: policy only requires repair of structural damage; preventative works (land drain installation, patio re-profiling) are not covered; AXA not required to investigate or remove cause of subsidence; damage was 'inevitable' due to water percolation — not a fortuitous event; will only carry out structural repairs after consumer funds and completes preventative works",
        "Evidence Dispute": "Consumer's surveyor: further investigation needed before works begin; AXA's diagnosis insufficient; AXA not entitled to require consumer to fund unidentified works. AXA: established cause as excessive ground water; drainage installation will mitigate it; re-profiling is alternative or supplementary option; not required to identify which of three potential causes is primary. FOS: where removal of cause is required for effective repair, insurer must arrange and pay for it; land drain most likely effective mitigation; patio re-profiling also needed; AXA should carry out both before structural repairs; 'inevitability' argument rejected — major cause (neighbour's percolating water) was unknown to consumer",
        "Outcome Category": "Upheld",
        "Outcome": "AXA to: install land drainage to right-hand side of property and re-profile patio area at its own expense; carry out all necessary repairs and preventative work to deal with the subsidence damage; pay £150 D&I compensation; reimburse consumer's surveyor costs incurred in respect of the claim",
        "Compensation Awarded (£)": 150,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Where removal of the cause of subsidence is required for an effective and durable repair, the insurer must arrange and fund those preventative works — regardless of whether the policy's strict wording covers only repair of structural damage; this is analogous to insurer liability for underpinning; the 'inevitability' of damage does not relieve the insurer of this obligation; insurer is responsible for preventative works that are necessary to ensure insured repairs will last; consumer's professional fees for surveyors/engineers engaged in the claim may be recoverable where complaint is upheld",
        "Missing Evidence": "Specific identification of which of three potential causes was primary (FOS found this unnecessary — directed both repair schemes as complementary)",
        "Ombudsman Reasoning": "AXA accepted subsidence liability. Where cause removal is required for effective repair, insurer must fund it — equivalent to underpinning obligation which no insurer disputes. AXA's 'inevitability' argument rejected: the major cause (neighbour's percolating water) was unknown to the consumer and thus not foreseeable. Land drain installation most likely to mitigate cause; patio re-profiling also necessary to prevent surface water pooling. AXA should carry out both before structural repairs. Consumer's surveyor costs should be reimbursed.",
        "Workflow Insight": "Insurer liability for preventative works in subsidence claims: where the only way to carry out an effective structural repair is first to stabilise the ground (e.g. install drainage, remove tree, underpin), the insurer must fund this even if the policy's literal wording covers only structural damage repair; this is equivalent to underpinning liability; where multiple complementary repair schemes are proposed FOS may direct all of them if each addresses a distinct contributing cause; consumer professional fees for surveyors/engineers engaged to investigate or manage the claim may be recoverable if complaint is upheld",
        "AI Rule Candidate": "IF cause_removal_required_for_effective_durable_structural_repair THEN insurer_must_fund_preventative_works; IF insurer_proposes_alternative_repair_schemes_AND_they_are_complementary THEN FOS_may_direct_both_to_be_carried_out; inevitability_of_damage_does_not_relieve_insurer_of_obligation_to_fund_cause_removal_works; IF complaint_upheld AND consumer_incurred_surveyor_costs_investigating_claim THEN surveyor_costs_may_be_recoverable",
        "Source PDF": "DRN2707923.pdf",
    },
    {
        "Case ID": "SUBS-010",
        "FOS Decision ID": "DRN-2807339",
        "Insurer Name": "Aviva Insurance Limited",
        "FOS Decision Date": "7 Jun 2021",
        "Claim Type": "Buildings insurance — tree caused clay soil shrinkage and subsidence; tree removed, movement stopped; Aviva cash settled at £11,777.35; consumer disputed settlement was insufficient to cover all subsidence damage including damage to three external walls; Aviva contended significant wall damage was historic/pre-existing, not from current subsidence",
        "Movement Cause": "Nearby tree causing clay soil shrinkage; tree removed by Aviva; monitoring confirmed movement stopped post-removal",
        "Property Type": "Residential home",
        "Dispute Type": "Handling / Reinstatement Dispute",
        "Coverage Decision": "Accepted — Disputed Settlement",
        "Rejection Reason": "Aviva: cash settlement of £11,777.35 covers subsidence damage to main house; external wall damage is largely historic/pre-existing (lack of lateral restraint, historic settlement, old plaster, render failure) unrelated to current subsidence; subsidence expert found limited subsidence damage only on front wall; settlement is fair",
        "Evidence Dispute": "Consumer's engineer Mr L (2014): ongoing crack damage to all three external walls caused by recent and ongoing movement from tree; front wall and boundary wall in poor condition; evidence of historic repairs (steel ties) but also recent ongoing crack damage. Consumer's second engineer Mr W: movement due to soil shrinkage at corner; Aviva's settlement too low for complete wall remediation. Aviva's subsidence expert: identified historic damage to walls; only specifically commented on subsidence damage to front wall (limited); did not specifically assess subsidence damage to boundary wall or right-hand side wall. FOS: Mr L's evidence that all three walls showed ongoing crack damage from tree movement is persuasive; Aviva's expert's failure to assess two of the three walls means Aviva cannot credibly argue those walls have no subsidence damage; where insurer cannot distinguish insured (subsidence) from uninsured (historic/pre-existing) damage, FOS approach is to require insurer to cover all damage",
        "Outcome Category": "Upheld",
        "Outcome": "Aviva to increase cash settlement to cover: (1) full cost of putting right subsidence damage to main house per original schedule of works; (2) full cost of putting right all damage to the three external walls (including pre-existing where it cannot be distinguished from subsidence damage)",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Where insurer cannot distinguish between insured (subsidence) damage and pre-existing/uninsured damage at the same location, FOS requires insurer to cover total remediation cost; insurer's expert must specifically assess all affected structural elements for subsidence damage — failure to assess specific walls means insurer cannot later deny subsidence damage to those walls; cash settlement in lieu of repair must equal actual cost to consumer of having the work done — it cannot be discounted below consumer repair costs on basis of insurer's own contractor rates",
        "Missing Evidence": "Aviva's subsidence expert's specific assessment of boundary wall and right-hand side wall for subsidence damage (expert only commented on front wall)",
        "Ombudsman Reasoning": "Aviva's expert confirmed subsidence from tree. Consumer's engineer Mr L assessed all three external walls and found ongoing crack damage from tree movement — this evidence is persuasive. Aviva's expert failed to specifically assess the boundary wall and right-hand side wall for subsidence damage, so Aviva cannot credibly claim those walls were unaffected. Aviva could not distinguish between subsidence damage and historic pre-existing damage — FOS standard approach: require insurer to cover all damage. Aviva chose cash settlement rather than repair — settlement must be sufficient for consumer to actually commission the repairs.",
        "Workflow Insight": "Insurer instructing a subsidence expert must ensure the expert specifically assesses all affected structural elements — failure to assess specific walls or features means the insurer has no evidential basis to deny subsidence damage to those elements; where mixed subsidence/pre-existing damage is present and the insurer cannot segregate them, FOS requires the insurer to fund total remediation of all damage at that location; cash settlement in lieu of repair must be calibrated to the actual cost the consumer will incur to commission the required work — not the insurer's own contractor cost",
        "AI Rule Candidate": "IF insurer_expert_fails_to_specifically_assess_structural_element THEN insurer_cannot_deny_subsidence_damage_to_that_element; IF insurer_cannot_distinguish_subsidence_damage_from_pre_existing_damage THEN FOS_requires_insurer_to_fund_total_remediation_of_all_damage; IF insurer_cash_settles_instead_of_repairing THEN settlement_must_equal_actual_cost_to_consumer_of_commissioning_repairs; cash_settlement_cannot_be_discounted_below_consumer_repair_cost_on_basis_of_insurer_contractor_rates",
        "Source PDF": "DRN-2807339.pdf",
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
