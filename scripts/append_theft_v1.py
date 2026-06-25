"""
Standard append script for Theft Case Database — Schema v1 (21 columns).
Column 6 is "Entry / Theft Method" (how the theft was committed or item came to be stolen).

Usage
-----
1. Read the source PDF(s) and extract the fields listed in NEW_CASES below.
2. Add one dict per case to NEW_CASES following the extraction rules.
3. Run from the repo root:
       py scripts/append_theft_v1.py

Appends NEW_CASES rows to:
    knowledge/case-databases/Theft_Case_Database.xlsx

===========================================================================
FIELD EXTRACTION RULES
===========================================================================

Case ID             : Format THEFT-NNN (zero-padded to 3 digits)
FOS Decision ID     : DRN-XXXXXXX or DRNXXXXXXX as printed in the PDF
Insurer Name        : Formal registered name from the FOS decision
FOS Decision Date   : DD Mon YYYY — accept-or-reject deadline in final paragraph;
                      use "Not stated in document" if deadline not printed
Claim Type          : Policy type, physical incident and nature of dispute in one sentence
Entry / Theft Method: How the theft was committed or how items came to be stolen
                      e.g. "Forced entry — front door kicked in"
                           "Theft by invited persons — items taken during house move"
                           "Snatch theft from person — bag grabbed in street"
                           "Motor vehicle theft — keys left in ignition"
Property Type       : "Residential home" / "Motor vehicle (personal)" /
                      "Commercial vehicle" / "Personal gadget / mobile phone" /
                      "Personal travel insurance" etc.
Dispute Type        : Controlled vocab (7 values)
Coverage Decision   : Controlled vocab (5 values)
Rejection Reason    : Insurer's stated reason for declining or disputing
Evidence Dispute    : What evidence each party relied on
Outcome Category    : Controlled vocab (4 values)
Outcome             : Full FOS remedy instructions
Compensation Awarded (£) : Integer — D&I only; 0 if none
Is Core Case        : Controlled vocab (5 values)
                      Use "No — Commercial" for motor, travel, gadget or commercial
                      policies that are outside home/contents insurance scope
Key Policy Clause   : Policy wording or FOS/FCA principle applied
Missing Evidence    : Evidence that was absent and affected the outcome
Ombudsman Reasoning : How the ombudsman weighed the evidence
Workflow Insight    : Operational rule for the claims workflow
AI Rule Candidate   : Machine-evaluable rule for the rules engine
Source PDF          : Filename only (e.g. DRN0600921.pdf)
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
    "Entry / Theft Method",
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
# NEW CASES — Batch 1: THEFT-001 to THEFT-010
# Note: 7 of 10 cases are non-home-insurance (motor / travel / gadget /
# commercial van) — flagged Is Core Case = "No — Commercial" or
# "No — Broker Dispute" as appropriate.
# ---------------------------------------------------------------------------
NEW_CASES: list[dict] = [
    {
        "Case ID": "THEFT-001",
        "FOS Decision ID": "DRN0600921",
        "Insurer Name": "Zenith Insurance plc",
        "FOS Decision Date": "8 Feb 2018",
        "Claim Type": "Motor (car) insurance — vehicle stolen overnight 20-21 June 2015, recovered damaged; Zenith appointed insurance investigator whose report identified multiple inconsistencies in Mr A's account; theft claim rejected as not credible; complaint also covers simultaneous accident claim (rear-end collision 20 June 2015) in which Zenith settled with third-party driver",
        "Entry / Theft Method": "Motor vehicle theft — car stolen overnight from road (evening 20 June to morning 21 June 2015); vehicle recovered damaged; specific entry or key method not established; insurance investigator found inconsistencies in complainant's account undermining claim credibility",
        "Property Type": "Motor vehicle (personal car insurance)",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Insurance investigator's detailed report highlighted multiple inconsistencies in Mr A's evidence and raised credibility concerns; policy entitled Zenith to reject claim under fraud provisions; policy obligation on Mr A to disclose 'correct and complete information to the best of your knowledge and belief' — failure may result in refusal of claim",
        "Evidence Dispute": "Mr A: theft was genuine; Zenith wrongly rejected claim and accused him of being untruthful; Zenith failed to obtain witness statements he provided for accident claim. Zenith: insurance investigator's detailed report showed multiple inconsistencies making theft claim not credible; solicitor advised accident defence prospects below 50% so Zenith entitled to settle. FOS: investigator's report and all evidence reviewed; inconsistencies and credibility concerns sufficient to justify rejection; Zenith made reasonable but unsuccessful efforts to obtain witness statements; Zenith within rights to settle accident claim per policy terms and solicitor advice placing prospects below 50%.",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — theft claim rejection fair and reasonable in light of investigator's report and evidence inconsistencies; accident claim settlement by Zenith also fair and reasonable within policy terms",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "Policy fraud provisions entitle insurer to reject a claim where evidence inconsistencies undermine the complainant's credibility; policy obligation to provide 'correct and complete information' is valid basis for rejection where investigator's report reveals material inconsistencies; insurer's right to settle third-party accident claims where solicitor advises defence prospects below 50% ('reasonable prospects' threshold of 51%+); solicitor's professional assessment of sub-50% prospects is reasonable basis for settlement without policyholder consent",
        "Missing Evidence": "Specific itemised inconsistencies from the investigator's report (FOS summarised as 'a number of inconsistencies' without itemising); witness statements that Mr A put forward for the accident claim (Zenith made efforts but was unable to obtain them)",
        "Ombudsman Reasoning": "Insurance investigator appointed; detailed report prepared after gathering evidence and taking a statement from Mr A; report highlighted multiple inconsistencies and raised credibility concerns; FOS satisfied Zenith took all available evidence and the investigator's report into account before rejecting; decision fair and reasonable and in accordance with policy terms; for accident: solicitor advised defence prospects below 50%; policy required 'reasonable prospects' (51%+); Zenith made efforts to obtain witness statements without success; Zenith within right to settle",
        "Workflow Insight": "Where an insurance investigator's detailed report identifies material inconsistencies in the claimant's account and raises credibility concerns, an insurer may reject a theft claim under fraud provisions without proving fraud to a criminal standard; if an insurer makes reasonable efforts to obtain evidence but is unable to do so, this weakens the claimant's case and the insurer is not at fault; insurer may exercise its right to settle a third-party accident claim without the policyholder's consent where solicitor advice places prospects below the policy's reasonable prospects threshold",
        "AI Rule Candidate": "IF investigation_report_identifies_material_inconsistencies AND raises_credibility_concerns THEN insurer_entitled_to_reject_theft_claim_under_fraud_provisions; IF solicitor_advice_places_defence_prospects_below_50_percent THEN insurer_entitled_to_settle_accident_claim_without_policyholder_consent; IF insurer_makes_reasonable_efforts_to_obtain_evidence_and_fails THEN failure_to_obtain_weakens_claimants_case_not_insurers",
        "Source PDF": "DRN0600921.pdf",
    },
    {
        "Case ID": "THEFT-002",
        "FOS Decision ID": "DRN0709326",
        "Insurer Name": "Legal & General Insurance Limited",
        "FOS Decision Date": "26 May 2019",
        "Claim Type": "Home insurance — February 2018 burglary (brick thrown through back door window; items stolen following day); L&G voided policy from inception and declined theft claim for non-disclosure of prior accidental damage claim on jointly-held home policy with husband Mr B (claim made 7 months before current policy inception)",
        "Entry / Theft Method": "Forced entry — brick thrown through back door window; property burgled the following day with items stolen",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "L&G voided policy from inception: Mrs N failed to disclose at inception a May 2017 accidental damage claim (laptop, £1,299) made under a jointly-held home policy with husband Mr B — only 7 months before she took out the disputed policy in December 2017; broker's application question asked about any claims in the last five years; L&G's underwriting acceptance criteria required no claims within previous 12 months; L&G would not have issued the policy if the claim had been disclosed",
        "Evidence Dispute": "Mrs N: laptop claim was made by her partner under a policy for a different property; disputed policy in her sole name so no prior claims to disclose; Mr B is not her husband. L&G: claims database showed Mrs N as named joint policyholder on property A schedule with Mr B (December 2016 – December 2017); claim made under that joint policy May 2017; schedule showed both as 'married'; phone call recordings obtained; Mrs N confirmed Mr B is her husband in later email to FOS adjudicator. FOS: joint policy schedule confirmed Mrs N as joint policyholder; claim clearly within five-year question scope; in 15 February 2018 call Mrs N did not disclose that Mr B was her husband despite opportunity; L&G underwriter's acceptance criteria (no claims in prior 12 months) confirmed by documentary evidence; L&G would not have issued policy — voidance justified.",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — L&G entitled to void policy from inception for careless misrepresentation; policy voided from inception means no cover exists at time of theft; theft claim fairly declined; L&G confirmed premiums returned to Mrs N",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Consumer must take 'reasonable care' not to make a misrepresentation when taking out insurance (Consumer Insurance (Disclosure and Representations) Act 2012); insurer may void from inception and refuse all claims if it shows (a) a misrepresentation was made and (b) it would not have entered into the contract on the same terms had correct information been given; misrepresentation obligation extends to claims under any policy where consumer is a named policyholder — including jointly-held policies — not only sole-name policies; voiding from inception means there is no cover at the time of any claim; broker application question asking about 'claims in the last five years' for the consumer 'or anyone who normally lives with you' is sufficiently clear to require disclosure of claims under jointly-held policies with a spouse; insurer must produce documentary evidence of underwriting criteria to justify voidance",
        "Missing Evidence": "Direct evidence of what was said during the policy sale (sold through a broker online — statement of facts showed 'no previous claims disclosed'); content of phone call recordings (FOS ombudsman obtained and listened to these; Mrs N did not volunteer during calls that Mr B was her husband despite opportunities to do so)",
        "Ombudsman Reasoning": "Joint policy schedule shows Mrs N as joint policyholder on property A with Mr B (married status confirmed); laptop claim made under that joint policy 7 months before she took out new policy; broker asked clear five-year claims question; Mrs N had opportunity in 15 February 2018 call to clarify Mr B was her husband and the claim jointly hers — she did not do so; L&G underwriter criteria (no claims in prior 12 months) confirmed by underwriter note seen by FOS; policy voided from inception means no cover; theft claim fairly declined; premiums returned",
        "Workflow Insight": "Where a consumer is a joint policyholder on another policy at the time of a prior claim, that claim must be disclosed when answering any question about prior claims — even if the consumer's new policy is in their sole name; insurer is entitled to treat a claim as undisclosed if the consumer omits it and the insurer can show it would not have issued the policy; insurer must produce evidence of its underwriting criteria — bare assertion is insufficient; voiding from inception removes all cover retrospectively so any intervening claim (including the triggering claim) is automatically excluded",
        "AI Rule Candidate": "IF consumer_is_joint_policyholder_on_prior_policy AND claim_made_under_that_joint_policy THEN claim_must_be_disclosed_on_new_policy_even_if_new_policy_is_sole_name; IF consumer_fails_to_disclose_prior_claim AND insurer_proves_it_would_not_have_issued_policy THEN insurer_may_void_from_inception_and_decline_all_claims; IF policy_voided_from_inception THEN no_cover_exists_for_any_claim_in_voided_period; insurer_must_produce_evidence_of_underwriting_criteria_not_merely_assert_it",
        "Source PDF": "DRN0709326.pdf",
    },
    {
        "Case ID": "THEFT-003",
        "FOS Decision ID": "DRN0999356",
        "Insurer Name": "Legal and General Insurance Limited",
        "FOS Decision Date": "10 Aug 2015",
        "Claim Type": "Household insurance — Miss S took out L&G household policy online March 2013; later made accidental damage claim for damaged mobile phone which triggered review of claims history; L&G found (1) wrong date given for prior theft claim — stated 16 months before but actually 5 months before application — and (2) undisclosed prior accidental damage claim 11 months before (no payment received as below excess); L&G cancelled policy and refunded premiums; Miss S seeks removal of cancellation record from insurance databases",
        "Entry / Theft Method": "Not specified in decision — prior theft claim on previous household policy; circumstances of that theft not described; this decision concerns the non-disclosure of the prior claim history rather than the mechanics of the theft itself",
        "Property Type": "Residential home (rented — tenant)",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "(1) Miss S gave incorrect date for prior theft claim — said it occurred 16 months before application but it actually occurred 5 months before; L&G did not offer cover to tenants with one or more claim in the last three years; (2) Miss S failed to disclose prior accidental damage claim (11 months before; below excess; no payment received) — L&G asked about all claims not just settled ones; if correct information had been given, Miss S would have been refused cover under both the tenant criteria and the (incorrectly applied) owner-occupier criteria",
        "Evidence Dispute": "Miss S: medical conditions affecting memory caused her to give the wrong date for the theft claim; did not consider the accidental damage claim as a 'claim' since no payment received — she treated it as only an enquiry; accidental damage claim was not 'settled' so need not be declared; adjudicator initially found L&G system incorrectly applied owner-occupier criteria to Miss S as a tenant. L&G: claim questions were clear; cancellation was correct. FOS: even accepting memory issues, Miss S remembered the theft claim (she declared it, just misdated it) and had documents she could have checked for the correct date; accidental damage claim was recorded by insurer, she had letters with the claim reference number; question asked about 'any claims' not 'settled claims'; even under the erroneously applied owner-occupier criteria (max one claim in last 12 months) both claims were within 12 months — refusal would have resulted regardless of the system error.",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — Miss S acted carelessly when providing claims history; L&G would have refused cover if correct information had been given even under the incorrectly applied owner-occupier criteria; cancellation was justified; FOS cannot require deletion of valid and accurate cancellation record",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Consumer must take 'reasonable care' not to make a misrepresentation; medical conditions may affect memory but do not excuse carelessness where the consumer is aware of a claim and has supporting documents they could have checked; a claim is a 'claim' for disclosure purposes whether or not payment was received — an enquiry recorded by the insurer and assigned a claim reference number is a claim; insurers ask about 'any claims' not only 'settled' or 'paid' claims; where an insurer's system applies incorrect underwriting criteria but the correct criteria would also have led to refusal, the insurer may still cancel the policy; FOS cannot direct deletion of an accurate cancellation record",
        "Missing Evidence": "Sufficient medical evidence to establish that Miss S's memory condition specifically caused the misrepresentation in this instance (FOS acknowledged memory problems but found evidence insufficient to attribute the specific misdating to those conditions); full L&G underwriting guide for tenant vs owner-occupier criteria (adjudicator reviewed this — showed tenant criteria more restrictive than owner-occupier)",
        "Ombudsman Reasoning": "Miss S knew she had made the theft claim (she declared it but misdated it); she had documents from the claim she could have checked before providing the date; medical conditions do not explain why she could not check documents; accidental damage claim — she had letters referring to 'claim' and a claim reference number; L&G's question was clear ('any claims'); L&G system incorrectly applied owner-occupier criteria to a tenant; but both claims were within 12 months — even owner-occupier criteria (max one claim in last 12 months) would have led to refusal because there were two; cancellation was entitled; FOS cannot require deletion of accurate cancellation record",
        "Workflow Insight": "When reviewing a misrepresentation complaint, first determine whether the consumer took 'reasonable care'; if not, determine what would have happened if correct information had been given; if the insurer would have refused cover on correct information, cancellation is justified even if the insurer's system made a separate error in applying underwriting criteria; a claim that resulted in no payment (because below excess) is still a 'claim' that must be disclosed; an insurer cannot be required to delete an accurate cancellation record merely because the consumer finds it difficult to obtain new insurance as a result",
        "AI Rule Candidate": "IF consumer_declares_claim_but_gives_wrong_date AND consumer_had_supporting_documents_available THEN consumer_has_not_taken_reasonable_care; IF claim_recorded_by_insurer_and_assigned_reference_number THEN claim_must_be_disclosed_regardless_of_whether_payment_was_received; IF insurer_system_applies_incorrect_criteria_but_correct_criteria_also_leads_to_refusal THEN insurer_still_entitled_to_cancel; FOS_cannot_direct_deletion_of_accurate_cancellation_record; question_about_any_claims_includes_below_excess_nil_payment_claims",
        "Source PDF": "DRN0999356.pdf",
    },
    {
        "Case ID": "THEFT-004",
        "FOS Decision ID": "DRN-1313333",
        "Insurer Name": "Lloyds Bank General Insurance Limited",
        "FOS Decision Date": "6 Mar 2020",
        "Claim Type": "Contents insurance — August 2019 theft claim for items missing after house move; Mr R believed items stolen by family members invited to help him unpack; Lloyds declined as policy excluded theft by persons invited into the home; Mr R had previously added personal belongings cover following an earlier declined accidental damage claim for a phone",
        "Entry / Theft Method": "Theft by invited persons — items disappeared during home move; Mr R believed family members lawfully invited to help him unpack were responsible; alternatively removal company (but items would have been taken in April 2019 before policy began in May 2019)",
        "Property Type": "Residential home",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Policy theft cover explicitly excludes theft by anybody invited into the home; Mr R's own account stated he believed invited family members were responsible; policy documents clearly set out this exclusion; alternatively, if removal company responsible, theft would have occurred in April 2019 before policy began in May 2019 — outside policy period",
        "Evidence Dispute": "Mr R: as a victim of theft he should be covered; adding personal belongings cover meant he would be covered for any contents claim; policy not adequately explained during sale; he should be covered regardless of who committed the theft. Lloyds: policy theft section explicitly excludes theft by any person invited into the home — clearly set out in policy documents; sales script used in branch shows clear distinction between accidental damage and personal belongings cover; adviser's recollection given approximately one week after sale is credible; Mr R chose not to take personal belongings cover as it would not cover his mobility scooter; items taken by removal company would predate policy start.",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — policy exclusion for theft by invited persons is clearly communicated in policy documents and through the sale process; adding personal belongings cover does not override the theft exclusion for invited persons; items taken before policy inception are outside the policy period",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Insurers are entitled to limit theft cover to specific circumstances and exclude theft by invited persons — such limitation is valid if clearly stated in policy documents and communicated at sale; 'theft by invited persons' exclusion applies where the policyholder voluntarily allowed the alleged thief access to the premises; adding a supplementary section (e.g. personal belongings cover) does not override exclusions in other sections of the policy; theft occurring before policy inception is not covered regardless of the type of cover purchased; where a policy is sold in a branch without a recording, FOS may rely on an adviser's contemporaneous recollection if given shortly after the sale",
        "Missing Evidence": "Recording of the in-branch policy sale conversation (none available — sale was in person; FOS relied on adviser's recollection given approximately one week after sale)",
        "Ombudsman Reasoning": "Policy theft section clearly states cover does not apply to theft by persons invited into the home; Mr R invited his family members in — no doubt they were invited; Lloyds acted fairly in declining; sales script reviewed — makes clear distinction between accidental damage and personal belongings cover; adviser's recollection (~1 week post-sale) persuasive; personal belongings cover does not override theft-by-invited-persons exclusion; items taken by removal company would predate policy start in May 2019 (Mr R moved in April 2019) — Lloyds not liable, previous insurer could be approached; no insurance policy covers everything — insurers entitled to define covered circumstances in policy documents",
        "Workflow Insight": "A contents policy with a 'theft by invited persons' exclusion will not respond where the alleged thief was present with the policyholder's permission — the exclusion is determinative; when selling a policy and a consumer adds a supplementary section, the adviser should clarify that this does not override exclusions in other sections; where a theft precedes policy inception, the claim falls to the previous insurer regardless of cover purchased; in-branch sales without recordings may rely on contemporaneous adviser recollection where the recollection is given shortly after the sale",
        "AI Rule Candidate": "IF alleged_thief_was_invited_onto_premises THEN theft_by_invited_persons_exclusion_applies_and_claim_declined; IF theft_occurred_before_policy_inception_date THEN claim_is_outside_policy_period_regardless_of_cover_type; adding_supplementary_cover_section_does_not_override_exclusions_in_other_sections; IF_no_sale_recording AND adviser_recollection_given_shortly_after_sale THEN adviser_recollection_may_be_relied_upon",
        "Source PDF": "DRN-1313333.pdf",
    },
    {
        "Case ID": "THEFT-005",
        "FOS Decision ID": "DRN-1447440",
        "Insurer Name": "Assurant General Insurance Limited",
        "FOS Decision Date": "10 Aug 2020",
        "Claim Type": "Mobile phone and gadget insurance (bank account tech pack) — mid-November 2019 claim for two stolen phones (Ms L and her partner); Assurant declined both as the annual policy limit of 4 loss/theft claims per Tech Pack had already been reached; dispute about misleading information given during claim registration process that led Ms L to believe one phone would be covered",
        "Entry / Theft Method": "Mobile phone theft — circumstances of theft not specified; two phones stolen (not lost) in mid-November 2019",
        "Property Type": "Personal gadget / mobile phone insurance (tech pack via bank account — not home/contents)",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Policy limits loss/theft claims to 4 per Tech Pack per 12 months; Ms L had already made 4 successful claims (3 in December 2018, 1 in January 2019) before the November 2019 claim; annual limit was reached; no further loss/theft claims permitted until anniversary of the first claim",
        "Evidence Dispute": "Ms L: Assurant initially indicated both phones would be assessed and later declined both without adequate warning; £100 compensation insufficient to replace two phones. Assurant: Ms L was told her claim would be 'assessed' not 'accepted'; £50 compensation offered for misleading information; claim correctly declined as annual limit reached. FOS: policy terms and policy summary both clearly state 4 loss/theft claim limit; limit was reached; decline was correct; advisor set up the claim without being able to immediately check limit — but Ms L endured several days of expectation that a claim would succeed; Assurant acknowledges it could have been clearer; £100 appropriate compensation.",
        "Outcome Category": "Upheld in Part",
        "Outcome": "Assurant General Insurance Limited to pay Ms L £100 compensation for misleading information that caused distress and inconvenience; claim decline itself was fair and reasonable as annual limit had been reached",
        "Compensation Awarded (£)": 100,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "Annual claim limits for specific claim types (loss/theft) are enforceable where clearly stated in both policy terms and policy summary; where an advisor registers a claim without being able to immediately confirm eligibility, the insurer must clarify ineligibility as soon as established — failure to do so promptly gives rise to a compensation obligation for distress and inconvenience; compensation of £100 is appropriate where the consumer endured several days of expectation that a claim would be considered when it would not",
        "Missing Evidence": "Records of the initial claim registration call (to determine exactly what was represented to Ms L at outset about claim eligibility)",
        "Ombudsman Reasoning": "Policy terms and summary clearly set out 4 loss/theft claim limit per Tech Pack per 12 months; Ms L had made 4 successful claims in the previous 12 months — limit reached; decline was correct; advisor setting up the claim could not immediately know the limit was reached — Ms L likely told both phones would be processed at outset; Ms L had expectation for several days that at least one claim would succeed; Assurant acknowledges it could have been clearer; £100 appropriate for distress and inconvenience; £50 offered by Assurant was insufficient",
        "Workflow Insight": "When an advisor registers a claim without immediately verifying whether the annual claim limit has been reached, they must check the limit at the earliest opportunity and communicate ineligibility promptly — failure to communicate promptly creates a legitimate compensation liability; annual claim limits that are clearly stated in both policy terms and policy summary are enforceable without modification",
        "AI Rule Candidate": "IF annual_claim_limit_reached THEN insurer_must_communicate_ineligibility_at_earliest_opportunity_not_allow_expectation_to_build; IF consumer_endures_days_of_expectation_due_to_delayed_notification_of_ineligibility THEN compensation_approximately_£100_is_appropriate; annual_claim_limits_clearly_stated_in_policy_terms_and_summary_are_enforceable",
        "Source PDF": "DRN-1447440.pdf",
    },
    {
        "Case ID": "THEFT-006",
        "FOS Decision ID": "DRN-1744488",
        "Insurer Name": "Vitality Health Limited",
        "FOS Decision Date": "19 Jun 2020",
        "Claim Type": "Travel insurance — 2019 mobile phone theft while on holiday; phone stolen from complainant's hand while he had fallen asleep in a nightclub; Vitality declined as phone was 'unattended' (policy defined as 'not in full view of, and not in a position to prevent unauthorised interference') and reasonable care had not been taken",
        "Entry / Theft Method": "Snatch theft from person while asleep — phone taken from complainant's hand while he was asleep in a nightclub on holiday; opportunistic theft exploiting unconscious state",
        "Property Type": "Personal travel insurance (not home/contents)",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Policy excludes theft where belongings are left 'unattended' — defined as 'when you are not in full view of, and not in a position to prevent unauthorised interference with, your belongings'; and where 'reasonable care has not been taken to avoid the theft or loss'; Vitality: falling asleep in a busy public place (nightclub) with phone in hand does not satisfy being in full view of it or being in a position to prevent interference, and does not constitute reasonable care",
        "Evidence Dispute": "Mr C: phone was not 'unattended', 'unsecured', 'forgotten' or 'left behind' — he was physically holding it when stolen; he only fell asleep 'momentarily'; police report confirms genuine theft; policy wording should not apply where owner is physically present and holding the item. Vitality: policy definition of 'unattended' includes being asleep — cannot be 'in full view' or 'in a position to prevent interference' while asleep in a busy public place; reasonable care not taken. FOS: agreed that Mr C physically held phone but he was asleep — could not see phone and not in a position to prevent interference; policy definition of 'unattended' satisfied; not unreasonable for Vitality to conclude reasonable care not taken in a busy public place while asleep.",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — Vitality acted fairly in declining claim; policy definition of 'unattended' was satisfied as complainant was asleep and unable to see or protect the phone",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "The policy definition of 'unattended' (not in full view of, and not in a position to prevent unauthorised interference) applies when the policyholder is asleep — physical possession of the item does not amount to being in full view or able to prevent interference if the owner is unconscious; the 'reasonable care' exclusion applies to falling asleep in a busy public place with a valuable in hand; the fact that a theft is genuine and supported by a police report does not override a valid circumstance-based exclusion; locked accommodation exception explicitly carved out in policy terms distinguishes the sleeping-in-hotel-room scenario",
        "Missing Evidence": "Precise duration of the sleep episode ('momentarily' per Mr C) — if clearly established as extremely brief this might affect the analysis but FOS did not find it materially distinguished the case from the policy's plain terms",
        "Ombudsman Reasoning": "Circumstances of theft not in dispute — phone taken from complainant's hand while asleep in a nightclub on holiday; being asleep means complainant was not in full view of phone and not in a position to prevent unauthorised interference; policy definition of 'unattended' therefore met; Vitality not unreasonable in concluding reasonable care not taken in a busy public place while asleep; genuine theft confirmed by police report does not override circumstance-based exclusion; hypothetical of items stolen from locked hotel room is explicitly covered by the policy — distinguishable from this case",
        "Workflow Insight": "Travel insurers applying an 'unattended' exclusion should apply the policy's own definition — where the definition requires the policyholder to be in full view of and able to prevent interference with items, the exclusion is met when the policyholder is asleep regardless of physical possession; a police report confirming a theft is genuine does not override circumstance-based exclusions; 'reasonable care' is fact-specific and applies where the circumstance (asleep in a busy public place) was preventable",
        "AI Rule Candidate": "IF policyholder_asleep_when_theft_occurred THEN policy_definition_of_unattended_satisfied_regardless_of_physical_possession; IF unattended_exclusion_met THEN claim_declined_even_if_police_report_confirms_genuine_theft; falling_asleep_in_busy_public_place_with_valuables_satisfies_reasonable_care_exclusion; locked_accommodation_exception_applies_only_to_items_secured_in_locked_hotel_room",
        "Source PDF": "DRN-1744488.pdf",
    },
    {
        "Case ID": "THEFT-007",
        "FOS Decision ID": "DRN-2188041",
        "Insurer Name": "Royal & Sun Alliance Insurance Plc",
        "FOS Decision Date": "11 Nov 2020",
        "Claim Type": "Commercial van insurance — company van stolen while parked; claim settled by RSA; dispute about RSA recording the director (named driver, last to use the van before theft) on the Claims Underwriting Exchange (CUE) as 'driver' for the claim, resulting in increased personal motor insurance costs for the director",
        "Entry / Theft Method": "Motor vehicle theft — company van stolen while parked at company/director's address; specific method of entry not detailed; dispute concerns CUE recording only, not the claim circumstances",
        "Property Type": "Commercial vehicle (van insurance for limited company)",
        "Dispute Type": "Claim Recording / Administrative Dispute",
        "Coverage Decision": "Accepted",
        "Rejection Reason": "Not applicable — claim was accepted and settled; dispute concerns how the settled claim was recorded on the CUE database",
        "Evidence Dispute": "Company C: Mr B (director, named driver) was recorded on CUE as 'driver' for the theft claim which increased his personal motor insurance costs; Mr B was not driving the van at the time of theft; vehicle owned by the company not Mr B; policy terms do not state that named drivers will be individually recorded on CUE. RSA: C told RSA that Mr B was the last driver before the theft (main driver, parked at his address, only other named driver had not used it); standard industry practice to record last driver for theft claims on CUE; policy terms explain incident information passed to various registers — sufficient notice. FOS: standard practice to record last driver for theft on CUE; recording does not mean driver is blamed; RSA reasonably inferred Mr B was last driver from the claim notification call; policy terms adequate — not required to specifically state the driver's name would be included.",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — RSA acted in accordance with standard industry practice in recording Mr B as driver on CUE; not required to do anything differently",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "It is standard industry practice for insurers to record on CUE the driver (or last driver for theft claims) of a vehicle at the time of an incident; policy terms notifying that 'incident information will be passed to various registers' provides adequate notice of CUE recording without requiring explicit statement that the driver's name will be included; being recorded on CUE as driver does not mean the driver is held at fault — it records who was using the vehicle; the fact that the policyholder is a company rather than the named driver individually does not change the standard CUE recording practice",
        "Missing Evidence": "Absolute certainty about who was the last driver of the van (RSA inferred from notification call — Mr B was main driver, parked at his address, other named driver confirmed he had not used it — but RSA was not directly told and acknowledged it could not be known for absolute certain)",
        "Ombudsman Reasoning": "Standard industry practice for insurers to record last driver for theft claims on CUE; recording does not mean driver at fault or liable; from the claim notification call RSA had sufficient information to reasonably conclude Mr B was last driver; policy terms explain incident information passed to registers — sufficient without specifying CUE or that driver's name would form part of that information; cannot be known with absolute certainty that Mr B was last driver but it was most likely; RSA acted fairly and reasonably",
        "Workflow Insight": "Named drivers on commercial or company vehicle policies should be aware that being recorded as the last driver of a stolen vehicle on CUE is standard practice — even where the policyholder is a company, the named driver's personal insurance records may be affected; insurers need not specifically state in policy terms that the driver's name will form part of CUE records — general notice about incident information being passed to registers is sufficient",
        "AI Rule Candidate": "IF named_driver_is_last_driver_of_stolen_vehicle THEN insurer_entitled_to_record_driver_on_CUE_as_standard_practice; CUE_recording_of_last_driver_does_not_imply_fault_or_liability; policy_terms_notifying_that_incident_information_passed_to_registers_provides_adequate_notice_of_CUE_recording; IF insurer_reasonably_concludes_person_was_last_driver_from_claim_notification THEN recording_that_person_on_CUE_is_fair",
        "Source PDF": "DRN-2188041.pdf",
    },
    {
        "Case ID": "THEFT-008",
        "FOS Decision ID": "DRN2547216",
        "Insurer Name": "AXA Insurance UK Plc",
        "FOS Decision Date": "Not stated in document",
        "Claim Type": "Motor (car) insurance — Mr L took out motor policy via price comparison website; AXA discovered he had not disclosed two prior theft claims within the previous three years; AXA voided policy from inception and applied a £25 administration fee; Mr L disputes the voidance and the administration fee",
        "Entry / Theft Method": "Motor vehicle theft (prior undisclosed claims) — two prior vehicle theft claims not disclosed at inception; specific circumstances of those prior thefts not described in this decision",
        "Property Type": "Motor vehicle (personal car insurance)",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Mr L failed to disclose two theft claims within the previous three years when applying for motor insurance via price comparison website; application question was clear; AXA's underwriting policy: would not offer insurance if more than one theft claim was disclosed; AXA would not have issued the policy if the claims had been declared; policy terms entitled voidance from inception and application of a £25 administration fee for failure to declare claims",
        "Evidence Dispute": "Mr L: he entered all information correctly including details of previous claims. AXA: when Mr L entered his details on the price comparison website (auto-transferred to AXA) there was no mention of two claims; AXA was not aware of them; schedule showed no claims or accidents in last three years which Mr L did not query; AXA would not have offered insurance if claims disclosed; £25 admin fee reasonable for tasks carried out. FOS: satisfied AXA asked a clear question; not persuaded Mr L disclosed the two theft claims at outset; Mr L did not query the schedule showing no claims; AXA's underwriting policy not to accept applicants with more than one theft claim confirmed; voidance fair; admin fee per policy terms and reasonable.",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — AXA acted fairly in voiding motor insurance policy from inception; £25 administration fee appropriate per policy terms; Mr L failed to disclose two prior theft claims",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "Motor insurer's policy terms entitling it to void from inception where policyholder fails to declare all motoring accidents or claims in the previous three years are enforceable; consumer should check the resulting policy schedule and query any inaccuracies — failure to notice that the schedule stated 'no claims' when two claims had occurred is a factor against the consumer; an insurer's underwriting rule not to offer cover where more than one theft claim was disclosed is a legitimate criterion justifying voidance if undisclosed claims would have triggered it; administration fee charged per policy terms on voidance is reasonable",
        "Missing Evidence": "Records of what Mr L actually entered on the price comparison website (AXA stated no mention of two claims when details auto-transferred from comparison site to AXA)",
        "Ombudsman Reasoning": "AXA asked a clear question about previous claims; Mr L's details entered on price comparison website and auto-transferred to AXA — no mention of two theft claims; Mr L did not question the schedule showing no claims or accidents in last three years; AXA's underwriting policy not to offer cover if more than one theft claim disclosed; confirmed AXA would not have issued policy; voidance fair; administration fee per policy terms and reasonable",
        "Workflow Insight": "Policyholders using price comparison websites must verify that auto-transferred details accurately reflect their claims history — failure to check the resulting schedule for accuracy (especially where it shows 'no claims' when claims exist) weakens any argument that correct information was provided; an insurer's stated underwriting policy not to accept applicants with more than one theft claim is a legitimate basis for voidance where that criterion was not met due to non-disclosure",
        "AI Rule Candidate": "IF application_auto_populated_from_comparison_site AND schedule_shows_no_claims THEN consumer_should_verify_accuracy_and_failure_to_query_supports_non_disclosure_finding; IF insurer_underwriting_policy_excludes_applicants_with_more_than_one_theft_claim AND claims_not_disclosed THEN insurer_entitled_to_void_from_inception; administration_fee_per_policy_terms_on_voidance_is_enforceable_and_reasonable",
        "Source PDF": "DRN2547216.pdf",
    },
    {
        "Case ID": "THEFT-009",
        "FOS Decision ID": "DRN2725005",
        "Insurer Name": "BISL Limited",
        "FOS Decision Date": "Not stated in document",
        "Claim Type": "Motor insurance (broker administered) — BISL sent two contradictory non-payment cancellation letters with different response deadlines (12 April 2012 letter gave 18 May 2012 deadline; 16 April 2012 letter cancelled from 25 April 2012); Mr D relied on first letter and contacted BISL on 4 May 2012 to pay; policy already cancelled; subsequent vehicle theft occurred with no cover in place; complaint that policy cancelled in error",
        "Entry / Theft Method": "Motor vehicle theft — vehicle stolen while motor insurance policy was erroneously lapsed due to broker administrative error; circumstances of theft not described in decision",
        "Property Type": "Motor vehicle (personal motor insurance administered by broker BISL)",
        "Dispute Type": "Broker Conduct Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "BISL: second cancellation letter (16 April 2012) superseded first letter (12 April 2012); no reason to assume second letter was not received; policy correctly cancelled on 25 April 2012 for non-payment of direct debit; theft claim appeared suspicious",
        "Evidence Dispute": "Mr D: he did not receive the second cancellation letter; he relied on the first letter which gave 18 May 2012 as the deadline; he contacted BISL on 4 May 2012 genuinely believing he was within the deadline; policy should not have been cancelled when he contacted to pay. BISL: two letters sent; second letter superseded first; policy correctly cancelled; theft claim suspicious. FOS: BISL made administrative error by sending two letters with two different response deadlines; from Mr D's 4 May 2012 call it is clear he genuinely believed he had until 18 May — consistent with receiving only the first letter; consumer entitled to rely on first deadline received; BISL had indicated it would accept responsibility if cancellation was its error; broker must stand in insurer's shoes where its error causes a coverage gap.",
        "Outcome Category": "Upheld",
        "Outcome": "BISL Limited required to: deal with the theft claim; provide Mr D with a letter for future insurers confirming the policy was cancelled in error by BISL; pay Mr D £100 compensation for distress and inconvenience",
        "Compensation Awarded (£)": 100,
        "Is Core Case": "No — Broker Dispute",
        "Key Policy Clause": "Where a broker sends two contradictory letters with different response deadlines, the consumer is entitled to rely on the first deadline received — the broker cannot assume the later letter was received and acted upon; where a broker's administrative error leads to a policy being cancelled when it should not have been, the broker must 'stand in the insurer's shoes' and deal with any claim that arises during the period when cover should have been in place; FOS standard approach: broker who causes a policy gap through its error must take responsibility for claims arising in that gap",
        "Missing Evidence": "Proof of delivery of the second cancellation letter (BISL assumed it was received but could not prove delivery; Mr D denied receiving it — critical to determining which deadline he was entitled to rely on)",
        "Ombudsman Reasoning": "BISL sent two letters with contradictory response deadlines — administrative error; Mr D's 4 May 2012 call confirms he genuinely believed he had until 18 May (consistent with receiving only the first letter); FOS usual approach where broker error leads to policy not being in place is for broker to stand in insurer's shoes and deal with the claim; BISL had previously indicated it would accept responsibility if the cancellation was its error; BISL content to deal with claim; £100 compensation for distress and inconvenience",
        "Workflow Insight": "Brokers must ensure that when sending follow-up cancellation notices they do not create contradictory deadlines — if two letters go out with different deadlines, the consumer is entitled to rely on the first received; where a broker's error causes a coverage gap, the broker must fund the claim rather than leaving the consumer without recourse; broker's advance conditional commitment to accept responsibility if error is confirmed will be held to that commitment by FOS",
        "AI Rule Candidate": "IF broker_sends_two_contradictory_cancellation_notices_with_different_deadlines THEN consumer_entitled_to_rely_on_first_deadline_received; IF broker_error_causes_policy_lapse THEN broker_must_stand_in_insurers_shoes_and_pay_claim_arising_in_lapsed_period; IF broker_commits_to_accept_responsibility_if_error_confirmed THEN FOS_will_hold_broker_to_that_commitment",
        "Source PDF": "DRN2725005.pdf",
    },
    {
        "Case ID": "THEFT-010",
        "FOS Decision ID": "DRN2794102",
        "Insurer Name": "Markerstudy Insurance Company Limited",
        "FOS Decision Date": "30 Oct 2015",
        "Claim Type": "Commercial motor insurance — Company R's van stolen during employee delivery stop in February 2015; Markerstudy rejected claim alleging driver fraud (driver initially told police he left keys in van; later told Markerstudy he believed he may have been pickpocketed); FOS found no evidence of fraud and directed Markerstudy to reconsider claim",
        "Entry / Theft Method": "Motor vehicle theft during delivery stop — van stolen from outside store while employee driver was making a delivery; driver initially reported keys left in van to police; later reported he believed keys were pickpocketed before the theft; vehicle access method uncertain",
        "Property Type": "Commercial vehicle (motor/van insurance for limited company)",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Markerstudy: driver first reported to police that he left keys in van when stolen — constitutes a fraudulent misrepresentation inconsistent with the later account; inconsistency between police account (keys left in van) and later account given to Markerstudy (pickpocketed) undermines claim credibility; fraud alleged and claim rejected on that basis",
        "Evidence Dispute": "Company R / Mr E: driver was in shock when reporting to police and gave his initial impression; upon reflection and when calmer he believed he may have been pickpocketed before the van was stolen; driver recalled using his keys before entering the store, offered a description of the potential thief, and Markerstudy's own investigation confirmed he locked the van at other delivery stops. Markerstudy: the first account given to police is more reliable; inconsistency between accounts is indicative of fraud; store owner made general comments about the theft (though declined to provide a named statement or be identified). FOS: both accounts (keys left in van / pickpocketed) are possible and plausible; driver likely in shock when giving initial account; changed account does not amount to fraud without stronger corroborating evidence; store owner's unnamed and unsubstantiated comments cannot be relied upon; insurer's own investigation confirmed good vehicle security practices at other delivery stops.",
        "Outcome Category": "Upheld",
        "Outcome": "Markerstudy Insurance Company Limited required to reconsider the theft claim in line with remaining policy terms and conditions; pay £150 compensation for distress and inconvenience caused by the fraud allegation",
        "Compensation Awarded (£)": 150,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "Fraud allegation requires sufficient corroborating evidence — an insurer cannot decline a theft claim as fraudulent solely on the basis of an inconsistency between an initial police report made in shock and a later more considered account; a driver giving an initial impression at the scene and revising that account when calmer is not fraud without corroborating evidence; unsubstantiated comments from a witness who refuses to be named or provide a statement cannot form the basis of a fraud finding; insurer's own investigation evidence supporting the claimant's general credibility (e.g. good vehicle security practice at other delivery stops) must be weighed in the claimant's favour",
        "Missing Evidence": "Named statement from the store owner (declined to provide one — making his general comments inadmissible for fraud purposes); CCTV footage from the store or surrounding area that could corroborate or contradict either account of the key/pickpocket scenario",
        "Ombudsman Reasoning": "Markerstudy had reason to doubt the account but had no evidence of fraud; both accounts (keys left in van / keys pickpocketed) are possible and plausible; driver in state of shock when reporting to police — initial account reflects his first impression of what most likely happened; upon reflection and when calmer he believed he was pickpocketed and offered a description of the potential thief; Markerstudy's own investigation showed driver locked van at other delivery stops — supports general credibility; store owner's comments unsubstantiated (unnamed, no statement) — cannot be relied upon; inconsistency between accounts does not amount to fraud; must distinguish between an inconsistency and deliberate fraud; £150 compensation for being falsely accused of fraud",
        "Workflow Insight": "Insurers should distinguish between an inconsistency in a claimant's account and actual fraud — where the inconsistency is explicable (shock, initial impression later revised) and is not supported by other evidence, it is insufficient to sustain a fraud rejection; unsubstantiated witness comments (anonymous, no statement prepared) must not be relied upon as the basis for a fraud allegation; insurer's own investigation findings that support the claimant's general credibility must be weighed in the claimant's favour rather than ignored",
        "AI Rule Candidate": "IF account_changed_from_initial_police_report AND change_attributable_to_shock_or_reflection THEN inconsistency_alone_insufficient_for_fraud_finding; IF witness_refuses_to_be_named_or_provide_statement THEN witness_comments_cannot_form_basis_of_fraud_rejection; IF insurer_own_investigation_supports_claimant_credibility THEN findings_must_weigh_in_claimants_favour; fraud_allegation_requires_sufficient_corroborating_evidence_beyond_mere_inconsistency",
        "Source PDF": "DRN2794102.pdf",
    },
]


# ---------------------------------------------------------------------------
# Style helpers
# ---------------------------------------------------------------------------
def _row_fill(even: bool) -> PatternFill:
    color = "F5E6E6" if even else "FFFFFF"
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
        repo_root, "knowledge", "case-databases", "Theft_Case_Database.xlsx"
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
