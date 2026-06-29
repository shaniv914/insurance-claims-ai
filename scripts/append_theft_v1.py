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
# NEW CASES — Batch 2: THEFT-011 to THEFT-020
# Note: 9 of 10 cases are non-home-insurance (motor / travel / gadget /
# commercial retail) — flagged Is Core Case = "No — Commercial" or
# "No — Handling Dispute" as appropriate.
# Only THEFT-018 (LV household contents burglary; jewellery limits) is
# Is Core Case = Yes.
# ---------------------------------------------------------------------------
NEW_CASES: list[dict] = [
    {
        "Case ID": "THEFT-011",
        "FOS Decision ID": "DRN2801347",
        "Insurer Name": "Aviva Insurance Limited",
        "FOS Decision Date": "Not stated in document",
        "Claim Type": "Motor (car) insurance — car stolen; Aviva cancelled policy for non-disclosure of prior claims and convictions (would have offered cover on different terms so policy remained in force); spare key left in glove box by named driver Mr Y; Aviva declined theft claim under 'keys in car' policy exclusion; dispute about whether the exclusion was sufficiently highlighted in the policy summary",
        "Entry / Theft Method": "Motor vehicle theft using key left in vehicle — spare key left in glove box by named driver Mr Y; vehicle stolen using that key; specific method of gain beyond key availability not described",
        "Property Type": "Motor vehicle (personal car insurance)",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Spare key left in glove box of car by named driver Mr Y triggered the 'keys in car' policy exclusion; Aviva also cancelled (not voided) policy for non-disclosure — as it would have offered cover on different terms, cover was in force at time of theft but claim declined solely on the keys-in-car exclusion",
        "Evidence Dispute": "Mr T and Mr Y: 'keys in car' exclusion not sufficiently highlighted — adjudicator agreed it was 'buried' in a block of text in the policy summary. Aviva: exclusion contained in policy summary under the heading 'significant or unusual exclusions' as per FOS guidance; provided evidence it had informed Mr T the claim was declined on this exclusion. FOS (ombudsman): adjudicator found exclusion buried — ombudsman disagreed; policy summary longer than usual but first page was a cover sheet, second page very clear and easy to read, significant and unusual exclusions on third page under a bold heading; reasonable to expect Mr T to read first two or three pages of a Key Facts summary and note a bold heading for significant or unusual exclusions; exclusion sufficiently highlighted; parties made no further comment after provisional decision.",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — Aviva acted fairly and reasonably in declining the theft claim; 'keys in car' exclusion was sufficiently highlighted in the policy summary",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "'Keys in car' exclusion is enforceable where placed under a bold 'significant or unusual exclusions' heading in a Key Facts policy summary — even on page 3 of a longer-than-usual summary; consumer is expected to read at least the first two or three pages of a Key Facts summary and note a distinct bold heading for significant or unusual exclusions; where an insurer cancels (rather than voids) for non-disclosure because it would have offered cover on different terms, the policy remains in force at the time of the claim and exclusions still apply; a spare key left in the vehicle by a named driver triggers the keys-in-car exclusion regardless of whether the policyholder was aware the key was there",
        "Missing Evidence": "Specific details of why Mr Y left the spare key in the glove box and whether Mr T was aware of its presence",
        "Ombudsman Reasoning": "Adjudicator concluded exclusion was 'buried' in a block of text on page 3 of the policy summary; ombudsman reviewed summary — first page is a cover sheet, second page very clear and easy to read, significant and unusual exclusions on page 3 under a bold heading; not unreasonable to expect policyholder to read first two or three pages of a Key Facts summary; bold heading makes exclusion locatable; exclusion was sufficiently highlighted; Aviva correctly communicated decline reason to Mr T; neither party made further comment after provisional decision; complaint not upheld",
        "Workflow Insight": "A 'keys in car' exclusion placed under a bold 'significant or unusual exclusions' heading in the Key Facts policy summary is enforceable — even if located on page 3 of a longer-than-usual summary; when cancelling (not voiding) for non-disclosure, the policy remains in force and all exclusions continue to apply to any claim arising during the policy period; a named driver who leaves a spare key in the vehicle triggers the keys-in-car exclusion even if the main policyholder was unaware the key was there",
        "AI Rule Candidate": "IF spare_key_left_in_vehicle AND policy_contains_keys_in_car_exclusion THEN exclusion_applies_regardless_of_who_left_key; IF exclusion_listed_under_bold_significant_or_unusual_exclusions_heading_in_key_facts_summary THEN exclusion_is_sufficiently_highlighted; IF non_disclosure_leads_to_cancellation_not_voidance THEN policy_remains_in_force_and_all_exclusions_continue_to_apply",
        "Source PDF": "DRN2801347.pdf",
    },
    {
        "Case ID": "THEFT-012",
        "FOS Decision ID": "DRN2886197",
        "Insurer Name": "UK Insurance Limited",
        "FOS Decision Date": "Not stated in document",
        "Claim Type": "Travel insurance — bag stolen from car footwell at service station during a toilet break; Mr and Mrs T took turns so one was always within 8-10 feet of the car; UKI declined as valuables and personal belongings were left 'unattended' in a motor vehicle; complainants dispute the car was ever unattended",
        "Entry / Theft Method": "Theft from vehicle — bag removed from car footwell at service station while neither occupant was watching or able to notice the theft; theft entirely undetected until occupants returned to drive away; estate car meant bag could not be hidden in the boot",
        "Property Type": "Personal travel insurance",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Policy excluded cover for theft of personal possessions and valuables left in an unattended motor vehicle; UKI concluded the car was 'unattended' as neither Mr nor Mrs T was in a position to prevent the theft or even notice when someone approached",
        "Evidence Dispute": "Mr and Mrs T: car not left unattended — took turns for toilet break; one always within 8-10 feet; travelling with a dog; estate car meant bag could not be locked in boot as items visible through rear and side windows; close proximity made taking the bag unnecessary. UKI: belongings left unattended in motor vehicle — exclusion applies. FOS: accepted they took turns and one always in vicinity; but neither had any idea someone had opened the car door until about to drive away — they were not in a position to prevent or even notice the theft; car was therefore 'unattended'; distinguished from case where consumer sees theft but cannot prevent it (e.g. thief runs quickly) — those should be covered; here there was no awareness at all.",
        "Outcome Category": "Not Upheld",
        "Outcome": "UK Insurance Limited entitled to reject Mr and Mrs T's claim; no award made against UK Insurance Limited",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "Where a travel policy excludes theft from 'unattended' motor vehicles and provides no definition, the word is given its ordinary common sense meaning — a vehicle is unattended when the owner is not in a position to prevent a theft or even to notice when someone approaches; physical proximity alone is insufficient if the owner is entirely unaware the theft is occurring; key distinction: an insurer should pay where the consumer witnesses the theft but cannot prevent it (e.g. thief runs quickly) — that is different from being wholly unaware the theft occurred",
        "Missing Evidence": "None material — the factual circumstances were accepted by both parties; the dispute turned entirely on the legal interpretation of 'unattended'",
        "Ombudsman Reasoning": "Policy excludes theft from 'unattended' motor vehicle; no policy definition so ordinary common sense meaning applies; accepted Mr and Mrs T took turns and one always in vicinity with dog; but both letters confirm neither had any inkling someone had opened the car door until about to drive away — they were not 'attending' to the car or possessions inside it; key distinction: if consumer witnesses theft and fails to prevent it (e.g. quick thief) insurer should meet claim; here neither was aware theft was occurring — car was unattended; no reasonable basis to require UKI to make any payment",
        "Workflow Insight": "Travel insurance 'unattended vehicle' exclusions apply where the owner was entirely unaware the theft was occurring — physical proximity is insufficient if the owner could not notice or prevent the theft; the determinative question is whether the owner was in a position to notice and prevent the theft, not how far away they were; a consumer unable to secure valuables in a boot due to vehicle type should take the valuables with them or remain in a position to observe the vehicle interior",
        "AI Rule Candidate": "IF consumer_unaware_theft_occurred_until_after_the_fact THEN vehicle_was_unattended_for_exclusion_purposes; IF consumer_witnesses_theft_but_cannot_prevent_it THEN unattended_exclusion_does_not_apply_and_insurer_should_meet_claim; proximity_alone_insufficient_to_satisfy_attended_requirement_if_consumer_could_not_notice_or_prevent_theft",
        "Source PDF": "DRN2886197.pdf",
    },
    {
        "Case ID": "THEFT-013",
        "FOS Decision ID": "DRN-3028208",
        "Insurer Name": "Covea Insurance Plc",
        "FOS Decision Date": "11 Nov 2021",
        "Claim Type": "Motor (car) insurance — car stolen; Covea avoided (voided) policy from inception under CIDRA for non-disclosure of two 2019 speeding convictions at renewal; insured had declared one 2016 conviction; secondary complaint about £4,000 damage to recovered car alleged to have been caused by Covea's collection and storage",
        "Entry / Theft Method": "Motor vehicle theft — car stolen; specific method of entry not described; key issue is non-disclosure of motoring convictions at renewal leading to policy avoidance from inception",
        "Property Type": "Motor vehicle (personal car insurance)",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Mr O failed to disclose two 2019 speeding convictions at renewal — on the renewal call he said 'nothing's changed' and did not query the statement of fact which listed only the 2016 conviction; Covea's underwriting criteria (reviewed by FOS in confidence) showed it would not have offered any insurance if the two 2019 convictions had been declared; CIDRA careless qualifying misrepresentation — policy avoided from inception; premiums refunded; post-recovery damage assessed by senior engineer as consistent with theft and described as 'no or minimal' in a photo-based inspection",
        "Evidence Dispute": "Mr O: told broker about convictions at outset. Covea: renewal call recording shows Mr O said 'nothing's changed' when asked if anything had changed; statement of fact sent for checking listed only the 2016 conviction but Mr O did not query the omission; underwriting criteria (shown to FOS — confidential) demonstrates no cover would have been offered with those convictions; senior engineer confirmed reported damage is consistent with theft not with Covea's collection or storage. FOS: listened to renewal call — Mr O said 'nothing's changed' despite two new speeding convictions; had previously declared 2016 conviction so knew disclosure was required; should have noticed omission on statement of fact sent for checking; CIDRA careless qualifying misrepresentation established; avoidance + premium refund is correct remedy; post-recovery damage explanation accepted.",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — Covea acted fairly in avoiding motor policy from inception for careless misrepresentation under CIDRA; Covea is not responsible for post-recovery damage to Mr O's car",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "Under CIDRA, answering 'nothing's changed' at renewal when two new motoring convictions exist is a careless misrepresentation — prior disclosure of one earlier conviction demonstrates the consumer knew declarations were required; the statement of fact sent for checking creates an independent obligation to verify accuracy — failure to notice the omission of material information removes any defence; insurer must evidence underwriting criteria to justify avoidance — confidential underwriting guide shown to FOS satisfies this; CIDRA remedy for careless qualifying misrepresentation where insurer would not have offered cover at all: avoid from inception and refund premiums; a photo-based post-theft vehicle inspection calibrated for total-loss assessment does not catalogue all damage — 'no or minimal damage' in that context is not equivalent to 'no damage'",
        "Missing Evidence": "What Mr O said to broker at original policy inception (not available); whether the broker failed to pass on any disclosure would have changed the analysis; full details of the 2019 speeding conviction dates and types (established by Covea's driving licence check during the claim call)",
        "Ombudsman Reasoning": "Renewal call listened to — Mr O said 'nothing's changed' despite two 2019 speeding convictions; had previously declared 2016 conviction so was aware of the disclosure obligation; statement of fact sent for checking listed only the 2016 conviction — Mr O did not query the omission; Covea underwriting criteria reviewed in confidence — no cover would have been offered with those convictions; CIDRA careless qualifying misrepresentation established; avoidance + premium refund is correct CIDRA remedy; post-recovery damage: report said 'no or minimal' (not 'no damage') and senior engineer explained £4,000 is minimal vs a £27,000 car in theft context; most damage interior and not visible in photos; report was photo-based and produced for total-loss assessment; Covea not responsible for post-collection damage",
        "Workflow Insight": "Renewal calls with a generic 'has anything changed?' question require policyholders to proactively volunteer new convictions — answering 'no' without considering specific material changes is careless misrepresentation; the statement of fact sent for checking creates an independent verification duty — policyholders who fail to check it lose the ability to argue information was correctly provided; insurers investigating post-theft vehicle damage should clearly contextualise photo-based inspection reports (calibrated for total-loss assessment) to avoid disputes about whether 'no or minimal damage' means the vehicle sustained no damage at all",
        "AI Rule Candidate": "IF consumer_answers_nothing_changed_at_renewal AND new_convictions_exist THEN careless_misrepresentation_established; IF statement_of_fact_omits_material_information AND consumer_fails_to_notice_and_query THEN consumer_has_not_taken_reasonable_care; IF CIDRA_careless_misrepresentation AND insurer_would_not_have_offered_cover_at_all THEN avoidance_from_inception_and_premium_refund_is_correct_remedy; IF senior_engineer_confirms_damage_consistent_with_theft AND repair_cost_small_relative_to_vehicle_value THEN insurer_not_responsible_for_post_collection_damage",
        "Source PDF": "DRN-3028208.pdf",
    },
    {
        "Case ID": "THEFT-014",
        "FOS Decision ID": "DRN3072290",
        "Insurer Name": "AXA Insurance UK Plc",
        "FOS Decision Date": "20 Feb 2015",
        "Claim Type": "Motor (car) insurance — AXA applied additional premium after discovering via the CUE database an undisclosed theft claim Mr V had made on a prior policy for a different vehicle; Mr V disputes that a claim on a separate vehicle or policy need be disclosed and argues his No Claims Discount on the current vehicle is unaffected",
        "Entry / Theft Method": "Prior motor vehicle theft claim — circumstances of the prior theft on a different vehicle are not described in this decision; no current theft claim; dispute concerns whether the prior claim must be disclosed and whether AXA can apply additional premium as a result",
        "Property Type": "Motor vehicle (personal car insurance)",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Accepted",
        "Rejection Reason": "Not applicable — current policy accepted and in force; dispute concerns AXA's entitlement to apply additional premium after discovering via CUE that Mr V had not disclosed a prior theft claim on a separate vehicle policy",
        "Evidence Dispute": "Mr V: prior claim was on a separate policy for a different vehicle; NCD earned on current vehicle and is unaffected; he has already been penalised by the previous insurer; prior claim is irrelevant to current policy. AXA: claims history transcends individual vehicles and policies; the application question asked about any motor claims in the past five years; NCD and claims history are separate considerations; additional premium was automatically calculated and is correct; AXA entitled to exercise commercial discretion in its underwriting. FOS: AXA entitled to set its own questions and premium calculation methodology; question was clear — any motor claims in past five years; prior theft claim on a separate vehicle falls within that question; NCD (vehicle-specific) and overall claims history (person-specific) are distinct; AXA legitimately exercised its commercial judgement.",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — AXA legitimately applied additional premium reflecting correct claims history; AXA was entitled to exercise its commercial judgement in its underwriting",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "A motor insurer is entitled to ask about all motor claims made in the past five years across any vehicle or policy — claims history is not vehicle-specific; No Claims Discount earned on a specific vehicle and overall claims history are separate underwriting factors that may be assessed independently; where a question about claims history is clear and comprehensive, the insurer is entitled to apply its underwriting criteria (including premium adjustments) when it discovers an undisclosed prior claim; the insurer has commercial discretion to decide which factors it uses in risk assessment and premium calculation; prior penalisation by another insurer for the same claim does not prevent a subsequent insurer from also taking that claim into account",
        "Missing Evidence": "Not applicable — factual circumstances not in dispute; the entire dispute concerned the insurer's legal entitlement to use the prior claims history in its underwriting of the current policy",
        "Ombudsman Reasoning": "AXA entitled to decide what questions to ask and how to calculate premiums; question was clear — any motor claims in past five years; prior theft claim on a different vehicle under a different policy falls within that question; once aware of the undisclosed claim, AXA was entitled to apply its underwriting criteria; NCD (vehicle-specific) is distinct from overall claims history (person-specific); additional premium was legitimately calculated; Mr V free to seek lower premiums from other providers if he considered AXA's premium too high",
        "Workflow Insight": "Policyholders must understand that prior motor claims history for insurance purposes covers all vehicles and policies — a prior theft claim on a different vehicle must be disclosed when asked about any motor claims in the past five years; NCD earned on a specific vehicle does not neutralise the insurer's right to underwrite against the policyholder's overall claims history; being previously penalised by one insurer for a claim does not prevent a subsequent insurer from independently taking that same claim into account in its risk assessment",
        "AI Rule Candidate": "IF insurer_asks_about_any_motor_claims_in_past_five_years THEN prior_theft_claim_on_any_vehicle_any_policy_must_be_disclosed; insurer_may_treat_NCD_and_overall_claims_history_as_separate_underwriting_factors; IF insurer_discovers_undisclosed_prior_claim_via_CUE THEN entitled_to_apply_underwriting_criteria_including_premium_adjustment; prior_penalisation_by_previous_insurer_does_not_preclude_current_insurer_from_underwriting_same_claim",
        "Source PDF": "DRN3072290.pdf",
    },
    {
        "Case ID": "THEFT-015",
        "FOS Decision ID": "DRN3099586",
        "Insurer Name": "AXA Insurance UK Plc",
        "FOS Decision Date": "22 Apr 2016",
        "Claim Type": "Commercial shop and restaurant insurance — prior FOS decision required AXA to reinstate policy and settle theft and arson claims; when AXA settled the claims it paid no interest; V complains AXA should pay interest on the delayed payment and compensation for the delay",
        "Entry / Theft Method": "Commercial premises theft and arson (shop/restaurant) — specific method of entry not described; this decision concerns interest on delayed claim settlement only, not the theft or arson circumstances",
        "Property Type": "Commercial (shop and restaurant insurance policy — not home/contents)",
        "Dispute Type": "Handling / Reinstatement Dispute",
        "Coverage Decision": "Accepted",
        "Rejection Reason": "Not applicable — claims accepted and settled pursuant to prior FOS decision; dispute concerns whether interest must be paid on the time-lag between the date of loss and settlement",
        "Evidence Dispute": "V: AXA should pay interest on delayed settlement and compensation for the delay. AXA: not responsible for any avoidable delay so no interest or compensation owed. Adjudicator: interest should be paid but no compensation. AXA disagreed and requested ombudsman review. FOS (ombudsman): AXA misunderstands the purpose of interest — it is a component of indemnity to put the policyholder back in the position they would have been in if the insured event had never happened; time-lag between loss and payment requires interest regardless of whether the delay was avoidable or the insurer's fault; compensation for distress/inconvenience is distinct and only payable where avoidable delay or insurer mistakes caused additional harm — no avoidable delay here; 8% simple per annum is the appropriate interest rate.",
        "Outcome Category": "Upheld in Part",
        "Outcome": "AXA Insurance UK Plc to pay interest (less tax if properly deductible) at 8% simple per annum on the theft claim and material damage elements of the arson claim from the respective dates of loss to the date of payment; no compensation for delay awarded",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "The principle of indemnity requires an insurer to put the policyholder back in the same position as if the insured event had not occurred — where there is a time-lag between the insured event and payment, interest must be paid to achieve true indemnity regardless of whether the delay was avoidable or the insurer's fault; interest and compensation for distress/inconvenience are distinct obligations — interest is a component of indemnity; compensation for distress is only payable where avoidable delay or insurer mistakes caused additional harm to the policyholder; FOS applies 8% simple per annum as standard interest rate on insurance settlement payments",
        "Missing Evidence": "Precise dates of each loss event and corresponding settlement payment dates (needed to calculate interest quantum; these were not disputed in this decision)",
        "Ombudsman Reasoning": "Prior FOS decision required AXA to reinstate policy and settle claims; AXA settled but paid no interest; AXA argued no interest was owed because the delay was not its fault — ombudsman explains this misunderstands the purpose of interest; the indemnity principle requires interest for the time-lag between loss and payment regardless of fault; compensation for distress is separate and only arises from avoidable insurer delay or mistakes — this case required detailed investigation and AXA was not responsible for avoidable delay; no compensation; 8% simple per annum is appropriate interest rate",
        "Workflow Insight": "Insurers settling claims (including pursuant to FOS decisions) must include interest from the date of loss to the date of payment to achieve full indemnity — the absence of fault for the delay does not remove the interest obligation; interest should be offered proactively as part of any settlement rather than waiting to be directed; compensation for distress is a separate, additional obligation that only arises when the insurer's own avoidable fault or mistakes caused identifiable additional harm",
        "AI Rule Candidate": "IF time_lag_exists_between_date_of_loss_and_settlement_payment THEN interest_at_8_percent_simple_per_annum_is_payable_regardless_of_fault; interest_and_distress_compensation_are_distinct_obligations_and_must_not_be_conflated; IF no_avoidable_delay_or_insurer_mistakes THEN no_distress_compensation_owed_even_where_interest_is_payable; interest_must_be_included_in_any_insurance_settlement_to_achieve_the_indemnity_principle",
        "Source PDF": "DRN3099586.pdf",
    },
    {
        "Case ID": "THEFT-016",
        "FOS Decision ID": "DRN3125088",
        "Insurer Name": "Tesco Underwriting Limited",
        "FOS Decision Date": "6 Dec 2019",
        "Claim Type": "Motor (car) insurance — house burgled in December 2018; both house and car keys stolen; partner's car stolen immediately but Miss W's car was not; Tesco arranged locksmith and advised Miss W to safeguard her car; three days later Miss W's car stolen from outside home by person who had possession of the stolen key; Tesco rejected theft claim for failure to take reasonable protective steps",
        "Entry / Theft Method": "Motor vehicle theft using stolen keys — house burgled; car and house keys stolen; partner's car stolen at the time of the burglary; Miss W's car not immediately taken; three days later, person holding the stolen key drove away Miss W's car from outside home; no forced entry to vehicle required as thief possessed the key",
        "Property Type": "Motor vehicle (personal car insurance)",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Tesco rejected the theft claim on the grounds that Miss W failed to take reasonable steps to protect her car after being advised to do so; thief had Miss W's car key and knew where the car was kept; Miss W had access to the spare car key (she was using it) and could have moved the car to a safe location or added additional physical locks; placing bins in front of a car behind a low wall is insufficient when the thief possesses the key and knows the car's location",
        "Evidence Dispute": "Miss W and Mr J: car kept behind a wall with bins in front — reasonable protective measures; Tesco should have provided specific advice about protective measures before rejecting rather than giving examples afterwards. Tesco: Miss W was advised to safeguard her car when keys were reported stolen; spare key available — she could have moved car or added extra locks; bins behind a low wall are inadequate; apologised for failing to call back as promised and gave agent feedback. FOS: reasonable to expect Miss W to take protective steps knowing the thief had her key and knew her car's location; spare key available — she was using it; could have moved car or added locks; wall is low so car clearly visible; bins in front inadequate against a thief with a key; Tesco's failure to call back was poor service but apology + agent feedback was sufficient remedy; the short callback delay did not affect the claim outcome.",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — Tesco's rejection of the theft claim was fair and reasonable; apology and agent feedback was sufficient remedy for the service failure in not calling back as promised",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "Where a policyholder is aware that a thief has possession of the car key and knows the vehicle's location, a reasonable care obligation arises to take proactive protective steps before a locksmith can attend — this may include relocating the vehicle or adding physical locks; placing objects in front of a vehicle behind a low wall is not a reasonable protective measure when the thief holds the key; access to a spare key means the policyholder is in a position to move the vehicle and is expected to do so; a short unexplained delay by an insurer in calling back does not affect the claim outcome where the claim was promptly decided within three working days",
        "Missing Evidence": "Specific content of the advice Tesco gave Miss W when it told her to 'safeguard' her car at the time keys were reported stolen (ombudsman found general advice to safeguard was sufficient — Tesco was not required to itemise specific measures in advance)",
        "Ombudsman Reasoning": "Miss W and Mr J knew thief had key and knew car location — obligation to take reasonable protective steps arose at that point; spare key available — Miss W was using it; could have moved car to safe location or added extra locks; photo shows car visible over low wall; bins in front of car inadequate against a thief who holds the key; Tesco's failure to call back acknowledged — apology and agent feedback is sufficient remedy; three-working-day delay in calling back did not affect outcome as claim was then promptly decided; complaint not upheld",
        "Workflow Insight": "A theft claim may be declined where the policyholder was forewarned that the thief had the car key and knew the vehicle location, but failed to take available protective steps (such as relocating the vehicle) before a locksmith attended; the availability of a spare key is a critical factor — if the policyholder can access the vehicle they are expected to take protective action; an insurer's service failure in not making a promised callback does not affect the claim outcome if the claim was handled and decided promptly overall",
        "AI Rule Candidate": "IF policyholder_aware_thief_holds_car_key AND knows_vehicle_location AND spare_key_available THEN policyholder_must_take_reasonable_protective_steps_including_relocating_vehicle; IF protective_measures_inadequate_against_keyed_access THEN reasonable_care_obligation_not_met_and_claim_may_be_declined; IF insurer_service_failure_did_not_affect_claim_outcome THEN apology_and_internal_feedback_is_proportionate_remedy",
        "Source PDF": "DRN3125088.pdf",
    },
    {
        "Case ID": "THEFT-017",
        "FOS Decision ID": "DRN-3197736",
        "Insurer Name": "AmTrust Europe Limited",
        "FOS Decision Date": "19 Jan 2022",
        "Claim Type": "Gadget insurance — laptop noticed missing at airport during check-in for return flight from overseas trip; AmTrust declined for two reasons: (1) laptop stored in cargo hold — later conceded this was a misunderstanding of Mr A's account; (2) travelling against FCO advice — policy general exclusion 33 requires the claim to be 'a direct result' of such travel; FOS found that health-related FCO advice is not causally linked to a theft and the 'direct result' test was not met",
        "Entry / Theft Method": "Theft at airport — laptop noticed missing at airport check-in for return flight; Mr A reported to airport security but left family to follow up as he needed to catch his flight; exact method of theft not established; Mr A believed it was stolen not lost",
        "Property Type": "Personal gadget insurance (laptop — standalone gadget policy, not home/contents)",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "(1) Laptop stored in cargo hold — exclusion for hold storage applied; AmTrust later conceded this was a misunderstanding of Mr A's account. (2) Mr A had travelled against FCO advice — policy general exclusion 33 excludes claims that occur 'as a direct result' of travelling to a country where the FCO advise against all (but essential) travel",
        "Evidence Dispute": "Mr A: laptop was not in cargo hold; had valid personal reasons for travel; FCO advice was health-related and has no causal link to a theft. AmTrust: cargo hold exclusion applied (later conceded incorrect); FCO exclusion 33 still applies as Mr A travelled against FCO advice. FOS: cargo hold misunderstanding resolved during call and confirmed in AmTrust's FOS response; FCO advice was health-related — a health advisory does not carry a close causal relationship to the theft of a laptop; the word 'direct' in 'direct result' requires a close causal connection between the nature of the FCO advice and the specific loss; the risk of laptop theft was not heightened by a health advisory; exclusion 33 cannot fairly be applied.",
        "Outcome Category": "Upheld",
        "Outcome": "AmTrust Europe Limited must reconsider Mr A's theft claim under his policy in line with the remaining terms and conditions, without relying on general exclusion 33 (travelling against FCO advice); FOS is not directing AmTrust to pay the claim — it may still assess the claim on remaining policy terms",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "An FCO travel exclusion requiring a claim to be 'a direct result' of travelling against advice demands a close causal connection between the nature of the FCO advice and the specific type of loss claimed — 'direct' means more than temporal or geographic coincidence; where FCO advice is health-related, it does not carry a direct causal relationship to a theft loss — health advice does not increase the risk of theft; an insurer cannot rely on an FCO exclusion unless the particular advisory directly caused or heightened the risk of the specific type of loss sustained; where an insurer declines on an incorrect factual premise (cargo hold) that is later conceded, remaining grounds are examined carefully",
        "Missing Evidence": "Valid police crime reference report from local authorities (required by policy definition of theft — Mr A reported to airport security and left family to follow up; whether a valid crime reference was ever obtained is unclear and remains for AmTrust to assess on reconsideration)",
        "Ombudsman Reasoning": "Cargo hold misunderstanding resolved — AmTrust confirmed in the call and in its FOS response; FCO exclusion 33 requires claim to be a 'direct result' of FCO travel advice; 'direct' denotes a close causal connection — not merely 'as a result'; FCO advice was health-related; a health advisory has no close causal link to the theft of a laptop; the risk of theft was not heightened by the health-related FCO advice; exclusion 33 cannot fairly be applied; AmTrust must reconsider without relying on that exclusion; FOS is not directing payment — AmTrust may assess the claim on remaining policy terms",
        "Workflow Insight": "When relying on an FCO exclusion with a 'direct result' causation test, the insurer must assess whether the nature of the FCO advice (health, security, political unrest) is causally linked to the specific type of loss — health advisories do not increase theft risk and cannot be treated as directly causing a theft; insurers should not decline on a factual misunderstanding and then rely on an exclusion whose own causation test is not satisfied; 'direct result' exclusions require analysis of the causal chain, not just the presence of FCO advice",
        "AI Rule Candidate": "IF FCO_exclusion_requires_direct_result AND FCO_advice_is_health_related AND claim_is_for_theft THEN direct_causation_test_not_met_and_exclusion_cannot_apply; direct_result_means_close_causal_connection_between_nature_of_FCO_advice_and_specific_type_of_loss; IF insurer_declines_on_factual_misunderstanding_later_conceded AND remaining_exclusion_fails_causation_test THEN claim_must_be_reconsidered_on_remaining_policy_terms",
        "Source PDF": "DRN-3197736.pdf",
    },
    {
        "Case ID": "THEFT-018",
        "FOS Decision ID": "DRN3351038",
        "Insurer Name": "Liverpool Victoria Insurance Company Limited",
        "FOS Decision Date": "25 Apr 2016",
        "Claim Type": "Household contents insurance — 2015 home burglary; more than £20,000 of jewellery stolen; LV applied a valuables limit of £5,000 and paid a further £3,000 for unspecified personal possessions; Mr G claims LV told him at inception (2002) that these limits would not apply to theft claims",
        "Entry / Theft Method": "Burglary — home broken into in 2015; more than £20,000 of jewellery stolen; specific method of forced entry not described in the decision",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Accepted — Disputed Settlement",
        "Rejection Reason": "Not applicable — theft claim accepted but settled at the policy's valuables limit (£5,000) and unspecified personal possessions sub-limit (£3,000); Mr G disputes the application of these limits, claiming they do not apply to theft; LV maintains limits always applied and are clearly shown on all schedule documents since at least 2008",
        "Evidence Dispute": "Mr G: when he first took out the policy in 2002 LV told him the valuables and personal possessions limits did not apply to theft claims; LV never informed him this had changed; if he had known limits applied he would have arranged higher cover. LV: limits clearly shown on policy schedule documents; Mr G was told he could change them; £100 paid for complaint handling about how the claim was managed. FOS: no recording or documents available from 2002 inception; standard industry approach is for limits to apply to all claims — it is unlikely Mr G was told limits didn't apply to theft; earliest available documents from 2008 show £5,000 valuables limit; 2015 schedule identical; policy summary explains limits apply; full policy states contents covered up to limits shown on schedule; no exception for theft; Mr G admitted not reading renewal documents sent each year.",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — LV entitled to apply valuables and unspecified personal possessions limits to Mr G's theft claim; no additional payment required beyond what LV had already paid",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Policy limits for specified categories (valuables, unspecified personal possessions) apply to all claims including theft unless the policy explicitly states otherwise — there is no industry standard or practice of limits not applying to theft claims; an unsubstantiated allegation that verbal assurances at inception override clearly-stated schedule limits will not be upheld without supporting evidence, especially where the allegation contradicts standard industry practice; policy limits clearly shown on the schedule and renewal documents bind the policyholder whether or not those documents were read — the obligation to read and understand renewal documents rests with the policyholder; jewellery or valuables exceeding the schedule limit require separate specified item cover to obtain adequate protection",
        "Missing Evidence": "Recording or documentation from the original 2002 policy inception (not available — earliest LV documents from 2008; the absence of inception records is the central evidential gap; no contemporaneous evidence of what was said about limits at the point of sale)",
        "Ombudsman Reasoning": "No evidence available from 2002 inception — no telephone recording or documents; Mr G alleges limits were said not to apply to theft but standard industry practice is for limits to apply to all claims; unlikely he was told otherwise; earliest documents from 2008 show £5,000 valuables limit on the schedule; 2015 schedule identical; policy summary and full policy explain limits apply; no exception for theft stated in any document; Mr G admitted not reading renewal documents each year — would have seen and understood limits if he had done; limits fairly applied to his theft claim; LV not required to pay more; £100 voluntarily paid for complaint handling already noted",
        "Workflow Insight": "Where a policyholder alleges an oral representation at inception overrides clearly stated policy limits, FOS requires contemporaneous evidence — undocumented telephone conversations from many years before the claim are insufficient to override schedule limits; policy schedules and renewal documents sent each year confirm current limits — policyholders who do not read renewals cannot later claim ignorance of limits shown thereon; handlers should proactively advise policyholders with high-value jewellery to consider whether the valuables sub-limit is adequate and whether separate specified item cover should be arranged",
        "AI Rule Candidate": "IF consumer_alleges_verbal_promise_at_inception_that_limits_do_not_apply AND no_supporting_evidence AND standard_practice_contradicts_allegation THEN FOS_will_not_uphold_allegation; policy_schedule_limits_apply_to_all_claims_including_theft_unless_policy_explicitly_states_otherwise; IF consumer_admits_not_reading_renewal_documents THEN consumer_cannot_claim_ignorance_of_limits_clearly_shown_on_schedule; jewellery_exceeding_valuables_limit_must_be_separately_specified_to_obtain_adequate_cover",
        "Source PDF": "DRN3351038.pdf",
    },
    {
        "Case ID": "THEFT-019",
        "FOS Decision ID": "DRN-3711750",
        "Insurer Name": "Royal & Sun Alliance Insurance Limited",
        "FOS Decision Date": "3 Jan 2023",
        "Claim Type": "Home insurance — primary claim for escape of water (EOW); during RSA-appointed repairers' attendance, jewellery went missing from Miss B's home; Miss B believed RSA's contractors stole it; dispute covers (1) EOW claim handling failures (missed appointments, delays, lost earnings, poor quality work); (2) whether jewellery theft should be dealt with under the EOW claim or as a separate second claim",
        "Entry / Theft Method": "Theft by invited contractor — jewellery went missing from home while RSA-appointed painters were attending to carry out escape of water repair works; Miss B believed RSA's repairers were responsible; police investigation was ongoing at time of this FOS decision",
        "Property Type": "Residential home",
        "Dispute Type": "Handling / Reinstatement Dispute",
        "Coverage Decision": "Accepted",
        "Rejection Reason": "Not applicable — EOW claim accepted; jewellery theft not declined; RSA agreed to consider theft as a separate second claim with the £200 excess waived (or £200 paid as compensation if theft claim not accepted); dispute concerns whether the theft should be subsumed within the EOW claim and whether RSA's EOW handling compensation is adequate",
        "Evidence Dispute": "Miss B: RSA's repairers clearly stole jewellery on the balance of probabilities; should be dealt with under EOW claim not as a second claim; cash-in-lieu payment too low; having two claims on record is unfair and will increase premiums. RSA: theft is a separate incident from EOW; police investigation ongoing — cannot determine who stole jewellery; £1,740 (lost earnings) + £250 (D&I) is fair for EOW handling failures; cash in lieu must be at RSA repairer rates per policy terms; excess waived for theft claim (or £200 paid as compensation) + £125 for door lock. FOS: not appropriate to determine criminal responsibility — that is a matter for criminal courts; RSA cooperated with police; theft is a separate incident from EOW and two-claim treatment is correct; RSA's EOW compensation package is fair; cash in lieu at RSA repairer rates is policy-compliant; RSA's offer on the theft claim is fair and reasonable.",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — RSA's offer to waive the £200 excess for the theft claim (or pay it as compensation if the theft claim is not accepted) is fair and reasonable; RSA's approach to EOW claim handling compensation and cash-in-lieu settlement is upheld",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Handling Dispute",
        "Key Policy Clause": "FOS cannot determine whether a specific person committed a theft — that is a criminal matter for the courts; theft of property occurring during contractor attendance for a separate insured peril (escape of water) is a distinct insurable incident and should be treated as a second claim — the fact that the contractors were on site due to the first claim does not merge the two incidents; an insurer's cash-in-lieu settlement is limited to what the insurer's own approved repairers would have charged per standard policy terms; waiving the excess for a theft claim where the insurer's own contractors are the suspected thieves is a fair and reasonable outcome; if police later conclude the insurer's repairers committed the theft, the insurer should then consider whether it is fair to bring the theft claim under the original EOW claim",
        "Missing Evidence": "Police investigation outcome (ongoing at time of FOS decision — critical for determining whether RSA's contractors were responsible and whether the theft claim should subsequently be dealt with under the EOW claim); proof of ownership and value of the missing jewellery (required for any theft claim assessment)",
        "Ombudsman Reasoning": "EOW handling: RSA entitled to use own repairers per policy terms; cash in lieu limited to RSA repairer rates; £1,740 lost earnings + £250 D&I is fair for the significant handling failures (missed appointments, delays, contractor damage); further payment not required. Theft of jewellery: FOS cannot find RSA's repairers stole the jewellery — theft is a criminal matter for courts; theft reported to police; RSA cooperated with police investigation; theft is a separate incident from EOW and two-claim treatment is correct; excess waiver for theft claim (or £200 as compensation if claim not accepted) + £125 for door lock is fair; Miss B cannot be compensated for increased premiums resulting from two-claim recording where the two-claim approach is correct; if police conclude RSA's repairers committed the theft, RSA should reconsider whether to deal with it under the EOW claim at that point",
        "Workflow Insight": "Where property goes missing during contractor visits arranged by the insurer for a separate claim, it should be treated as a distinct second claim with its own excess and process — not subsumed into the original claim simply because the same contractors are involved; FOS will not determine criminal responsibility for suspected contractor theft; the insurer should cooperate with the police investigation and consider the theft claim on its own merits; cash-in-lieu settlements should be based on the insurer's own approved repairer rates per policy terms, not the consumer's preferred repairer's quote",
        "AI Rule Candidate": "IF theft_occurs_during_insurer_appointed_contractor_visit THEN treat_as_separate_claim_not_subsumed_within_original_claim; FOS_cannot_determine_criminal_responsibility_for_alleged_contractor_theft; cash_in_lieu_settlement_limited_to_insurer_own_repairer_rates_per_policy_terms; IF insurer_cooperates_with_police_on_contractor_theft_investigation THEN insurer_has_not_acted_unreasonably_in_its_handling",
        "Source PDF": "DRN-3711750.pdf",
    },
    {
        "Case ID": "THEFT-020",
        "FOS Decision ID": "DRN-3764842",
        "Insurer Name": "I.M.S. (London) Limited",
        "FOS Decision Date": "21 Jul 2023",
        "Claim Type": "Commercial retail insurance (sole trader) — theft of a Vitesse statue at a classic car and autojumble show while trader was packing items into the boot of his car; IMS declined under a policy endorsement (ENDP4639900) requiring either forcible/violent entry or permanently sited security guards for theft cover at exhibition premises; no forced entry occurred and club volunteers do not qualify as permanently sited security guards",
        "Entry / Theft Method": "Opportunistic theft at exhibition — Vitesse statue stolen while the sole trader was packing items into the boot of his car at a classic car and autojumble show; no forcible or violent entry; theft occurred during loading in an open outdoor setting",
        "Property Type": "Commercial (sole trader retail insurance — stock and contents at exhibitions)",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Policy endorsement ENDP4639900 requires either (i) forcible or violent entry or (ii) exhibition premises protected by permanently sited security guards for theft cover at exhibitions; no forced entry occurred; club volunteers managing car parks, entrances and issuing wristbands are not 'permanently sited security guards'; event organiser confirmed to the loss adjuster that no paid security guards were used at the event; endorsement conditions not met",
        "Evidence Dispute": "C (Mr W): club volunteers who managed car parks, entrances and issued wristbands constitute security measures and satisfy the endorsement requirement for permanently sited security guards. IMS: event organiser confirmed no security guards were employed — only club volunteers; volunteers do not meet 'permanently sited security guards' standard; endorsement was added when cover limits were increased — signals requirement for professional security. FOS: accepted no forced entry occurred; considered whether volunteers satisfy 'permanently sited security guards'; volunteers had various responsibilities (car parks, entrances, wristbands) and are not permanently sited professional security staff; exhibitions are high-risk events; endorsement added as condition of increased cover limits — suggests requirement for permanent professional security guards; IMS acted fairly in declining.",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — IMS acted fairly and reasonably in declining the theft claim by relying on the policy endorsement; IMS not required to do anything differently",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "A retail policy endorsement requiring 'permanently sited security guards' at exhibition premises for theft cover (where no forcible/violent entry occurred) means paid, professional security staff permanently stationed at the event — club volunteers with multiple administrative responsibilities (car parks, entrances, wristband issuance) do not meet this standard; exhibitions are recognised as high-risk events for stock theft — endorsements mandating professional security are a reasonable and enforceable insurer risk control; an endorsement added as a condition of increasing cover limits is a deliberate underwriting risk control that must be strictly complied with; the absence of both forcible entry and permanently sited security guards means the endorsement is breached and the claim is excluded",
        "Missing Evidence": "Full scope of volunteer duties and deployment at the event (assessed by FOS but found insufficient); copy of original endorsement negotiation documents showing the intent behind 'permanently sited security guards' (not needed — FOS satisfied the standard means paid professional security)",
        "Ombudsman Reasoning": "No forcible or violent entry — accepted by both parties; must therefore consider whether the exhibition premises were protected by permanently sited security guards; event organiser confirmed no paid security guards at event — only club volunteers; volunteers had various responsibilities including car parks, entrances, and wristband issuance and do not satisfy 'permanently sited security guards'; exhibitions are high-risk events so a requirement for professional security is reasonable; endorsement was added when cover limits were increased — signals a deliberate requirement for permanent professional security rather than ad hoc volunteer arrangements; C did not comply with the endorsement; IMS acted fairly and reasonably",
        "Workflow Insight": "Commercial traders attending exhibitions must verify their policy endorsements and confirm professional security guard arrangements before each event — a 'permanently sited security guards' requirement is not satisfied by event volunteers with general administrative roles; sole traders should seek written confirmation from event organisers about security arrangements to verify endorsement compliance before attending; endorsements added as conditions of increased cover limits are deliberate high-risk controls that must be strictly met",
        "AI Rule Candidate": "IF exhibition_theft_endorsement_requires_permanently_sited_security_guards AND only_volunteer_marshals_present THEN endorsement_not_met_and_theft_claim_excluded; permanently_sited_security_guards_means_paid_professional_security_staff_not_volunteers_with_multiple_responsibilities; IF no_forcible_or_violent_entry AND no_permanently_sited_security_guards THEN exhibition_theft_endorsement_is_breached; IF endorsement_added_as_condition_of_increased_cover_limits THEN endorsement_requirements_are_deliberate_risk_controls_to_be_strictly_applied",
        "Source PDF": "DRN-3764842.pdf",
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
