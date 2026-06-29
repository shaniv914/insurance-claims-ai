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
# NEW CASES — Batch 3: THEFT-021 to THEFT-030
# Note: All 10 cases are non-home-insurance (motor / motorcycle / motorhome /
# commercial motor trade / commercial property) — all flagged
# Is Core Case = "No — Commercial" or "No — Broker Dispute".
# No Is Core Case = Yes cases in this batch.
# ---------------------------------------------------------------------------
NEW_CASES: list[dict] = [
    {
        "Case ID": "THEFT-021",
        "FOS Decision ID": "DRN4331324",
        "Insurer Name": "Acromas Insurance Company Limited",
        "FOS Decision Date": "15 Mar 2017",
        "Claim Type": "Motor (car) insurance — car stolen after policyholder left keys in ignition with engine running while visiting a friend for 10–15 minutes; Acromas declined under 'keys in vehicle' exclusion in both Loss or Damage section and General Exceptions; Acromas initially gave misleading information that it would meet the claim, corrected next day; £100 compensation paid for the misleading information",
        "Entry / Theft Method": "Motor vehicle theft — keys left in ignition with engine running while policyholder went into friend's house for 10–15 minutes; car stolen during that absence; policyholder did not intend to leave it unattended but made a conscious decision to go inside",
        "Property Type": "Motor vehicle (personal car insurance)",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Policy 'keys in vehicle' exclusion applied — car left unattended with keys in ignition and engine running for 10–15 minutes; exclusion stated in both the Loss or Damage section and General Exceptions of the policy, and in the Certificate of Insurance; Acromas also stated in the policy booklet that Mr M must protect the car from loss; £100 compensation already paid for initial misleading information",
        "Evidence Dispute": "Mr M: didn't intend to leave car unattended; was invited in by friend briefly. Acromas: exclusion clearly applies — keys in ignition while unattended; exclusion sufficiently highlighted in multiple policy sections. FOS: Mr M made a conscious decision to go inside for 10–15 minutes; policy exclusion stated in multiple locations (Loss or Damage section, General Exceptions, Certificate of Insurance); Acromas corrected incorrect information next day; £100 compensation for brief misleading period is fair; complaint not upheld",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — Acromas' decision to decline the theft claim was reasonable; 'keys in vehicle' exclusion sufficiently highlighted; £100 compensation for misleading information was proportionate",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "A 'keys in vehicle' exclusion applies whenever a vehicle is left unattended with keys in the ignition regardless of the policyholder's intention or brevity of absence; the exclusion is sufficiently highlighted when stated in multiple policy sections (Loss or Damage section, General Exceptions, and Certificate of Insurance); where an insurer initially gives misleading information that it will meet a claim but corrects it the next day and pays £100 compensation, the remedy is proportionate; the policyholder's general duty in the policy booklet to protect the vehicle reinforces the exclusion",
        "Missing Evidence": "No material missing evidence — facts not in dispute; Mr M accepted he left the car with keys in the ignition for 10–15 minutes while visiting his friend",
        "Ombudsman Reasoning": "Mr M made a conscious decision to go into his friend's house for 10–15 minutes, leaving the car with keys in the ignition and engine running; policy exclusion applies in exactly these circumstances; exclusion stated in Loss or Damage section, General Exceptions and Certificate of Insurance; Acromas gave misleading information initially but corrected it next day and paid £100; complaint not upheld",
        "Workflow Insight": "A 'keys in vehicle' exclusion is enforceable where the policyholder consciously leaves the vehicle — even briefly — with keys in the ignition; stating the exclusion in multiple policy documents (Loss or Damage section, General Exceptions, Certificate of Insurance) creates a strong case for sufficient disclosure; when an insurer initially gives wrong information about a claim but corrects it within 24 hours, modest D&I compensation (£100) is the appropriate remedy",
        "AI Rule Candidate": "IF keys_left_in_ignition AND vehicle_left_unattended THEN keys_in_vehicle_exclusion_applies_regardless_of_duration_or_stated_intention; IF exclusion_stated_in_multiple_policy_sections_including_certificate THEN exclusion_is_sufficiently_highlighted; IF insurer_gives_misleading_claim_decision_corrects_within_24_hours AND pays_100_compensation THEN remedy_is_proportionate",
        "Source PDF": "DRN4331324.pdf",
    },
    {
        "Case ID": "THEFT-022",
        "FOS Decision ID": "DRN4423879",
        "Insurer Name": "Liverpool Victoria Insurance Company Limited",
        "FOS Decision Date": "19 Apr 2018",
        "Claim Type": "Motor (motorcycle) insurance purchased through broker — motorbike stolen; LV declined under garaging endorsement requiring the motorbike to be kept in a 'properly constructed and locked garage'; Mr K's storage area had a brick-frame entrance and metal shutter but rear opened to garden, sides were part timber/plastic or glass, and roof was made of plastic panels; LV could not clearly define 'properly constructed' but assessed physical construction from photographs",
        "Entry / Theft Method": "Motor vehicle theft (motorcycle) — motorbike stolen from storage area that appeared garage-like at the entrance but was structurally inadequate; rear opened to garden, sides non-solid (timber/plastic/glass), roof plastic panels; precise theft method not described",
        "Property Type": "Motor vehicle (personal motorcycle insurance)",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Policy garaging endorsement required motorbike to be kept in a 'properly constructed and locked garage'; the storage area had a brick-frame entrance and metal shutter but the rear was not secure (opened to garden), sides were part timber/plastic or glass, and the roof was plastic panels per online image provided by LV — not a properly constructed garage",
        "Evidence Dispute": "Mr K: the storage area looks like a garage; LV has not clearly defined 'properly constructed'; endorsement terms are ambiguous. LV and investigator: photos show the area is not secure — rear opens to garden, sides non-solid, roof plastic. FOS: reviewed photos and online image; entrance has appearance of a garage (brick frame, metal shutter) but rear not secure, opens to garden; sides part timber and part plastic/glass; roof plastic panels; not a properly constructed and locked garage; complaint not upheld",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — LV's decision to decline was reasonable; Mr K's motorbike was not kept in a properly constructed and locked garage as required by the garaging endorsement",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "A 'properly constructed' garage requires full structural integrity on all sides including the rear and roof — an enclosure with an impressive brick-frame entrance and metal shutter does not satisfy the requirement if the rear opens to a garden, sides are non-solid (timber/plastic), and the roof is plastic panels; where a policy endorsement requires a 'properly constructed and locked garage' without defining the term, the ordinary meaning applies: the entire structure must be fully enclosed and secure; the visual appearance of the entrance alone is insufficient if the remainder of the structure is inadequate",
        "Missing Evidence": "No material missing evidence — the physical construction of the storage area was assessed from photographs provided by Mr K and online images provided by LV",
        "Ombudsman Reasoning": "Policy schedule included garaging endorsement requiring properly constructed and locked garage; policy does not define 'properly constructed'; reviewed photographs — entrance has brick frame and metal shutter giving appearance of a garage; inside: rear not secure, opens to garden; sides part timber and part plastic/glass; LV also provided online image showing plastic panel roof; not a properly constructed and locked garage; endorsement conditions not met; LV's decision was fair and reasonable; complaint not upheld",
        "Workflow Insight": "When assessing a garaging endorsement, handlers must look beyond the entrance — the rear, side walls and roof must all provide genuine structural security; photographs of all sides (including rear and roof) should be obtained; the absence of a policy definition of 'properly constructed' does not assist the claimant — the ordinary reasonable meaning requires full structural security on all sides",
        "AI Rule Candidate": "IF garaging_endorsement_requires_properly_constructed_locked_garage AND storage_area_has_unsecured_rear_or_non_solid_sides_or_plastic_roof THEN endorsement_not_satisfied; properly_constructed_garage_requires_full_structural_integrity_on_all_sides_including_rear_and_roof; absence_of_policy_definition_of_properly_constructed_means_ordinary_reasonable_meaning_applies",
        "Source PDF": "DRN4423879.pdf",
    },
    {
        "Case ID": "THEFT-023",
        "FOS Decision ID": "DRN-4621255",
        "Insurer Name": "U K Insurance Limited",
        "FOS Decision Date": "20 Mar 2024",
        "Claim Type": "Motor (car) insurance — theft claim accepted and settled by UKI; Mr S raised post-settlement complaints: (1) UKI incorrectly told him he had 4 years' NCD when correct figure was 2 years after the theft claim reduced it; (2) UKI classified the theft claim as 'at fault'; (3) UKI failed to call him back; Mr S has medical conditions and disabilities that heighten his susceptibility to stress; incorrect NCD information caused him to underestimate premium increases with his new insurer",
        "Entry / Theft Method": "Motor vehicle theft — car stolen; circumstances of theft not disputed; all issues in this decision concern post-settlement NCD information error and claim classification rather than the theft circumstances",
        "Property Type": "Motor vehicle (personal car insurance)",
        "Dispute Type": "Claim Recording / Administrative Dispute",
        "Coverage Decision": "Accepted",
        "Rejection Reason": "Not applicable — theft claim accepted and settled by UKI; dispute concerns incorrect NCD information given post-settlement and the 'at fault' classification of the theft claim",
        "Evidence Dispute": "Mr S: UKI told him he had 4 years' NCD (wrong — 2 years correct after theft claim reduced it); theft should not be 'at fault' as it was outside his control; T&Cs unclear; medical conditions and disabilities heightened impact of the error. UKI: NCD error acknowledged; 'at fault' classification correct industry practice where no third party to recover from; policy terms clear on NCD reduction; £50 initial compensation offered. FOS: NCD error admitted; 'at fault' classification is correct and policy wording is clear; Mr S's documented medical conditions make the error's impact greater — total compensation of £300 is fair (deducting £50 already paid)",
        "Outcome Category": "Upheld in Part",
        "Outcome": "UKI must pay Mr S £300 total compensation (deducting sums already paid); 'at fault' classification of the theft claim upheld as correct industry practice; NCD error is upheld as UKI's mistake warranting increased compensation given Mr S's medical conditions",
        "Compensation Awarded (£)": 300,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "A theft claim where no third party can be pursued is correctly classified as a 'fault/bonus-disallowed' claim — the term 'at fault' describes the insurer's inability to recover costs, not the policyholder's culpability; policy terms stating 'if you claim on your policy, we may reduce the NCD' are clear and apply to theft claims; where an insurer misstates NCD information causing a policyholder to underestimate premium increases with a new insurer, compensation must reflect the actual impact — including any documented medical conditions and disabilities that heighten the effect of the error",
        "Missing Evidence": "None material on the NCD error — UKI admitted the mistake; the uplift in compensation reflects Mr S's documented medical conditions and their sensitivity to stress",
        "Ombudsman Reasoning": "'At fault' classification is correct industry practice — insurers record claims as 'at fault' (bonus-disallowed) when there is no third party to pursue; NCD policy terms are clear; NCD error admitted by UKI; Mr S had to notify new insurer and pay more than expected; Mr S has medical conditions seriously impacted by stressful situations — error's impact on him was greater than on most; £50 insufficient; £300 total (net of amount already paid) is fair and reasonable",
        "Workflow Insight": "Insurers must verify NCD information before communicating it post-claim, as errors directly affect premiums with the next insurer; a theft claim is correctly classified 'at fault' (bonus-disallowed) in industry databases where no third party can bear the insurer's costs; when assessing compensation for NCD misinformation, policyholders' disclosed medical conditions that heighten susceptibility to financial stress are relevant to the quantum of any award",
        "AI Rule Candidate": "IF theft_claim AND no_third_party_to_recover_from THEN claim_correctly_classified_as_fault_and_NCD_may_be_reduced; IF insurer_misstates_NCD_information_post_settlement THEN compensation_must_reflect_actual_financial_and_emotional_impact; IF policyholder_has_documented_medical_conditions_heightening_impact_of_insurer_error THEN compensation_uplift_is_appropriate",
        "Source PDF": "DRN-4621255.pdf",
    },
    {
        "Case ID": "THEFT-024",
        "FOS Decision ID": "DRN-4800291",
        "Insurer Name": "U K Insurance Limited trading as NIG",
        "FOS Decision Date": "30 Aug 2024",
        "Claim Type": "Commercial motor trade insurance — vehicle theft claim made October 2023 plus a simultaneous fire damage claim on the same policy; NIG experienced delays and communication issues; M (limited company) failed to provide vehicle keys to NIG for 3+ months; CCTV footage required specialist recovery and was recovered after NIG's final response; NIG offered £350 compensation for handling issues; M disputed NIG's right to require a formal interview and claimed no obligation to provide evidence after 30 days from policy expiry",
        "Entry / Theft Method": "Motor vehicle theft (commercial motor trade) — specific method of theft not described; claim under investigation at time of FOS decision; a second fire damage claim on the same policy was made shortly after the theft claim, adding complication",
        "Property Type": "Commercial (motor trade insurance — not home/contents)",
        "Dispute Type": "Handling / Reinstatement Dispute",
        "Coverage Decision": "Accepted",
        "Rejection Reason": "Not applicable — claim not declined; dispute concerns NIG's handling delays, its right to require a formal interview as part of the investigation, and the interpretation of the policy's 30-day evidence provision",
        "Evidence Dispute": "M: NIG has enough information to settle the claim; not obliged to provide evidence more than 30 days after policy expiry; formal interview not contractually required. NIG: M delayed providing keys for 3+ months; CCTV required specialist recovery; simultaneous fire and theft claims warranted detailed investigation; formal interview is reasonable; £350 offered for handling issues. FOS: M caused delays by not providing keys timely; simultaneous fire claim justifies more detailed investigation; one formal interview request is reasonable in complex circumstances; policy 30-day limit applies to when M must provide evidence, not when NIG can request it; £350 adequate for inconvenience; limited company cannot experience distress",
        "Outcome Category": "Not Upheld",
        "Outcome": "NIG's offer of £350 compensation for handling delays and communication issues is fair and reasonable; NIG is entitled to require a formal interview before settling the claim; M must engage with the formal interview process",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "A policy clause requiring evidence within 30 days of the loss event imposes a duty on the policyholder to provide evidence promptly — it does not impose a deadline on when the insurer can request evidence; a request for a formal interview is reasonable where simultaneous theft and fire claims exist on the same policy; a limited company cannot claim distress — compensation is limited to inconvenience; a single formal interview request (distinguished from multiple requests) falls within the policyholder's general obligation to provide reasonably required evidence; insurer's claim investigation obligations persist after policy expiry if the claim was made during the policy period",
        "Missing Evidence": "CCTV footage content (recovered after NIG's final response); outcome of formal interview with M's director (not yet conducted at time of FOS decision)",
        "Ombudsman Reasoning": "Simultaneous theft and fire claims on same policy justify detailed investigation; M caused delay by not providing keys for over 3 months — no claim decision appropriate until keys provided; CCTV only recovered after final response so reasonable no decision was made; 30-day clause in policy applies to policyholder's duty to provide evidence, not to when insurer can request it; one formal interview is reasonable in complex circumstances; £350 adequate for inconvenience; limited company cannot experience distress; complaint not upheld",
        "Workflow Insight": "Where two serious claims occur close together on the same policy (theft and fire), insurers are entitled to conduct detailed investigations including requesting formal interviews; policyholders who delay providing required evidence cannot subsequently complain about claim delays; a single formal interview request in complex multi-claim circumstances is reasonable; a company complainant cannot receive distress compensation — awards limited to inconvenience",
        "AI Rule Candidate": "IF simultaneous_theft_and_fire_claims THEN more_detailed_investigation_including_formal_interview_is_justified; IF policyholder_delays_providing_required_evidence THEN insurer_not_responsible_for_resulting_delay; policy_30_day_evidence_provision_imposes_duty_on_policyholder_not_deadline_on_insurer; IF complainant_is_limited_company THEN distress_award_not_available_only_inconvenience",
        "Source PDF": "DRN-4800291.pdf",
    },
    {
        "Case ID": "THEFT-025",
        "FOS Decision ID": "DRN-4979740",
        "Insurer Name": "Admiral Insurance (Gibraltar) Limited",
        "FOS Decision Date": "25 Oct 2024",
        "Claim Type": "Motor (car) insurance — home burglary February 2024; Mrs G's car keys stolen and car taken; car recovered by police within one hour with no damage; Mrs G arranged for car lock to be recoded and locks replaced; she wanted to claim under the policy 'lost or stolen keys' benefit (section 2.4: up to £500 reimbursement, no excess, no NCD impact); Admiral insisted on treating as a full theft claim under section 2.1 (excess payable, NCD affected)",
        "Entry / Theft Method": "Theft of car keys during home burglary — home broken into while occupants slept; car keys stolen and car taken but recovered by police within one hour with no damage; Mrs G arranged for car lock to be recoded and house and car locks replaced",
        "Property Type": "Motor vehicle (personal car insurance)",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Accepted — Disputed Settlement",
        "Rejection Reason": "Not applicable — claim not declined; Admiral accepted a claim but insisted on applying section 2.1 (full theft terms with excess and NCD impact) rather than section 2.4 (lost/stolen keys benefit: no excess, no NCD impact, up to £500)",
        "Evidence Dispute": "Mrs G: only claiming for key/lock replacement; car was recovered before she contacted Admiral; section 2.4 does not exclude itself because the car was also stolen. Admiral: section 2.1 must apply because the car was stolen — insisted on full theft terms. Investigator agreed with Mrs G. FOS: car recovered before Mrs G contacted Admiral; Admiral has not dealt with a vehicle theft loss; Mrs G only claims for key/lock replacement under section 2.4; no policy wording prevents section 2.4 from being used in these circumstances; applying section 2.1 with harsher terms is unfair; complaint upheld",
        "Outcome Category": "Upheld",
        "Outcome": "Admiral must meet Mrs G's claim under section 2.4 of the policy — no excess, no NCD impact, up to £500 reimbursement; Admiral to pay 8% simple interest per annum on reimbursement from date Mrs G paid to date Admiral reimburses; Mrs G must provide proof of payment",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "Where a motor policy provides a distinct 'lost or stolen keys' benefit (no excess, no NCD impact) and the policyholder only claims for key/lock replacement, the insurer must apply that specific benefit rather than the broader theft section — even where the car was also stolen — provided the car has been recovered and no vehicle loss or damage claim is being made; an insurer cannot apply harsher policy terms in place of a specific named benefit unless the policy wording explicitly excludes the benefit in those circumstances; when a vehicle is recovered before the claim is made, the insurer has not dealt with a 'theft claim' for the vehicle and cannot rely on the theft event to override the keys-only benefit",
        "Missing Evidence": "None material — dispute was entirely one of policy interpretation; Mrs G must provide proof of payment for Admiral's reimbursement",
        "Ombudsman Reasoning": "Section 2.4 covers lost/stolen keys with no excess and no NCD impact if claiming only under that benefit; car recovered by police before Mrs G contacted Admiral; no loss or damage to car; Mrs G only claiming for key/lock replacement; Admiral has not had to deal with a vehicle theft loss; applying section 2.1 with excess and NCD impact is unfair when no vehicle loss claim is being made; policy does not exclude section 2.4 where car was also stolen but recovered; complaint upheld; 8% simple interest on reimbursement",
        "Workflow Insight": "Handlers must identify the specific policy benefit that actually applies to the consumer's loss — where keys are stolen but the vehicle is recovered undamaged and no vehicle loss/damage claim is made, the 'lost/stolen keys' benefit applies; the mere fact that a vehicle was temporarily stolen does not convert a key-replacement claim into a full theft claim where the vehicle is recovered before the claim is made; policyholders should be informed that claiming under the keys-only benefit preserves their NCD",
        "AI Rule Candidate": "IF consumer_claims_only_for_lost_stolen_keys AND vehicle_recovered_before_claim_made AND no_vehicle_loss_or_damage_claimed THEN lost_stolen_keys_policy_benefit_applies_not_full_theft_section; insurer_cannot_apply_full_theft_terms_if_policy_does_not_explicitly_exclude_keys_benefit_in_theft_circumstances; IF vehicle_recovered_before_insurer_dealt_with_theft THEN insurer_has_not_processed_theft_claim_and_cannot_override_keys_benefit",
        "Source PDF": "DRN-4979740.pdf",
    },
    {
        "Case ID": "THEFT-026",
        "FOS Decision ID": "DRN-5011646",
        "Insurer Name": "Mulsanne Insurance Company Limited",
        "FOS Decision Date": "25 Oct 2024",
        "Claim Type": "Motor (car) insurance — car stolen March 2023; Mulsanne declined because the policy and schedule required a Thatcham-approved tracking device as a condition of theft cover; vehicle had no tracker fitted; Dr M also has a separate complaint about the sales process that led to this policy term (handled separately as a different regulated activity)",
        "Entry / Theft Method": "Motor vehicle theft — car stolen; specific method of entry not described; the policy security requirement (Thatcham-approved tracker) was not met at time of loss",
        "Property Type": "Motor vehicle (personal car insurance)",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Policy wording stated theft claims only covered if required security requirements are met; Dr M's schedule specifically required a Thatcham-approved tracking device; no tracker was fitted to the vehicle; security requirement not met; complaint about the sales process is addressed separately as a different regulated activity",
        "Evidence Dispute": "Dr M: when the policy was taken out, he did not agree that the car had a tracker; the sales process was unfair. Mulsanne: policy and schedule clearly require Thatcham-approved tracker for theft cover; no tracker fitted — not disputed. FOS: policy wording and schedule are clear; no tracker fitted — undisputed; sales process complaint is a separate regulated matter handled separately; on this claim-only complaint, Mulsanne applied the policy terms fairly; complaint not upheld",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — Mulsanne fairly and reasonably declined the theft claim; policy terms and schedule clearly required a Thatcham-approved tracking device which was not present",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "Where a policy schedule specifically requires a Thatcham-approved tracking device as a condition of theft cover, the absence of any tracker means theft cover does not apply — even if the policyholder disputes the sales process that led to that requirement; FOS treats the sales process complaint and the claim decision complaint as separate regulated activities; on a claim-only complaint, FOS assesses whether the insurer fairly applied the policy terms as they exist, not whether those terms should have been included; a complaint about the fairness of the sales process that generated a policy term does not affect the enforceability of that term in a claim decision",
        "Missing Evidence": "Whether Dr M was clearly informed at point of sale that a tracker was required (addressed in the separate sales complaint, not in this decision)",
        "Ombudsman Reasoning": "Policy says theft claims only covered if required security requirements met; schedule specifies Thatcham-approved tracking device; no tracker fitted — not disputed by Dr M; Dr M's concern about the sales process is being handled separately; on the current complaint about the claim decision, Mulsanne applied the policy terms fairly; complaint not upheld",
        "Workflow Insight": "A tracker endorsement on a motor policy is an absolute condition of theft cover — there is no dispensation for vehicles where the policyholder disputes having agreed to the requirement during the sales process; FOS treats the sales process complaint and the claim decision complaint as distinct regulated activities; insurers can rely on clear policy schedule terms in a claim complaint even when the sales process is disputed in a parallel complaint",
        "AI Rule Candidate": "IF policy_schedule_requires_thatcham_tracker AND no_tracker_fitted THEN theft_claim_excluded; sales_process_complaint_is_separate_regulated_activity_from_claim_decision_complaint; IF insurer_applies_clear_policy_schedule_endorsement THEN claim_decline_is_fair_regardless_of_parallel_sales_process_dispute",
        "Source PDF": "DRN-5011646.pdf",
    },
    {
        "Case ID": "THEFT-027",
        "FOS Decision ID": "DRN-5275985",
        "Insurer Name": "Wakam",
        "FOS Decision Date": "5 Apr 2025",
        "Claim Type": "Motor (motorhome) insurance — motorhome stolen from outdoor storage yard; Wakam repudiated the claim citing significant corrosion (rust) in breach of the policy condition 'You shall maintain Your Motorhome in efficient condition'; motorhome later recovered badly damaged; Wakam subsequently conceded it could not fairly rely on mileage discrepancy or misrepresentation grounds; FOS found the rust was not causally connected to the theft under ICOBS 8.1.2(B); £400 D&I compensation ordered",
        "Entry / Theft Method": "Motor vehicle theft (motorhome) — motorhome stolen from outdoor storage yard; specific method of theft not described; vehicle later recovered with extensive internal damage and items stolen from it",
        "Property Type": "Motor vehicle (personal motorhome insurance)",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Policy condition requiring motorhome to be maintained in 'efficient condition' allegedly breached due to significant corrosion (rust) visible on post-recovery inspection; Wakam argued rust was pre-existing and should have been noticed at the last MOT; Wakam also initially relied on mileage discrepancies and possible misrepresentation but later conceded these grounds could not be fairly relied upon",
        "Evidence Dispute": "Mr S: motorhome passed MOT on 31 March 2023 with no advisory notices relating to rust; rust is characteristic of a 30-year-old vehicle stored outdoors; he had invested significantly in the motorhome. Wakam: inspection photos show substantial corrosion from scale rust to penetrating rust; rust of this severity would have been visible at last MOT. FOS: MOT passed March 2023 with no rust advisories; 14 months elapsed between MOT and Wakam's post-recovery inspection; rust may have developed or worsened in that time; even if rust breached the condition, Wakam has not shown causal connection between the rust and the theft (ICOBS 8.1.2(B)); mileage and misrepresentation grounds conceded by Wakam; pre-existing rust deductible from indemnity settlement; £400 D&I; 8% interest from initial rejection",
        "Outcome Category": "Upheld",
        "Outcome": "Wakam must reconsider the theft claim in line with remaining policy terms and limits, including damage to the motorhome and stolen accessories/equipment; pre-existing rust damage (identified by Wakam's engineer's report) may be deducted from any settlement; 8% simple interest per annum from date of Wakam's initial rejection to date of settlement; £400 D&I compensation for avoidable trouble and upset caused by Wakam's unreasonable handling",
        "Compensation Awarded (£)": 400,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "Under ICOBS 8.1.2(B), a rejection of a consumer policyholder's claim for breach of a policy condition is unreasonable unless the circumstances of the claim are connected to the breach; an insurer cannot rely on a vehicle maintenance condition to repudiate a theft claim without demonstrating causal connection between the breach and the theft itself; a recent MOT pass with no rust advisories is strong evidence that rust visible at post-theft inspection may have developed or worsened after the MOT; a stolen vehicle that is later recovered remains a theft claim in respect of post-theft damage — the insurer cannot reclassify it as an accidental damage total loss claim simply because the vehicle was recovered; pre-existing unrelated damage (rust) may be deducted from any theft indemnity settlement but the claim cannot be fully declined without causal connection to the loss",
        "Missing Evidence": "Evidence establishing causal connection between the corrosion and the theft (e.g. that rust around locks facilitated entry) — absent and critical; without this causal link ICOBS 8.1.2(B) prevents reliance on the condition breach to decline the claim",
        "Ombudsman Reasoning": "Motorhome passed MOT 31 March 2023 with no rust advisories; 14 months before Wakam's post-recovery inspection — rust may have developed or worsened; Wakam cannot persuasively demonstrate causal connection between rust and the theft; ICOBS 8.1.2(B) requires causal connection between breach and loss; mileage and misrepresentation grounds conceded; vehicle remains a theft claim even after recovery; pre-existing rust excluded from settlement; Wakam's handling caused avoidable distress — £400 D&I; 8% interest from initial rejection to settlement date",
        "Workflow Insight": "Before repudiating a theft claim on a general vehicle condition clause, insurers must affirmatively establish causal connection between the breach and the theft — ICOBS 8.1.2(B) prevents reliance on unconnected breaches; a recent MOT with no relevant advisories substantially weakens any claim that a defect was pre-existing; post-theft inspection findings must be compared to the last MOT record before asserting a pre-existing breach; the indemnity principle allows pre-existing unrelated damage to be excluded from settlement but does not permit the entire claim to be declined without causal connection to the loss",
        "AI Rule Candidate": "IF claim_repudiated_for_policy_condition_breach AND breach_not_causally_connected_to_theft THEN repudiation_unreasonable_under_ICOBS_8_1_2_B; IF vehicle_passed_recent_MOT_without_relevant_advisories THEN insurer_cannot_assume_defect_was_pre_existing_without_further_evidence; pre_existing_unrelated_damage_may_be_deducted_from_theft_settlement_but_full_decline_not_permitted_without_causal_connection; stolen_and_recovered_vehicle_remains_theft_claim_for_post_theft_damage",
        "Source PDF": "DRN-5275985.pdf",
    },
    {
        "Case ID": "THEFT-028",
        "FOS Decision ID": "DRN5412255",
        "Insurer Name": "Lloyds Bank Insurance Services Limited",
        "FOS Decision Date": "Not stated in document",
        "Claim Type": "Home insurance — policy brokered by Lloyds TSB originally voided by underwriter for alleged non-disclosure; Lloyds TSB later admitted the mis-sale and provided 12 months' own-policy cover at the original premium plus £100 D&I compensation; Ms H made a theft claim under the remedial policy (accepted and settled); Ms H subsequently struggled to find affordable renewal insurance in the open market, alleging Lloyds TSB bore ongoing responsibility for higher premiums caused by her theft claim now on record",
        "Entry / Theft Method": "Home theft — specific theft circumstances not described in this decision; theft claim was accepted and settled under the remedial Lloyds TSB policy; the dispute concerns the downstream effect on premium costs, not the theft event itself",
        "Property Type": "Residential home",
        "Dispute Type": "Claim Recording / Administrative Dispute",
        "Coverage Decision": "Accepted",
        "Rejection Reason": "Not applicable — theft claim was accepted and settled during the remedial policy period; dispute concerns whether Lloyds TSB bears ongoing responsibility for higher market premiums resulting from the legitimately-recorded theft claim",
        "Evidence Dispute": "Ms H: Lloyds TSB as broker mis-sold the original policy; should have offered renewal at the original premium; theft claim on her record causes higher premiums in the market and at renewal; Lloyds TSB should bear responsibility. Lloyds TSB: the mis-sale was fully remedied — 12 months' cover at original premium, £100 D&I, CUE voidance record cleared; theft claim is legitimate and not connected to the original error; premium pricing is each insurer's commercial judgment. FOS: Lloyds TSB fully remedied the mis-sale; CUE voidance record confirmed cleared; theft claim on remedial policy is a legitimate claims record not connected to the original error; market and renewal premium pricing is each insurer's commercial judgment — FOS cannot interfere; complaint not upheld",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — Lloyds TSB appropriately remedied the original policy mis-sale; not responsible for market premium increases arising from the legitimately-made and recorded theft claim",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Broker Dispute",
        "Key Policy Clause": "A broker who mis-sells a policy discharges its remedial obligations by providing equivalent cover at the original premium and paying appropriate D&I compensation — it does not bear ongoing responsibility for how the market prices premiums reflecting a legitimately-made claim on the remedial policy; a theft claim made and accepted on a remedial policy is a legitimate claims record not attributable to the broker's original error once that error has been fully remedied; insurers have commercial discretion to price premiums at renewal based on claims history — FOS cannot interfere; once a CUE voidance/avoidance record is confirmed to have been removed, the broker has no further liability arising from the original policy avoidance",
        "Missing Evidence": "Not applicable — factual background not disputed; dispute concerned the legal extent of Lloyds TSB's ongoing responsibility after fully remediating the mis-sale",
        "Ombudsman Reasoning": "Lloyds TSB admitted the mis-sale and offered appropriate remediation (12 months' equivalent cover at original premium, £100 D&I); CUE avoidance record confirmed cleared; Ms H chose not to renew with Lloyds TSB as renewal premium increased — Lloyds TSB has commercial discretion over renewal pricing; open market insurers pricing higher due to theft claim history is their commercial judgment not attributable to Lloyds TSB's original error (which was remedied); the theft claim's existence on Ms H's record is not connected to the mis-sale; complaint not upheld",
        "Workflow Insight": "Where a broker mis-sells a policy, full remediation includes correcting any adverse CUE records arising from the mis-sale — once this is done and equivalent cover is provided at the original premium, the broker's obligations are discharged; a theft claim made on a remedial policy creates a legitimate claims record that the policyholder must account for in the open market — this is not attributable to the original broker error; FOS does not interfere in how insurers price premiums at renewal or in the open market",
        "AI Rule Candidate": "IF broker_provides_remedial_cover_at_original_premium_and_pays_DI_compensation AND corrects_CUE_record THEN broker_has_discharged_mis_sale_remediation_obligations; theft_claim_legitimately_made_on_remedial_policy_creates_valid_claims_record_not_attributable_to_original_broker_error; insurer_renewal_and_market_premium_pricing_is_commercial_judgment_FOS_cannot_interfere",
        "Source PDF": "DRN5412255.pdf",
    },
    {
        "Case ID": "THEFT-029",
        "FOS Decision ID": "DRN-5704833",
        "Insurer Name": "Complete Cover Group Ltd",
        "FOS Decision Date": "24 Oct 2025",
        "Claim Type": "Motor (car) insurance sold through broker Complete Cover Group Ltd — car stolen; insurer declined because the tracker subscription was not activated (tracker device was physically present on the car but the subscription was not continuously active as required by the endorsement); Mr B's representative complained that Complete Cover did not adequately explain the requirement to activate the subscription or how to do so",
        "Entry / Theft Method": "Motor vehicle theft — car stolen; method of theft not described; the endorsement condition not met because tracker subscription was not activated despite the tracker device being physically present",
        "Property Type": "Motor vehicle (personal car insurance)",
        "Dispute Type": "Broker Conduct Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Insurer declined because tracker subscription was not continuously active at time of theft loss; endorsement required: (1) Thatcham-approved tracking device; (2) subscription continuously active at time of loss; (3) subscription registered in policyholder's or spouse's name; (4) access to subscription account — condition (2) was not met as subscription was not activated",
        "Evidence Dispute": "Mr B's representative: Complete Cover did not draw the tracker subscription activation requirement sufficiently to attention; did not explain steps needed to activate; Mr B believed the subscription came with the car at purchase. Complete Cover: endorsement requirements shown at quotation with mandatory non-prepopulated checkbox acknowledgement; welcome letter clearly stated subscription must be continuously active; policy handbook gave full detail; it is not a broker's responsibility to provide vehicle-specific tracker activation instructions. FOS: requirement disclosed at multiple stages (mandatory checkbox, welcome letter, handbook); clearly presented in the welcome schedule document; it was Mr B's personal responsibility to ensure the subscription was active; broker not required to provide vehicle-specific activation instructions; Mr B could have sought clarification but did not; complaint not upheld",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — Complete Cover adequately disclosed the tracker subscription endorsement requirement at multiple stages; it is not a broker's responsibility to provide vehicle-specific tracker activation instructions",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Broker Dispute",
        "Key Policy Clause": "A broker adequately discloses a tracker subscription endorsement where: (1) the requirement is shown at quotation with a mandatory non-prepopulated acknowledgement checkbox the consumer must actively tick; (2) the welcome letter explicitly highlights it; (3) the policy handbook provides full detail; a broker is not required to provide vehicle-specific instructions for activating tracker subscriptions — that is the policyholder's personal responsibility; where endorsement requirements are clearly and repeatedly disclosed and the consumer fails to seek clarification, the broker is not liable for the insurer's resulting claim decline",
        "Missing Evidence": "Evidence that Mr B's tracker subscription was or was not active at the time of theft; explanation of why any subscription 'coming with the car' at purchase was not activated or was not registered in Mr B's own name",
        "Ombudsman Reasoning": "Complete Cover presented tracker endorsement at quotation with a mandatory non-prepopulated checkbox — Mr B had to actively acknowledge it before proceeding; welcome letter clearly stated subscription must be continuously active; policy handbook gave full detail; requirement not hidden — clearly in the welcome policy schedule document; Mr B's representative said he thought subscription came with the car — Complete Cover is not responsible for this assumption; broker not required to provide vehicle-specific activation steps for each customer type; Mr B could have contacted Complete Cover for clarification but did not; complaint not upheld",
        "Workflow Insight": "Brokers can satisfy tracker endorsement disclosure obligations through a combination of mandatory acknowledgement at quotation, clear welcome letter and detailed policy handbook; policyholders bear personal responsibility for ensuring tracker subscriptions are continuously active — this includes verifying that any subscription 'coming with the car' is activated and registered in the correct name; a policyholder who does not seek clarification on endorsement requirements cannot later hold the broker responsible for a resulting claim decline",
        "AI Rule Candidate": "IF tracker_endorsement_disclosed_at_quotation_via_mandatory_non_prepopulated_checkbox AND welcome_letter_and_handbook_also_reference_requirement THEN broker_has_adequately_disclosed_endorsement; IF consumer_fails_to_activate_tracker_subscription AND no_evidence_of_seeking_clarification THEN broker_not_liable_for_resulting_claim_decline; broker_not_required_to_provide_vehicle_specific_tracker_activation_instructions",
        "Source PDF": "DRN-5704833.pdf",
    },
    {
        "Case ID": "THEFT-030",
        "FOS Decision ID": "DRN-5852242",
        "Insurer Name": "U K Insurance Limited trading as NIG",
        "FOS Decision Date": "29 Oct 2025",
        "Claim Type": "Commercial property insurance — insolvent tenant (tenant two) vacated a commercial let property after becoming insolvent, making unauthorised alterations (partial redecoration, removal of light fixtures causing small holes, removal of aluminium kitchen sheeting) and removing fixtures and fittings; J (limited company landlord) claimed approximately £14,000 for malicious damage and approximately £8,300 for loss of rent; UKI declined — tenant's actions did not constitute 'malicious damage'; FOS agreed; no formal theft claim for removed items was ever submitted to UKI",
        "Entry / Theft Method": "Removal of fixtures and fittings by insolvent departing commercial tenant — tenant two vacated after insolvency; made unauthorised changes (partially painted walls, removed light fixtures, removed aluminium kitchen cladding, possibly removed coffee machine); no forced entry; tenant had legitimate access to the property under the lease",
        "Property Type": "Commercial (commercial let property insurance — not home/contents)",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Damage caused by an insolvent tenant vacating and making unauthorised alterations does not constitute 'malicious damage' — malicious damage requires deliberate intention to cause harm directed at the insured or the property; tenant's actions were more consistent with self-interest or refurbishment purposes than with spite or ill will toward J; without an accepted insured event, the loss of rent claim cannot be pursued; no formal theft claim was made to UKI so FOS cannot direct consideration of one",
        "Evidence Dispute": "J: tenant made unauthorised alterations and removals causing £14,000 damage constituting malicious damage — actions were intentional and caused financial harm; tenant has history of insolvency and not paying staff showing bad character. UKI: 'malicious' requires deliberate intent to cause harm — breach of lease insufficient. Investigator: agreed not malicious damage; recommended UKI consider a theft claim for removed items. FOS: damage not malicious — tenant acting in self-interest not with ill will toward J; partial redecoration consistent with refurbishment; small holes from light fixtures not evidence of intent to harm; aluminium cladding partially left on-site — more consistent with refurbishment than malice; tenant insolvency history does not establish intent to harm J; neglect or acting without care ≠ malicious damage; FOS cannot direct insurer to consider theft claim not formally made or complained about; complaint not upheld",
        "Outcome Category": "Not Upheld",
        "Outcome": "Complaint not upheld — UKI fairly declined the malicious damage claim; FOS will not require UKI to consider a theft claim that was never formally made as part of this complaint; J should make a formal theft claim to UKI",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Commercial",
        "Key Policy Clause": "'Malicious damage' under a commercial property policy requires deliberate intention to cause harm to the insured property or the insured landlord — breach of tenancy agreement and unauthorised alterations are insufficient without positive evidence of spite or ill will directed at the insured; acting without care, or neglect, does not evidence malicious damage unless there is also deliberate intent to cause harm; a tenant acting in its own interests (removing items for use elsewhere, refurbishment) is not acting maliciously even if this causes financial harm to the landlord; loss of rent cover depends on an accepted insured event being established; FOS can only consider complaints about claims formally made and declined or ignored by the insurer — it cannot require an insurer to consider a claim never formally submitted",
        "Missing Evidence": "Evidence of deliberate intent to harm J by tenant two (absent — all actions appear self-interested); formal theft claim submitted to UKI and UKI's response (not before FOS; J directed to make this claim directly to UKI)",
        "Ombudsman Reasoning": "Policy covers 'theft and malicious damage by tenant'; malicious damage means damage with deliberate intention to cause harm; partial redecoration consistent with refurbishment not malice; small holes from light fixture removal not evidence of intent to harm; aluminium sheeting partly left on site — more consistent with refurbishment; tenant removing items for another business shows self-interest not ill will; tenant insolvency history does not establish intent to harm J; neglect or acting without care ≠ malicious damage; no accepted insured event so no loss of rent; no formal theft claim made or complained about — FOS cannot review what was never before the insurer as a complaint; complaint not upheld",
        "Workflow Insight": "Landlords whose commercial tenants cause damage on vacating must formally make both a malicious damage claim and a separate theft claim — failing to formally make a theft claim means FOS cannot consider it; the 'malicious' element in a commercial property policy requires positive evidence of spite or ill will — significant financial harm to the landlord alone is insufficient; handlers should distinguish between damage serving the tenant's interests (acting without care) and damage intended to harm the landlord (malicious)",
        "AI Rule Candidate": "IF tenant_causes_damage_or_removes_items_on_vacating AND actions_serve_tenants_own_interests THEN malicious_damage_cover_does_not_apply; malicious_damage_requires_deliberate_intent_to_cause_harm_not_merely_breach_of_tenancy_or_financial_harm_to_landlord; IF loss_of_rent_claimed AND no_accepted_insured_event THEN loss_of_rent_claim_cannot_proceed; FOS_cannot_direct_insurer_to_consider_theft_claim_not_formally_made_and_not_complained_about",
        "Source PDF": "DRN-5852242.pdf",
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
