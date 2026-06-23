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
# NEW CASES — Batch 2: SUBS-011 to SUBS-020
# ---------------------------------------------------------------------------
NEW_CASES: list[dict] = [
    {
        "Case ID": "SUBS-011",
        "FOS Decision ID": "DRN2951368",
        "Insurer Name": "Covea Insurance Plc",
        "FOS Decision Date": "30 Oct 2020",
        "Claim Type": "Home insurance — subsidence claim; inadequate investigation in 2010 (superficial inspection only) led to Covea applying the £15,000 excess introduced in 2011 rather than the £1,000 excess applicable in 2010; consumer argues the 2017 tree-root subsidence is a continuation of the 2010 claim",
        "Movement Cause": "Tree root-induced clay soil shrinkage at front of property; relationship between weather conditions and tree roots caused cyclical cracking that first manifested before 2010; tree removed and movement stabilised after 2017 investigation",
        "Property Type": "Residential home",
        "Dispute Type": "Handling / Reinstatement Dispute",
        "Coverage Decision": "Accepted — Disputed Settlement",
        "Rejection Reason": "Covea: three distinct claims from three causes (1998, 2010, 2017); 2017 claim is subject to the £15,000 excess in force at that time; loss adjuster C found only differential movement in 2010 that was not worsening; tree implicated only from hot summer 2016 — a new event; CUE should record three claims",
        "Evidence Dispute": "Consumer: C's inspection in 2010 was superficial (three-day turnaround, no in-depth investigation); C subsequently said tree impact on 2010 cracks could not be ruled out; prior experts had recommended monitoring. Covea: original underpinning repairs (2002/3) remained in good order; C found differential movement in 2010 within the excess; 2017 investigation found tree impact from 2016 hot summer; three separate causes. FOS: prior history and experts' recommendations warranted more than a superficial investigation in 2010; C acknowledged tree impact on 2010 cracks could not be ruled out; tree root clay shrinkage is cyclical and weather-dependent — absence of worsening in one period does not confirm absence of cause; Covea has not shown on balance that 2010 cracks were unrelated to the tree; likely continuation from 2010; finely balanced but £1,000 excess is fair",
        "Outcome Category": "Upheld",
        "Outcome": "Record the 2017 claim as commencing in 2010; amend the CUE database to record two subsidence claims not three; apply the £1,000 excess applicable in 2010 (future claims remain subject to the current excess)",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Where a loss adjuster's inadequate investigation of an earlier claim prejudices the excess the consumer pays on a later related claim, FOS may direct the later claim be treated as a continuation of the earlier and apply the earlier excess; insurer duty under ICOBS to handle claims promptly and fairly, not unreasonably reject; tree root clay shrinkage is cyclical — weather-driven fluctuations in visible movement do not break the causal chain; insurer bears the burden of demonstrating that the later damage is a distinct new event to justify the higher excess applicable at the time of the later claim; CUE database must accurately reflect the number of distinct claims",
        "Missing Evidence": "Detailed investigation reports from the 2010 claim (none produced — C completed only a superficial inspection with no exploratory investigation within three days of attendance)",
        "Ombudsman Reasoning": "Property had prior subsidence history and experts had recommended further monitoring; Covea chose a new loss adjuster specifically to avoid thorough investigation; C's 2010 inspection was superficial — no in-depth investigation despite that history; C subsequently acknowledged tree impact on 2010 cracks could not be ruled out; tree root clay shrinkage is cyclical — not a constant; Covea has not shown on balance that 2010 cracks were unrelated to the tree; likely continuation from 2010 given original underpinning repairs (2002/3) remained in good order; finely balanced but fair outcome is £1,000 excess and two CUE entries",
        "Workflow Insight": "Loss adjusters instructed on a property with prior subsidence history must carry out a thorough investigation — a superficial visit that produces a final determination within three days is likely to be found inadequate by FOS; where an inadequate investigation in an earlier claim later prejudices the excess payable on a continuation claim, FOS will direct the higher excess to be replaced by the lower one applicable at the time of the earlier claim; tree root clay shrinkage claims are inherently cyclical — variation in movement between dry and wet years does not establish a new distinct event; insurer must demonstrate on balance that later damage is a new and distinct event before applying the higher current-period excess",
        "AI Rule Candidate": "IF inadequate_investigation_of_prior_claim AND later_damage_likely_related_by_same_cause THEN insurer_may_be_directed_to_treat_later_claim_as_continuation_applying_earlier_excess; IF property_has_prior_subsidence_history AND experts_recommended_further_investigation THEN superficial_inspection_is_inadequate; tree_root_clay_shrinkage_is_cyclical_so_absence_of_worsening_in_one_period_does_not_confirm_absence_of_cause; insurer_must_demonstrate_on_balance_that_later_damage_is_distinct_new_event_to_apply_higher_excess_applicable_at_later_date",
        "Source PDF": "DRN2951368.pdf",
    },
    {
        "Case ID": "SUBS-012",
        "FOS Decision ID": "DRN-3258437",
        "Insurer Name": "Red Sands Insurance Company (Europe) Limited",
        "FOS Decision Date": "2 Jun 2022",
        "Claim Type": "Home insurance — subsidence claim made in 2020 (same year as inception); Red Sands removed subsidence cover and declined claim on grounds that Mr P could not produce the specific certificate of adequacy or chartered surveyor's report required by the policy question, and therefore made a misrepresentation under CIDRA 2012",
        "Movement Cause": "Subsidence (cause and extent not concluded — loss adjuster had suggested monitoring before claim was declined when subsidence cover was removed; investigation incomplete)",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Red Sands: Mr P selected 'am able to provide' certificate of adequacy or chartered surveyor's report but could not produce either specific document; Red Sands would not have offered subsidence cover without those documents; Mr P should know more than a layman given his experience buying a property with prior subsidence history; damage may be linked to previous problems",
        "Evidence Dispute": "Mr P: openly disclosed prior subsidence history; believed his available evidence (confirmed accepted claim, loss adjuster explanation, drainage work completion, 2014 RICS mortgage valuation showing no structural movement) fulfilled the requirement; answered to best of knowledge as a layman. Red Sands: specific documentation required to manage risk; consumer could have used email or chat to clarify; consumer's experience of buying a subsidence-history property elevates his knowledge beyond layman level. FOS: CIDRA test is reasonable care not to misrepresent, not strict correctness; Mr P's evidence achieves broadly the same purpose as a certificate of adequacy — shows prior problem was dealt with normally; RICS mortgage valuation finding no structural movement goes further than a certificate would; buying a property with prior subsidence history does not import professional knowledge of subsidence or insurance; Red Sands is not disadvantaged by absence of a certificate format",
        "Outcome Category": "Upheld",
        "Outcome": "Red Sands to reinstate subsidence cover; consider the claim subject to policy terms and conditions (claim requires further investigation and monitoring before outcome can be determined)",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "CIDRA 2012 — test is reasonable care not to misrepresent, not strict correctness; where consumer's evidence achieves the same substantive purpose as specifically requested documentation (certificate of adequacy), insurer is not disadvantaged by absence of exact format and cannot establish a qualifying misrepresentation; RICS mortgage valuation finding no structural movement may provide broader reassurance than a certificate of adequacy; purchasing a property with a subsidence history does not import professional knowledge of buildings, subsidence or insurance claims — consumer remains a layman unless specific expertise is demonstrated",
        "Missing Evidence": "Certificate of adequacy from original insurer (not obtained); chartered surveyor's structural report (not obtained); conclusion on whether 2020 damage is linked to prior subsidence (monitoring required — investigation was abandoned when cover was removed)",
        "Ombudsman Reasoning": "CIDRA test is reasonable care; Mr P's documents show prior subsidence claim accepted by an insurer, cause identified (leaking drains), drainage work completed, RICS survey found no structural movement — this substantially achieves the purpose of a certificate of adequacy; Red Sands no more disadvantaged than if a certificate had been provided; Mr P is a layman — buying a house with subsidence history does not make him more expert in subsidence or insurance; not satisfied he failed his CIDRA duty; Red Sands has no CIDRA remedy; claim must be properly investigated",
        "Workflow Insight": "When relying on CIDRA to remove subsidence cover, insurer must show the consumer's answer failed to meet the 'reasonable care' standard — mere technical non-compliance with a document format requirement is insufficient; where the consumer's available evidence achieves the same substantive purpose as the specific document requested, the CIDRA test is met even if that exact format is absent; consumer's occupational or life experience importing relevant expertise must be specifically demonstrated — the fact of buying a property with subsidence history alone does not elevate the consumer above layman status; when removing cover and declining a claim simultaneously, the claim investigation should not be abandoned — FOS may require both steps to be undone",
        "AI Rule Candidate": "IF consumer_evidence_achieves_same_purpose_as_requested_document THEN CIDRA_reasonable_care_test_met_even_if_exact_format_absent; IF consumer_purchased_property_with_prior_subsidence_history THEN consumer_remains_layman_unless_specific_professional_expertise_demonstrated; IF insurer_removes_cover_under_CIDRA AND claim_investigation_incomplete THEN insurer_directed_to_reinstate_cover_and_complete_investigation; CIDRA_test_is_reasonable_care_not_strict_correctness",
        "Source PDF": "DRN-3258437.pdf",
    },
    {
        "Case ID": "SUBS-013",
        "FOS Decision ID": "DRN-3387540",
        "Insurer Name": "Accredited Insurance (Europe) Limited",
        "FOS Decision Date": "17 May 2022",
        "Claim Type": "Home insurance — new cracks noticed in 2020 (year of policy inception); AIL declined claim for pre-existing/maintenance-related damage exclusion; separately argued claim cost below £2,500 excess; also failed to renew policy after declining claim, leaving consumer without subsidence cover",
        "Movement Cause": "Subsidence caused by trees (AIL identified as cause after inspection; tree investigation and crack monitoring recommended before claim was declined)",
        "Property Type": "Residential home",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "AIL: (1) damage pre-existing or maintenance-related — homebuyer's survey noted property needed repair and prior subsidence had occurred; policy excludes pre-inception events; (2) schedule of works £1,900 is below £2,500 subsidence excess",
        "Evidence Dispute": "Mr P: discovered new cracks August 2020; prior subsidence (approx 10 years earlier) was repaired by insurer; homebuyer's survey found no significant structural defects, no active movement. AIL: homebuyer's survey noted prior subsidence; inspection showed internal cracks (photos not annotated). FOS: excess argument — total claim value includes investigation and monitoring costs (likely >£600) plus repair costs (£1,900) so overall value likely exceeds £2,500 excess; pre-existing exclusion — AIL's own inspection report did not comment on when subsidence started; homebuyer's survey found no active movement; no expert evidence that current subsidence pre-existed the January 2020 inception; burden of proof on AIL as party relying on exclusion; AIL failed to renew after declining claim — obligation to continue cover after accepted claim applies; wrongful denial effectively deprived consumer of subsidence cover at renewals 2021 and 2022",
        "Outcome Category": "Upheld",
        "Outcome": "Accept the subsidence claim; offer cover including subsidence from next renewal and confirm subsidence cover from January 2021 to January 2023; pay the difference between what Mr P paid for alternative cover and what he would have paid remaining with AIL; pay £500 compensation; consider consumer's professional costs (legal, arboricultural, tree surgery) incurred progressing claim after wrongful decline — confirm which are refunded or offset against excess",
        "Compensation Awarded (£)": 500,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Pre-existing damage exclusion — burden of proof on insurer relying on exclusion to show damage existed before inception; homebuyer's survey finding no active movement is strong evidence against pre-inception subsidence; excess applies to total claim value including investigation and monitoring costs, not just repair cost alone; insurer obliged to continue offering subsidence cover at renewal after accepting a claim; wrongful claim denial that leads to consumer losing subsidence cover at subsequent renewals creates liability for increased premium costs consumer paid to alternative insurers; consumer professional costs (surveyor, arboricultural, tree surgery) incurred in progressing a claim after wrongful denial may be refunded or offset against the excess",
        "Missing Evidence": "Expert opinion on when current subsidence began relative to policy inception; AIL inspection report timing commentary (not provided); confirmation of tree investigation and monitoring cost to establish whether excess was actually exceeded",
        "Ombudsman Reasoning": "Excess argument fails — total claim value includes tree investigation and crack monitoring (>£600) plus £1,900 repair costs, likely exceeding £2,500; pre-existing exclusion fails — AIL's inspection did not state when subsidence started; homebuyer's survey found no active movement; no expert evidence of pre-inception subsidence; burden on AIL not discharged; AIL obliged to continue cover after accepted claim; failure to renew created obligation to pay increased premium costs consumer incurred",
        "Workflow Insight": "Insurer relying on a pre-existing damage exclusion must obtain and produce expert evidence that the current subsidence began before policy inception — it is not sufficient to note that prior subsidence occurred years earlier and was repaired; excess must be calculated against total claim value including all investigation, monitoring, tree works and repair costs — it is not limited to the repair schedule alone; insurer that wrongly declines a claim and then fails to renew the policy creates a compounded liability: the claim itself plus the increased premium costs the consumer paid elsewhere and the loss of subsidence cover for those renewal years",
        "AI Rule Candidate": "IF pre_existing_damage_exclusion_relied_upon THEN insurer_must_produce_expert_evidence_that_subsidence_existed_before_inception; IF homebuyer_survey_found_no_active_movement THEN strong_evidence_against_pre_inception_subsidence; excess_must_be_calculated_against_total_claim_value_including_investigation_monitoring_and_tree_costs; IF insurer_wrongly_declines_claim AND fails_to_renew THEN liable_for_claim_plus_increased_premium_costs_at_alternative_insurers; consumer_professional_costs_after_wrongful_decline_may_be_offset_against_excess_or_refunded",
        "Source PDF": "DRN-3387540.pdf",
    },
    {
        "Case ID": "SUBS-014",
        "FOS Decision ID": "DRN-3427348",
        "Insurer Name": "Amtrust Europe Limited",
        "FOS Decision Date": "9 May 2022",
        "Claim Type": "Commercial landlord buildings insurance — cracking reported June 2021; Amtrust concluded longstanding movement from street view images (2012+) and 2008 pre-purchase survey; backdated removal of subsidence cover to 2017 (claiming misrepresentation at renewal) and declined claim on grounds movement is settlement not subsidence",
        "Movement Cause": "Longstanding differential settlement caused by different foundation depths between front and rear of building (front founded higher due to cellar construction; different depths cause very gradual differential downward movement); trees, drain leakage and clay shrinkage all ruled out by Amtrust specialists",
        "Property Type": "Residential let property (landlord/commercial policy — Insurance Act 2015 applies)",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": "Amtrust: (1) longstanding cracking and movement not disclosed at 2017 renewal when asked whether property had experienced stepped/diagonal cracking, subsidence, ground heave, landslip, movement or underpinning — answered 'No'; would not have provided subsidence cover if known; backdated subsidence cover removal to 2017. (2) Movement is settlement (differential foundation depths) not subsidence — settlement excluded under policy",
        "Evidence Dispute": "Mrs G: unaware of cracks as property was tenanted; disputes cracks were present; Amtrust's agents behaved badly during inspection. Amtrust: street view images from 2012 clearly show cracks; 2008 pre-purchase survey noted longstanding movement; landlord has responsibility to inspect and maintain; underwriting guide supports removing subsidence cover. FOS provisional: Mrs G misrepresented (should have answered 'Yes'); Insurance Act 2015 applies; BUT Amtrust's underwriting guide only removes subsidence cover for previous/existing confirmed subsidence — not for settlement or general movement; movement confirmed as settlement not subsidence; Amtrust did not demonstrate qualifying breach justifying subsidence cover removal under IA 2015. FOS final: Amtrust's oral argument that underwriters interpret guide more broadly than its text is insufficient — guide must demonstrate what Amtrust would have done; qualifying breach for subsidence cover removal not established; backdating unfair; claim decline for settlement is correct",
        "Outcome Category": "Upheld in Part",
        "Outcome": "Amtrust to reinstate subsidence cover backdated to 2017; claim decline upheld — damage caused by settlement which is excluded under the policy",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Insurance Act 2015 (commercial landlord policy) — insurer must demonstrate a qualifying breach to apply a remedy; insurer's underwriting guide must specifically support the action taken; if underwriting guide only removes subsidence cover for confirmed previous/existing subsidence, insurer cannot extend that to settlement or general movement/cracking — it must show qualifying breach for the specific scenario; settlement and subsidence are distinct and separately addressed perils — insurer must apply the correct exclusion for the correct peril; misrepresentation on a broad question (including cracking/movement) does not automatically create a qualifying breach justifying removal of subsidence cover if underwriting criteria do not specifically address that scenario",
        "Missing Evidence": "Amtrust underwriting guide provisions specifically addressing what the insurer would do if informed of cracking/movement that is not confirmed subsidence (guide only addressed confirmed subsidence explicitly); whether Mrs G received any communication about cracks from tenants",
        "Ombudsman Reasoning": "Mrs G misrepresented — street view images from 2012 show longstanding cracks; should have answered 'Yes' at 2017 renewal; Insurance Act 2015 applies; qualifying breach requires showing what Amtrust would have done — underwriting guide only covers confirmed previous/existing subsidence; damage confirmed as settlement (different foundation depths) not subsidence; policy already excludes settlement as standard, so removing subsidence cover specifically would not have changed outcome for settlement damage; Amtrust hasn't demonstrated qualifying breach to support backdated removal of subsidence cover; backdating unfair; claim decline for settlement is fair and correct",
        "Workflow Insight": "Commercial landlord insurers relying on Insurance Act 2015 to backdate removal of subsidence cover must produce underwriting guide provisions specifically covering the disclosed risk scenario — a guide that only addresses confirmed subsidence does not justify removal of cover where the disclosed issue is settlement, cracking or movement; settlement and subsidence must be distinguished at claim stage — settlement is typically excluded as standard while subsidence requires specific cover; where a broad disclosure question captures general cracking/movement that turns out to be settlement (not subsidence), and the policy already excludes settlement, removing subsidence cover in addition is disproportionate and unsupported by standard underwriting criteria; landlord-policy misrepresentation cases under IA 2015 require specific demonstration that the remedy taken matches what the underwriting guide actually says",
        "AI Rule Candidate": "IF insurance_act_2015_applies AND insurer_claims_qualifying_breach THEN underwriting_guide_must_specifically_support_action_taken; IF underwriting_guide_addresses_only_confirmed_subsidence THEN it_does_not_support_removal_of_cover_for_settlement_or_general_cracking; settlement_and_subsidence_are_distinct_perils_requiring_separate_exclusion_analysis; IF movement_confirmed_as_settlement_not_subsidence THEN settlement_exclusion_applies_and_claim_decline_is_fair; misrepresentation_on_broad_movement_question_does_not_automatically_justify_removal_of_subsidence_cover",
        "Source PDF": "DRN-3427348.pdf",
    },
    {
        "Case ID": "SUBS-015",
        "FOS Decision ID": "DRN-3581769",
        "Insurer Name": "Fairmead Insurance Limited",
        "FOS Decision Date": "3 Aug 2022",
        "Claim Type": "Home insurance — subsidence claim accepted (porch sinking, January 2020); dispute about whether drain repairs caused by the subsidence should be covered under one subsidence claim or a separate accidental damage claim with a separate £200 excess; also unresolved subsidence structural repairs and inadequate communication throughout",
        "Movement Cause": "Ground subsidence causing porch to sink; drains within area of influence damaged by subsidence movement; separate drain defects outside area of influence unrelated to subsidence; patio damage attributed to settlement not subsidence",
        "Property Type": "Residential home",
        "Dispute Type": "Handling / Reinstatement Dispute",
        "Coverage Decision": "Accepted — Disputed Settlement",
        "Rejection Reason": "Fairmead: drain repairs within the area of subsidence influence covered under subsidence claim; drain repairs outside the area of influence not connected to subsidence — separate accidental damage claim with separate £200 excess; patio damage = settlement (not covered); faulty drainage design and missing lintels = defects in design/workmanship (not covered)",
        "Evidence Dispute": "Mr and Mrs C: all drain repairs and all damage on three sides should be covered under one subsidence claim. Fairmead: drains within influence area covered under subsidence; drains outside influence area dealt with separately. FOS: internal notes, loss adjuster breakdown and two receipts for £200 excess confirm a separate £200 excess was charged for drain repairs that were actually connected to and caused by the subsidence — despite Fairmead's assertion to the contrary; drain damage caused by and contributing to subsidence must be included under the one subsidence claim with the one subsidence excess; drain repairs outside the area of influence and genuinely unrelated to subsidence can legitimately be a separate claim; consumer-funded private repairs without expert evidence linking them to subsidence: not reimbursable; faulty design/workmanship and missing lintels not covered; further site investigation needed to confirm drain repairs are lasting and effective",
        "Outcome Category": "Upheld in Part",
        "Outcome": "Carry out further investigation to confirm drain repairs under subsidence claim are lasting and effective; complete lasting and effective repair if not; investigate slabs near kitchen for subsidence damage and relay if confirmed; refund £200 accidental damage excess paid April 2020 with 8% simple interest from date of payment to settlement; update all internal and external records to reflect one subsidence claim; pay £500 compensation",
        "Compensation Awarded (£)": 500,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Drain repairs caused by subsidence and contributing to it must be included under the subsidence claim with a single subsidence excess — insurer cannot split them into a separate accidental damage claim; drain repairs outside the area of subsidence influence that are genuinely unrelated to the movement may be treated as a separate claim; excess applies only once per subsidence event — insurer charging a second excess for subsidence-related drain repairs is incorrect; insurer must maintain accurate internal and external records — if records show two separate claims for the same subsidence event, they must be corrected; insurer must be given first opportunity to investigate and rectify its own repairs before consumer can demand independent survey",
        "Missing Evidence": "Expert evidence demonstrating the privately-funded drain repairs (outside area of influence) were caused by or contributed to the subsidence; confirmation of whether slabs outside kitchen moved due to subsidence",
        "Ombudsman Reasoning": "Internal notes, loss adjuster breakdown and consumer receipts all confirm a separate £200 accidental damage excess was charged for drain repairs actually related to subsidence — this contradicts Fairmead's assertion that all was handled under one claim; the fair principle is that subsidence-related drain repairs belong under one subsidence claim with one excess; Fairmead's stated policy on this point (apply one subsidence claim for subsidence-related drains) was not followed in practice; further drain investigation needed as repairs may be defective; slabs to be reviewed at same time; communication was poor throughout; £500 compensation appropriate",
        "Workflow Insight": "When drain repairs are split between 'subsidence-related' and 'unrelated', the insurer must correctly categorise each part and charge only one excess for all subsidence-related elements — charging a separate accidental damage excess on drain repairs within the area of influence is incorrect and will be reversed; internal claim records must accurately reflect what was covered and how — inconsistency between notes and practice creates both a rectification obligation and a compensation liability; insurer must be given first opportunity to inspect and rectify its own repair work before consumer-requested independent surveys are authorised",
        "AI Rule Candidate": "drain_repairs_caused_by_subsidence_and_contributing_to_it_must_be_included_under_subsidence_claim_with_single_excess; IF insurer_charges_separate_accidental_damage_excess_for_drain_repairs_within_subsidence_area_of_influence THEN excess_must_be_refunded; drain_repairs_outside_area_of_influence_and_genuinely_unrelated_may_be_separate_claim; IF consumer_alleges_insurer_repairs_defective THEN insurer_must_reinvestigate_before_independent_survey_is_directed; insurer_records_must_accurately_reflect_one_subsidence_claim_for_one_subsidence_event",
        "Source PDF": "DRN-3581769.pdf",
    },
    {
        "Case ID": "SUBS-016",
        "FOS Decision ID": "DRN-3682901",
        "Insurer Name": "Society of Lloyd's",
        "FOS Decision Date": "7 Oct 2022",
        "Claim Type": "Home insurance — subsidence accepted (November 2020); insurer identified defective drains as the cause but refused to fund drain repairs under the subsidence claim, citing the accidental damage section's 'sudden' requirement; told consumers to arrange and fund drain repairs themselves before the subsidence claim could progress",
        "Movement Cause": "Ground movement caused by defective drains (leaking or washing away supporting ground); drain defects occurred gradually over time — not sudden; drain repair is required before the structural subsidence repairs can be carried out to prevent recurrence",
        "Property Type": "Residential home",
        "Dispute Type": "Handling / Reinstatement Dispute",
        "Coverage Decision": "Accepted — Disputed Settlement",
        "Rejection Reason": "Society of Lloyd's: drain defects fall under the accidental damage section of the policy; the accidental damage section requires a sudden event and drain defects occurred gradually; accidental damage section therefore does not apply; consumers must fund and arrange drain repairs themselves before insurer will progress subsidence repairs",
        "Evidence Dispute": "Society of Lloyd's: separate accidental damage section covers drains; 'sudden' requirement not met. FOS: insurer's obligation to carry out a lasting and effective repair when a valid subsidence claim is accepted extends beyond the literal scope of the policy — it includes repairing the underlying cause of the subsidence to prevent recurrence; drain repair is required for a lasting and effective repair; whether the drains are separately covered under the accidental damage section is irrelevant — they must be repaired under the subsidence section to fulfil the lasting and effective repair obligation; analogous to underpinning (which no insurer disputes); one excess only; Mr and Mrs F are in their 90s and paid from savings — aggravating factor for compensation",
        "Outcome Category": "Upheld",
        "Outcome": "Reimburse Mr and Mrs F for cost of drain repairs on receipt of evidence of costs; pay 8% simple interest per annum from date of payment to settlement; charge only one excess for the entire claim; deal with all repairs as one subsidence claim; pay £300 total compensation (£150 additional)",
        "Compensation Awarded (£)": 300,
        "Is Core Case": "Yes",
        "Key Policy Clause": "Insurer's duty to carry out a lasting and effective repair under an accepted subsidence claim extends to repairing the underlying cause of the subsidence (e.g. defective drains) even if that cause is not explicitly covered under the policy's accidental damage section; where drain repair is required before structural repairs can be durable, the insurer must fund it as part of the subsidence claim regardless of whether the 'sudden' requirement in the accidental damage section is met; this obligation is analogous to underpinning — it is a necessary precondition for an effective repair; only one excess applies to one subsidence event; consumer vulnerability (advanced age, using savings) is a relevant factor in setting compensation level",
        "Missing Evidence": "N/A — the factual situation was not in dispute; Lloyd's accepted the drains needed repair but disputed who should pay",
        "Ombudsman Reasoning": "Society of Lloyd's accepted subsidence and obligation to repair; lasting and effective repair requires removal of the underlying cause; drain repair is required to prevent subsidence recurring — without it, structural repairs will fail; whether drains are covered under the accidental damage section is irrelevant; insurer's obligation under the subsidence section overrides the separate section analysis; analogous to underpinning, which all insurers acknowledge as their obligation; one excess for one event; Mr and Mrs F are in their 90s and paid from savings — this aggravates the D&I impact and justifies £300 total compensation",
        "Workflow Insight": "When a subsidence investigation identifies defective drains as the root cause, the insurer must fund those drain repairs as part of the subsidence claim even if the policy's accidental damage section does not technically cover the drain defect — the lasting and effective repair obligation overrides the section-by-section analysis; an insurer cannot use an accidental damage 'sudden event' requirement to avoid funding drain repairs that are the identified cause of an accepted subsidence claim; one excess applies per subsidence event — a second excess for the drain element is incorrect; consumer age and financial vulnerability are relevant to D&I compensation quantum",
        "AI Rule Candidate": "IF defective_drains_are_identified_cause_of_accepted_subsidence THEN insurer_must_fund_drain_repairs_under_subsidence_claim; drain_repair_required_for_lasting_and_effective_repair_must_be_funded_regardless_of_accidental_damage_section_limitations; IF drain_defects_occurred_gradually_and_accidental_damage_section_requires_sudden_event THEN insurer_cannot_use_that_to_avoid_drain_repair_obligation_under_subsidence_section; one_excess_applies_per_subsidence_event; consumer_vulnerability_and_age_are_relevant_factors_in_compensation_quantum",
        "Source PDF": "DRN-3682901.pdf",
    },
    {
        "Case ID": "SUBS-017",
        "FOS Decision ID": "DRN-3929594",
        "Insurer Name": "AXA Insurance UK Plc",
        "FOS Decision Date": "24 Jul 2023",
        "Claim Type": "Home insurance — complaint about premium and subsidence excess increases at April 2021 and April 2022 renewals following accepted subsidence claim (November 2020 claim accepted; second claim recorded March 2022); AXA failed to respond to complaint",
        "Movement Cause": "Subsidence (cause not specified — first claim accepted and cash settled November 2022; second claim recorded March 2022 and under investigation)",
        "Property Type": "Residential home",
        "Dispute Type": "Handling / Reinstatement Dispute",
        "Coverage Decision": "Accepted — Disputed Settlement",
        "Rejection Reason": "AXA: subsidence excess increases based on underwriting criteria and claim reserve (amounts spent and expected to be spent on claim) at each renewal; premium increases correctly calculated per AXA's underwriting model; second claim at 2022 renewal further increased risk",
        "Evidence Dispute": "Mr and Mrs I: increases unfair. AXA: detailed premium calculations submitted; excess based on claim reserve; individual underwriter would not have calculated differently. FOS investigator: premium increases correct; excess increase not justified at the level set. AXA: provided further underwriting comments challenging investigator. FOS ombudsman: ABI guidance requires renewal terms to be reasonable not merely technically correct per underwriting formula; premium increases are both correct and reasonable given subsidence claim history and second claim at 2022 renewal; excess — AXA set it at £10,000 based on an overstated reserve; applying AXA's own criteria to actual accurate claim values, £5,000 was the appropriate excess at both 2021 and 2022 renewals; both parties accepted £5,000 figure",
        "Outcome Category": "Upheld",
        "Outcome": "AXA to reduce the subsidence excess at the 2021 and 2022 renewals to £5,000",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "ABI guidance on continuation of cover after subsidence claim requires insurer to offer renewal on reasonable terms — reasonableness is the standard, not merely mathematical accuracy of the underwriting formula; premium increases post-subsidence claim are justified if correctly calculated per the insurer's underwriting criteria AND not unreasonable; subsidence excess post-claim must reflect the actual claim value (reserve), not an inflated or over-estimated amount — FOS will review against both the insurer's own criteria and the accuracy of the reserve figures used; a second subsidence claim provides additional justification for a premium increase at that renewal",
        "Missing Evidence": "Accurate claim reserve at each renewal (AXA used overstated reserve to calculate excess — FOS applied actual settled values to determine £5,000 appropriate)",
        "Ombudsman Reasoning": "Premium increases correctly calculated per AXA criteria and reasonable given subsidence history and second claim; excess incorrectly set at £10,000 — based on overstated reserve figures; applying AXA's own underwriting criteria to accurate claim values, £5,000 is the fair excess at both renewals; ABI standard requires reasonable terms, not just formula compliance; second claim at 2022 renewal legitimately contributed to premium increase; AXA did not respond to the original complaint — a separate failing acknowledged",
        "Workflow Insight": "Post-subsidence claim excess setting must use accurate claim reserve values — using an overstated reserve to set a higher excess will be corrected by FOS to the level the insurer's own criteria would produce when applied to actual values; premium increases must satisfy the ABI reasonableness standard, not just the insurer's internal pricing model; when a second subsidence claim is recorded, it provides a stronger basis for renewal excess and premium adjustments; insurer failure to respond to a complaint on time is a separate service failing but does not override the substantive analysis of whether the pricing was fair",
        "AI Rule Candidate": "subsidence_excess_at_renewal_must_be_calculated_on_accurate_claim_reserve_not_overstated_estimate; IF insurer_excess_based_on_overstated_reserve THEN FOS_will_reduce_excess_to_level_justified_by_actual_claim_values_under_insurers_own_criteria; ABI_continuation_guidance_requires_renewal_terms_to_be_reasonable_not_merely_formula_correct; second_subsidence_claim_provides_additional_basis_for_excess_and_premium_increase_at_that_renewal",
        "Source PDF": "DRN-3929594.pdf",
    },
    {
        "Case ID": "SUBS-018",
        "FOS Decision ID": "DRN-4190935",
        "Insurer Name": "Society of Lloyd's",
        "FOS Decision Date": "15 Jan 2024",
        "Claim Type": "Home insurance — Society of Lloyd's underwriter (U) withdrew from the home insurance market; Mr M had an open subsidence claim from March 2020; underwriter provided a 3-month extension then a 12-month renewal with £10,000 excess, but did not renew at expiry (August 2021); broker could not find replacement cover with subsidence; ABI guidance on continuation of cover not followed",
        "Movement Cause": "Subsidence (cause not specified — claim accepted in 2020; investigation ongoing when cover was removed)",
        "Property Type": "Residential home",
        "Dispute Type": "Handling / Reinstatement Dispute",
        "Coverage Decision": "Accepted",
        "Rejection Reason": "Society of Lloyd's / underwriter U: commercial decision to withdraw from providing the policy type Mr M held; decision applied to all policyholders with that policy type, not just Mr M; U not obliged to provide insurance; broker was given adequate time to find alternative cover",
        "Evidence Dispute": "Mr M: cannot find subsidence cover elsewhere due to open claim; disadvantaged by market withdrawal. SoL: ABI guidance does not require insurers withdrawing from the market to make exceptions or find alternative cover for policyholders; withdrawal applied to all such policyholders not just Mr M; extension provided sufficient notice. FOS: ABI guidance applies regardless of ABI membership (industry good practice); guidance's block transfer provisions demonstrate its spirit — insurer ceasing to provide cover should arrange for policyholders not to be disadvantaged; if U had not withdrawn from the market, Mr M would have had continuing cover; U took no steps beyond a short extension to arrange alternative cover for Mr M; other insurers in similar positions have made exceptions, arranged policy transfers or subsidised specialist cover; inability to arrange alternative cover not actively demonstrated; not sufficient that continued cover is merely unwanted — must show it is impossible",
        "Outcome Category": "Upheld",
        "Outcome": "Society of Lloyd's to arrange for an underwriter (U or otherwise) to provide home insurance including subsidence cover to Mr M on reasonable terms, backdated to when his policy with U ended; premium and terms must not be set in a way that effectively prevents access to the ongoing subsidence cover",
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": "ABI guidance on continuation of cover following a subsidence claim applies to non-ABI members as a statement of industry good practice; the guidance applies even when an insurer withdraws from the home insurance market — withdrawal from the market is not listed as an exception; insurer withdrawing from the market must take reasonable steps to ensure the spirit of the guidance is achieved for policyholders with open/prior subsidence claims — this includes arranging alternative cover, subsidising specialist cover, or making an exception; the block transfer provisions of the guidance demonstrate the intent that stopping insurers must protect continuing-cover obligations; insurer preference not to provide cover is different from inability to provide cover — inability must be actively demonstrated; cover arranged for continuing-cover compliance must not be priced or structured so as to effectively prevent access",
        "Missing Evidence": "Evidence that Society of Lloyd's or underwriter U actively attempted and failed to arrange alternative subsidence cover for Mr M (never pursued)",
        "Ombudsman Reasoning": "ABI guidance applies regardless of ABI membership; guidance applies to market withdrawal scenario — not listed as exception; block transfer provisions show guidance intends stopping insurers to arrange ongoing cover; if U had remained in market, Mr M would have had continuing cover — market withdrawal directly disadvantaged him; U provided short extension only and then left Mr M without subsidence cover; other insurers in similar situations found solutions (exceptions, transfers, subsidised specialist cover); SoL asserted alternative impossible but never actively explored it; not persuaded that arranging cover is impossible; SoL must arrange cover — it does not need to set up the policy itself, only to arrange an underwriter to provide it; terms must not effectively prevent access",
        "Workflow Insight": "An insurer withdrawing from the home insurance market must take proactive steps to protect policyholders with open or prior subsidence claims — the ABI continuation guidance is not disapplied by market exit; acceptable steps include making an exception for affected policyholders, arranging a block transfer to another insurer, or subsidising the difference between the existing premium and a specialist provider; merely telling the broker to find alternative cover and providing a short extension is insufficient; insurer must actively explore whether alternative arrangements are possible before asserting they cannot be made; any alternative cover arranged must be genuinely accessible — priced or structured so as to effectively prevent access is non-compliance",
        "AI Rule Candidate": "ABI_continuation_guidance_applies_to_non_ABI_members_as_industry_good_practice; ABI_continuation_guidance_applies_even_when_insurer_withdraws_from_home_insurance_market; IF insurer_withdraws_from_market AND policyholder_has_open_subsidence_claim THEN insurer_must_take_reasonable_steps_to_arrange_alternative_cover; acceptable_steps_include_exception_for_affected_customers_OR_block_transfer_OR_subsidy_for_specialist_cover; insurer_preference_not_to_provide_cover_is_not_the_same_as_inability_inability_must_be_demonstrated; alternative_cover_must_not_be_priced_so_as_to_effectively_prevent_access",
        "Source PDF": "DRN-4190935.pdf",
    },
    {
        "Case ID": "SUBS-019",
        "FOS Decision ID": "DRN-4813489",
        "Insurer Name": "Kennett Insurance Brokers Limited",
        "FOS Decision Date": "23 Apr 2025",
        "Claim Type": "Buildings insurance — broker (Kennett) failed to declare Mr and Mrs D's 2003 subsidence claim when arranging new buildings insurance; insurer discovered discrepancy when a new subsidence claim was made in summer 2022 and avoided the policy; complaint against broker for omitting the prior claim from the application",
        "Movement Cause": "2003 subsidence (cause not specified — handled by prior insurer); 2022 subsidence (investigation not completed — insurer avoided policy before investigation)",
        "Property Type": "Residential home",
        "Dispute Type": "Broker Conduct Dispute",
        "Coverage Decision": "Not Applicable",
        "Rejection Reason": "Kennett: standard wording requires consumers to verify accuracy of insurance documentation; Mr and Mrs D never flagged the omission over many years; Mr D was advised not to move away from the insurer who handled the 2003 claim but chose to do so; cannot establish exactly how the subsidence claim was removed from records",
        "Evidence Dispute": "Mr and Mrs D: trusted Kennett as longstanding broker (40+ years); Kennett knew about 2003 subsidence claim; should not need to audit broker's work annually. Kennett: cannot be certain how or when claim was removed from records; consumers had opportunity to review paperwork. FOS: likely Kennett's error caused the omission before the 2014 policy was taken out; demands and needs statements from 2017–2021 all referenced a subsidence excess, reasonably leading consumers to believe prior claim was properly declared; some policy schedules should have alerted them but not all required a claims history section; Kennett accepts it should have checked more closely when arranging policies; distress caused by Kennett's failure to maintain accurate records over a long relationship is clear; £150 compensation appropriate; insurer's avoidance handled in a separate complaint",
        "Outcome Category": "Upheld",
        "Outcome": "Kennett Insurance Brokers Limited to pay Mr and Mrs D £150 compensation for distress and inconvenience",
        "Compensation Awarded (£)": 150,
        "Is Core Case": "No — Broker Dispute",
        "Key Policy Clause": "Broker has a professional duty to accurately record and carry forward material claims history — including prior subsidence claims — when arranging insurance for longstanding clients; a broker who has knowledge of a prior subsidence claim and later arranges insurance that omits it is likely responsible for that omission; longstanding broker relationship and consumer trust do not transfer the broker's duty of care to the consumer; inclusion of a subsidence excess in renewal documents issued over multiple years can reasonably mislead a consumer into believing the prior claim is properly declared",
        "Missing Evidence": "Exact process and timing by which the 2003 subsidence claim was removed from Mr and Mrs D's insurance records (passage of time makes certainty impossible)",
        "Ombudsman Reasoning": "2003 subsidence not in dispute; Kennett knew about it; Kennett accepts omission occurred; likely Kennett's error before 2014; demands and needs statements consistently referenced subsidence excess, giving consumers reasonable comfort that the prior claim was properly declared; primary responsibility rests with Kennett; consumers could have been more vigilant but Kennett bears the main duty; insurer avoidance separately handled; £150 compensation appropriate for distress from a 40-year broker relationship ending in avoidance",
        "Workflow Insight": "Brokers placing insurance for clients with known prior subsidence claims must verify that each new policy application and subsequent renewals accurately reflect the subsidence claim history — repeated year-on-year omission cannot be blamed on the consumer; where a broker's own documents (demands and needs statements) reference a subsidence excess, consumers will reasonably but incorrectly believe the prior claim is properly declared — this reduces consumer contributory responsibility; insurer avoidance arising from broker omission creates both a broker compensation liability and a separate insurer-facing complaint",
        "AI Rule Candidate": "broker_must_accurately_record_prior_subsidence_claim_in_each_new_policy_application_it_arranges; IF broker_knew_of_prior_subsidence_claim AND arranged_policy_that_omitted_it THEN broker_is_likely_responsible_for_omission; IF renewal_documents_reference_subsidence_excess THEN consumer_reasonably_believed_prior_claim_was_declared; broker_omission_creating_insurer_avoidance_creates_broker_liability_for_D_and_I_and_separate_insurer_complaint",
        "Source PDF": "DRN-4813489.pdf",
    },
    {
        "Case ID": "SUBS-020",
        "FOS Decision ID": "DRN-4883553",
        "Insurer Name": "AXA Insurance UK Plc",
        "FOS Decision Date": "30 Sep 2024",
        "Claim Type": "Home insurance — AXA removed subsidence, landslip and heave (SLH) cover entirely at the April 2023 renewal following two accepted subsidence claims (2020 and 2022); AXA cited risk appetite and claim history; SLH cover also absent from April 2024 renewal",
        "Movement Cause": "Subsidence (two accepted claims; specific causes not detailed in this decision — decision concerns cover removal, not cause investigation)",
        "Property Type": "Residential home",
        "Dispute Type": "Handling / Reinstatement Dispute",
        "Coverage Decision": "Accepted",
        "Rejection Reason": "AXA: two subsidence claims increased risk beyond risk appetite; scenario not addressed by underwriting criteria so individual underwriter decided to remove SLH cover; entitled to refuse cover outside risk appetite; offered to 'consider' SLH in future",
        "Evidence Dispute": "Mrs I: purchased specialist non-standard insurance specifically covering subsidence; unfair to remove cover because she made claims against it; left in vulnerable position with subsidence history and no cover. AXA: two claims increased risk; happy to 'consider' SLH; underwriting criteria did not address multiple claims scenario. FOS investigator: AXA (ABI member) failed to act in line with ABI guidance; 'consider' is insufficient; recommended reinstatement backdated. FOS ombudsman: ABI guidance requires continuation of subsidence cover after a claim; guidance acknowledges problem cases (repeated subsidence, non-disclosure) but directs cover be provided wherever possible even in those cases; two subsidence claims is the scenario the guidance is specifically designed to address; it cannot be a compelling reason to remove cover; AXA must reinstate SLH for April 2023 and April 2024 renewals; vague offer to 'consider' is not an offer of cover",
        "Outcome Category": "Upheld",
        "Outcome": "AXA to amend April 2023 and April 2024 policies to include subsidence, landslip and heave cover; pay £250 compensation",
        "Compensation Awarded (£)": 250,
        "Is Core Case": "Yes",
        "Key Policy Clause": "ABI guidance on continuation of cover — ABI member insurer that has dealt with a subsidence claim must normally continue to offer subsidence cover on the property; guidance explicitly acknowledges problem cases including repeated subsidence and requires cover to be provided wherever possible even in those cases; two subsidence claims cannot be a compelling reason to remove cover — they are the exact scenario the guidance is designed to remedy; insurer's 'risk appetite' is not a compelling reason that overrides ABI guidance obligations; offer to 'consider' providing SLH cover is not an offer of cover and does not satisfy the continuation obligation; only a risk fundamentally outside usual underwriting criteria (analogous to fraud) could justify departure from the continuation obligation",
        "Missing Evidence": "AXA's written underwriting criteria for scenarios involving two or more subsidence claims (decision was made by individual underwriter in absence of applicable criteria)",
        "Ombudsman Reasoning": "AXA is an ABI member — guidance binding; guidance requires continuation of SLH cover after subsidence claim; guidance recognises problem cases (repeated subsidence) and requires cover wherever possible even then; two claims is not a compelling reason to remove cover — it is the problem the guidance exists to address; AXA produced no compelling reason beyond two claims; AXA's 'consider' offer is insufficient — positive reinstatement required; April 2023 and 2024 policies must be amended; £250 compensation for unnecessary distress (related case DRN-3929594 had already reduced the excess to £5,000 at those same renewals)",
        "Workflow Insight": "ABI member insurers cannot rely on 'risk appetite' or 'claim history' as justification for removing SLH cover after multiple subsidence claims — the ABI continuation guidance specifically addresses repeated subsidence as a problem case and requires cover wherever possible; absence of underwriting criteria for a multiple-claim scenario means a commercial decision by an individual underwriter is not supported by written criteria and will be reviewed and overturned by FOS; the continuation obligation is not satisfied by an offer to 'consider' providing cover in future — it requires positive provision of cover; where a prior FOS decision has already corrected the excess for the same policy period, removal of SLH cover from the same renewal is an additional and independent breach of the ABI guidance",
        "AI Rule Candidate": "ABI_member_insurer_must_continue_SLH_cover_after_subsidence_claim; two_subsidence_claims_are_not_a_compelling_reason_to_remove_cover_they_are_the_scenario_the_guidance_addresses; IF insurer_removes_SLH_cover_citing_risk_appetite_or_claims_history THEN removal_is_unfair_unless_compelling_reason_beyond_subsidence_itself_exists; offer_to_consider_cover_is_not_an_offer_of_cover; IF individual_underwriter_decision_not_supported_by_written_underwriting_criteria THEN decision_is_more_easily_overturned_by_FOS; departure_from_ABI_continuation_obligation_requires_compelling_reason_analogous_to_fraud_not_merely_commercial_unwillingness",
        "Source PDF": "DRN-4883553.pdf",
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
