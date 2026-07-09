"""
Standard append script for all Unoccupied Property batches (schema v1, 21
columns — same schema as EOW v2 / Storm v1 / Flood v1 / Subsidence v1 /
Theft v1; column 6 = "Unoccupied Period / Circumstance").

Active — reuse this script for every future batch: replace NEW_CASES below
with the next batch's cases and run again. Appends only — never modifies
existing rows in knowledge/case-databases/Unoccupied_Property_Case_Database.xlsx.

Current contents: Batch 2 (UNOC-011 to UNOC-020).
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
    "Unoccupied Period / Circumstance",
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
    "Dispute Type": [
        "Coverage Dispute",
        "Endorsement / Exclusion Challenge",
        "Broker Conduct Dispute",
        "Claim Recording / Administrative Dispute",
        "Claim Quantum Dispute",
        "Third-Party / Liability Dispute",
        "Causation Dispute",
    ],
    "Coverage Decision": [
        "Declined — Full",
        "Declined — Partial",
        "Accepted",
        "Accepted — With Deductions",
        "Not Applicable",
    ],
    "Outcome Category": [
        "Upheld",
        "Upheld in Part",
        "Not Upheld",
        "Not Applicable",
    ],
    "Is Core Case": [
        "Yes",
        "No — Administrative",
        "No — Broker Dispute",
        "No — Quantum Only",
        "No — Time-Barred",
    ],
}

ROW_FILL_ODD  = PatternFill(start_color="F9EDF2", end_color="F9EDF2", fill_type="solid")
ROW_FILL_EVEN = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
ROW_FONT      = Font(name="Calibri", size=10)
ROW_BORDER    = Border(
    left=Side(style="thin", color="DDDDDD"),
    right=Side(style="thin", color="DDDDDD"),
    top=Side(style="thin", color="DDDDDD"),
    bottom=Side(style="thin", color="DDDDDD"),
)

NEW_CASES = [
    {
        "Case ID": "UNOC-011",
        "FOS Decision ID": "DRN-1623377",
        "Insurer Name": "Society of Lloyd's",
        "FOS Decision Date": "6 May 2020",
        "Claim Type": (
            "Unoccupied property insurance (undergoing refurbishment) — escape of water "
            "from attic pipe declined under Structural Works Clause endorsement; insurer "
            "maintained the clause excluded EOW/Trace and Access cover from the moment the "
            "endorsement was added, regardless of whether structural work had actually begun"
        ),
        "Unoccupied Period / Circumstance": (
            "Property unoccupied from November 2018 policy inception pending refurbishment; "
            "structural work disclosed at inception but had not yet started when the "
            "February 2019 escape of water occurred"
        ),
        "Property Type": "Residential unoccupied property (pre-refurbishment)",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": (
            "Structural Works Clause endorsement excluded Escape of Water and Trace and "
            "Access cover 'until all the structural work, as disclosed to us, has been "
            "fully completed'; insurer applied this from the date the endorsement was "
            "added, not from the date work actually started"
        ),
        "Evidence Dispute": (
            "Mr and Mrs S (via loss adjuster): endorsement should not bite until structural "
            "work had actually commenced; no work had started when the leak occurred. "
            "Lloyd's: endorsement applies 'at all times' from inception once disclosed, "
            "irrespective of whether work had started. FOS: Schedule of Cover endorsements "
            "section read as a whole is clear — EOW cover suspended 'at all times' and "
            "'until all structural work has been completed'; the loss adjuster's on-site "
            "report did not confirm heating clause compliance and had not completed its "
            "assessment; insurer retained final decision-making authority over the loss "
            "adjuster's report"
        ),
        "Outcome Category": "Not Upheld",
        "Outcome": (
            "Complaint not upheld — Lloyd's correctly declined the escape of water claim "
            "under the terms of the Structural Works Clause endorsement"
        ),
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": (
            "A Structural Works Clause endorsement that suspends Escape of Water and Trace "
            "and Access cover 'until all the structural work, as disclosed to us, has been "
            "fully completed' operates from the point the endorsement is added to the "
            "policy — it is not conditional on the disclosed works having actually started; "
            "reading the endorsements section as a whole (including the 'at all times' "
            "wording), a policyholder's honesty about future intended works does not "
            "prevent the exclusion applying immediately upon disclosure"
        ),
        "Missing Evidence": (
            "A completed loss adjuster assessment confirming compliance with the Heating "
            "Clause (Unoccupied Property) — the adjuster's report did not resolve this "
            "point and instead awaited a plumber's report on causation, but this was "
            "ultimately immaterial once the Structural Works Clause was found to apply"
        ),
        "Ombudsman Reasoning": (
            "Schedule of Cover stated endorsements apply 'at all times'; EOW/Trace and "
            "Access excluded until structural work fully completed, reinstated thereafter "
            "subject to the Heating Clause; no work had started but exclusion still applied "
            "per plain wording; loss adjuster's incomplete assessment was only one input "
            "into the insurer's overall decision, which properly considered the "
            "endorsement; claim correctly declined"
        ),
        "Workflow Insight": (
            "When a Structural Works Clause or similar endorsement is added upon disclosure "
            "of intended future works, claims handlers must apply the exclusion from the "
            "endorsement's effective date, not from the date works commence — "
            "policyholders should be warned at point of disclosure that cover for related "
            "perils (EOW, Trace and Access) is suspended immediately, not only once work "
            "begins; an incomplete or inconclusive loss adjuster report does not prevent "
            "the insurer relying on a clear policy exclusion when it makes its final "
            "coverage decision"
        ),
        "AI Rule Candidate": (
            "IF structural_works_clause_endorsement_added_on_disclosure_of_future_works "
            "AND policy_states_endorsements_apply_at_all_times "
            "THEN eow_and_trace_and_access_cover_excluded_from_endorsement_effective_date_regardless_of_whether_works_have_started; "
            "loss_adjuster_incomplete_assessment_does_not_override_clear_policy_exclusion_relied_upon_by_insurer_in_final_decision"
        ),
        "Source PDF": "DRN-1623377.pdf",
    },
    {
        "Case ID": "UNOC-012",
        "FOS Decision ID": "DRN-1719844",
        "Insurer Name": "Internet Insurance Services UK Ltd",
        "FOS Decision Date": "7 Jan 2021",
        "Claim Type": (
            "Broker mis-sale complaint — advised sale of buildings insurance policy "
            "(branded 'Residential Let Property Scheme') for a property that would be "
            "unoccupied during renovation; consumer alleged policy name implied it covered "
            "let properties and was therefore unsuitable/misleading, and that reduced cover "
            "(fire, lightning, explosion, aircraft only) for unoccupied periods wasn't "
            "adequately explained"
        ),
        "Unoccupied Period / Circumstance": (
            "Property unoccupied during a renovation/extension project from April 2010; "
            "policy provided reduced unoccupied-property cover for the works period, "
            "reverting to full cover once the consumer moved back in; policy itself lapsed "
            "in 2011"
        ),
        "Property Type": "Residential property under renovation (unoccupied during works)",
        "Dispute Type": "Broker Conduct Dispute",
        "Coverage Decision": "Not Applicable",
        "Rejection Reason": (
            "Not applicable to a coverage decline — dispute concerned suitability of the "
            "policy sold in 2010, discovered when a 2019 roof claim (caused by builder "
            "workmanship, not an insured peril) was declined as outside the "
            "fire/lightning/explosion/aircraft cover for unoccupied periods"
        ),
        "Evidence Dispute": (
            "Ms S: policy name ('Residential Let Property Scheme') implied it was for let "
            "properties, which caused confusion and made her doubt suitability; wanted "
            "broader cover. IIS: call recordings show advisor clearly explained the reduced "
            "level of cover available during the unoccupied works period ('all this policy "
            "will cover... is fire, lightning, earthquake, explosion and property owners' "
            "liability... won't cover theft, malicious damage, escape of water or storm "
            "damage'); policy provided exactly the cover Ms S asked for (buildings cover "
            "while unoccupied and under renovation). FOS: advisor's explanation of scope of "
            "cover was clear; policy name may have been confusing but the substance of "
            "cover (not the name) is what matters for suitability; no requirement the "
            "property be let for the policy to apply; no evidence of detriment from the "
            "policy name in the ten years since sale"
        ),
        "Outcome Category": "Not Upheld",
        "Outcome": (
            "Complaint not upheld — the policy sold was suitable for Ms S's stated need "
            "(buildings cover for an unoccupied property undergoing renovation); the level "
            "of cover was clearly explained on the calls even though the policy's product "
            "name may have caused some confusion"
        ),
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Broker Dispute",
        "Key Policy Clause": (
            "For an advised sale of unoccupied-property buildings cover, the broker must "
            "clearly explain the reduced scope of cover applicable during the unoccupied "
            "period (e.g. fire/lightning/explosion/aircraft only, excluding theft, "
            "malicious damage, escape of water and storm) — recorded confirmation of this "
            "explanation on the sales call is strong evidence of suitability; a confusing "
            "or unconventional product name (e.g. one associated with let properties) does "
            "not by itself render an unoccupied-property policy unsuitable where the "
            "substantive cover matches the consumer's stated needs and there is no "
            "requirement that the property actually be let"
        ),
        "Missing Evidence": (
            "Evidence that Ms S specifically requested cover for the builder's workmanship "
            "or latent defects — absent; the extensively recorded pre-sale calls did not "
            "include this requirement"
        ),
        "Ombudsman Reasoning": (
            "Two recorded calls showed IIS's advisor explained the unoccupied-period cover "
            "was limited to fire, lightning, earthquake, explosion and property owners' "
            "liability, excluding EOW/theft/storm; Ms S provided rebuild estimates and "
            "confirmed she'd return once works completed; policy matched her stated needs; "
            "the 'Residential Let Property Scheme' name was potentially confusing but its "
            "substance, not its name, governs suitability; no evidence of financial "
            "detriment from the name over ten years; not mis-sold"
        ),
        "Workflow Insight": (
            "Brokers selling unoccupied-property cover under a product name that could "
            "imply a different purpose (e.g. a let-property scheme) should proactively "
            "clarify the name's origin at the point of sale, even where the substantive "
            "cover matches the consumer's needs, to avoid downstream confusion; "
            "suitability assessments for unoccupied-property add-ons should focus on "
            "whether the explained scope of cover (not the product branding) matches what "
            "the consumer told the broker they needed"
        ),
        "AI Rule Candidate": (
            "IF broker_clearly_explains_reduced_unoccupied_period_cover_on_recorded_call "
            "AND cover_matches_consumers_stated_need "
            "THEN advised_sale_of_unoccupied_property_policy_is_suitable_regardless_of_potentially_confusing_product_name; "
            "policy_name_alone_does_not_establish_mis_sale_absent_evidence_the_substantive_cover_did_not_match_consumer_requirements"
        ),
        "Source PDF": "DRN-1719844.pdf",
    },
    {
        "Case ID": "UNOC-013",
        "FOS Decision ID": "DRN1874552",
        "Insurer Name": "Ageas Insurance Limited",
        "FOS Decision Date": "19 Feb 2015",
        "Claim Type": (
            "Commercial property owners insurance (two self-contained units) — escape of "
            "water from multiple burst pipes over Christmas period in a unit undergoing "
            "renovation after tenants vacated; heating not switched on due to works; "
            "insurer declined claim citing unoccupied-property exclusion and drain-down "
            "condition breach"
        ),
        "Unoccupied Period / Circumstance": (
            "One of two self-contained units vacated by tenants and unoccupied for more "
            "than 30 consecutive days while undergoing building works; director visited at "
            "least daily during the Christmas holiday period when works paused; other unit "
            "remained tenanted and unaffected"
        ),
        "Property Type": "Commercial property (two self-contained units; one unoccupied during renovation)",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": (
            "Policy excluded escape of water/freezing damage to premises unoccupied, "
            "untenanted, or not in active use for 30+ consecutive days; company was also "
            "in breach of the condition requiring water systems to be drained in "
            "unoccupied properties"
        ),
        "Evidence Dispute": (
            "B (via advisers): had been in occupation during building works and a director "
            "visited at least daily over Christmas, so the unit wasn't 'unoccupied'; "
            "exclusion/condition not market standard and not properly drawn to attention "
            "at inception; loss adjusters gave inconsistent reasons for declining. Ageas: "
            "unit was untenanted and had not been in active use in connection with the "
            "property-owner business for 30+ days; frequent visits for building-site "
            "purposes are not equivalent to occupation. FOS: policy's 'unoccupied' "
            "definition is disjunctive — meeting any one of unoccupied/untenanted/not in "
            "active use for 30+ days triggers the exclusion; the unit was untenanted "
            "regardless of the other tenanted unit; daily inspection visits to what was "
            "'essentially a building site' do not amount to occupation; exclusion not "
            "unusual and broker (not insurer) was responsible for explaining standard "
            "policy terms"
        ),
        "Outcome Category": "Not Upheld",
        "Outcome": (
            "Complaint not upheld — Ageas was entitled to decline the claim; the unit was "
            "unoccupied under the policy's disjunctive definition and the exclusion applied"
        ),
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": (
            "A disjunctive definition of 'unoccupied' (unoccupied OR untenanted OR not in "
            "active use for the relevant business for 30+ consecutive days) is satisfied "
            "if any single limb is met — a property does not need to be literally empty of "
            "all activity to be classed as unoccupied where it has been untenanted for the "
            "qualifying period, even if other visits (e.g. by a director carrying out or "
            "overseeing building works) occur frequently; frequent inspection visits to "
            "premises that are 'essentially a building site' do not constitute occupation; "
            "where only one part of a multi-unit building is untenanted, the unoccupied "
            "exclusion can apply to that part alone regardless of continued occupation "
            "elsewhere in the building; standard, non-unusual unoccupied-property "
            "exclusions and conditions sold via a broker do not need to be separately "
            "highlighted by the insurer"
        ),
        "Missing Evidence": (
            "Evidence that active occupation-level use (rather than periodic inspection "
            "during renovation) continued in the untenanted unit — absent; B accepted the "
            "unit had been vacated by tenants more than 30 days before the loss"
        ),
        "Ombudsman Reasoning": (
            "Definition applies if 'any…part of the building' is unoccupied/untenanted/not "
            "in active use for 30+ days; the vacated unit was untenanted for the qualifying "
            "period; building-works activity and daily director visits over Christmas did "
            "not equate to normal commercial occupation; loss adjusters' inconsistent "
            "explanations did not prejudice B given it had expert advisers; exclusion and "
            "drain-down condition were standard and did not require specific insurer "
            "highlighting when sold through a broker; claim properly declined"
        ),
        "Workflow Insight": (
            "Multi-unit or multi-tenant commercial property policies should be assessed "
            "unit-by-unit for unoccupancy purposes — continued occupation of one unit does "
            "not preserve cover for a vacated unit under a disjunctive 'unoccupied' "
            "definition; claims handlers should distinguish between genuine occupation-"
            "level activity and periodic building-site inspection visits when assessing "
            "whether the unoccupancy threshold has been met; loss adjusters should give "
            "consistent, clearly-referenced reasons for declining a claim even where the "
            "policyholder has expert representation"
        ),
        "AI Rule Candidate": (
            "IF policy_defines_unoccupied_disjunctively_as_unoccupied_OR_untenanted_OR_not_in_active_use_for_30_plus_days "
            "THEN any_single_limb_being_met_triggers_the_unoccupied_exclusion; "
            "IF only_activity_at_property_is_periodic_inspection_during_building_works "
            "THEN this_does_not_constitute_occupation_or_active_use; "
            "IF one_unit_of_multi_unit_property_is_untenanted_for_qualifying_period "
            "THEN unoccupied_exclusion_applies_to_that_unit_regardless_of_occupation_status_of_other_units"
        ),
        "Source PDF": "DRN1874552.pdf",
    },
    {
        "Case ID": "UNOC-014",
        "FOS Decision ID": "DRN2281195",
        "Insurer Name": "Society of Lloyd's",
        "FOS Decision Date": "18 Jul 2019",
        "Claim Type": (
            "Unoccupied property insurance (converted from a tenanted policy at renewal "
            "after Lloyd's discovered extended unoccupancy) — arson fire started by "
            "intruders/vandals; insurer declined citing breach of conditions requiring "
            "removal of refuse/waste materials and notification of illegal occupiers"
        ),
        "Unoccupied Period / Circumstance": (
            "Property unoccupied since November 2016 (previously let); renewed under a "
            "tenanted policy in May 2017 without Lloyd's being told of the unoccupancy; "
            "Lloyd's treated the claim as if an unoccupied-property policy had applied, "
            "since it would not have offered a tenanted policy had it known"
        ),
        "Property Type": "Residential unoccupied property (previously let; undergoing refurbishment at time of fire)",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": (
            "Breach of unoccupied-property policy condition requiring all refuse and "
            "waste materials to be removed from the interior; photographs showed a large "
            "quantity of burnt rubbish, discarded timber, waste products, plastic buckets "
            "and ripped-out kitchen carcasses; fire brigade report found the fire started "
            "with paper/cardboard, which contributed to the fire's spread"
        ),
        "Evidence Dispute": (
            "Mr G: unaware of the unoccupied-property conditions since he hadn't been told "
            "the property was being treated as unoccupied; items were retained for "
            "refurbishment (kitchen units to be fitted) and books/comics intended for "
            "charity, not waste; fire was started deliberately by third parties regardless "
            "of the property's contents. Lloyd's: would only have offered an unoccupied-"
            "property policy on renewal had it known of the unoccupancy, and did not offer "
            "a tenanted policy for empty properties; photographic evidence showed "
            "materials consistent with waste, not solely refurbishment items. FOS "
            "(provisional and final): fair to treat the claim as if under an unoccupied-"
            "property policy since the alternative was no cover at all; volume of debris "
            "exceeded what would reasonably be expected from refurbishment alone; breach "
            "of the waste-removal condition was materially connected to the fire loss "
            "since waste paper was the ignition source and fed the fire's spread; only one "
            "breached condition is needed to justify declining a claim"
        ),
        "Outcome Category": "Not Upheld",
        "Outcome": (
            "Complaint not upheld — Lloyd's fairly treated the claim under the terms "
            "applicable to an unoccupied property and was entitled to decline it for "
            "breach of the waste-removal condition, which was materially connected to the "
            "fire loss"
        ),
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": (
            "Where an insurer discovers after a loss that a property was unoccupied when "
            "a tenanted (not unoccupied) policy was in force, and it would only have "
            "offered unoccupied-property cover (with its own conditions) had it known, it "
            "is reasonable to assess the claim as though the unoccupied-property policy "
            "terms applied — the policyholder's unawareness of those specific conditions "
            "does not excuse compliance, since notifying the insurer of the true occupancy "
            "status was the policyholder's own responsibility; a condition requiring "
            "removal of refuse and waste materials from an unoccupied property is a "
            "risk-control measure specifically aimed at reducing arson risk, and a breach "
            "of it will bar a claim for fire damage where the ignition source and its "
            "spread are materially connected to the retained waste — this differs from "
            "breach of an unrelated condition (e.g. a burglar alarm requirement) in a "
            "flood claim, which would not be materially connected"
        ),
        "Missing Evidence": (
            "Clear itemised inventory distinguishing genuine refurbishment materials from "
            "discarded waste at the time of the fire — absent; the photographic evidence "
            "was relied upon instead and showed volumes inconsistent with refurbishment "
            "alone"
        ),
        "Ombudsman Reasoning": (
            "Property unoccupied since November 2016; renewed as tenanted May 2017 without "
            "notification; Lloyd's would not have offered a tenanted policy for an empty "
            "property, so treating the claim under unoccupied-property terms was "
            "reasonable; photographs showed 'heavily loaded' debris, rubbish and loose "
            "combustible materials disproportionate to refurbishment needs; fire brigade "
            "confirmed paper/cardboard ignition; breach of the waste-removal condition "
            "materially connected to the fire since it fed the blaze; only one condition "
            "breach needed to support declining the claim; insured's lack of connection to "
            "the arson itself is irrelevant where a separate condition breach applies"
        ),
        "Workflow Insight": (
            "When an insurer discovers post-loss that a tenanted policy was incorrectly in "
            "force over what was actually an unoccupied property, it is fair to assess the "
            "claim against the unoccupied-property terms it would have offered, rather "
            "than refusing cover outright — but this means the policyholder is held to "
            "conditions (e.g. waste removal, illegal-occupier notification) they were "
            "never told about, so materiality of any breach to the specific loss becomes "
            "the key fairness test; claims handlers must show a breach of a risk-control "
            "condition (like waste removal) is causally connected to the peril claimed for "
            "(fire) before relying on it to decline"
        ),
        "AI Rule Candidate": (
            "IF insurer_discovers_post_loss_that_tenanted_policy_was_in_force_over_actually_unoccupied_property "
            "AND insurer_would_only_have_offered_unoccupied_property_terms_if_notified "
            "THEN claim_may_fairly_be_assessed_against_unoccupied_property_policy_conditions_even_though_policyholder_was_unaware_of_them; "
            "IF waste_removal_condition_breached_AND_waste_materials_were_ignition_source_or_fuel_for_the_fire "
            "THEN breach_is_materially_connected_to_fire_loss_and_insurer_may_decline_claim; "
            "only_one_materially_connected_condition_breach_is_required_to_justify_declining_a_claim"
        ),
        "Source PDF": "DRN2281195.pdf",
    },
    {
        "Case ID": "UNOC-015",
        "FOS Decision ID": "DRN2912517",
        "Insurer Name": "UK Insurance Limited",
        "FOS Decision Date": "7 Aug 2017",
        "Claim Type": (
            "Commercial property insurance — escape of sewage from a tenanted shop/"
            "basement discovered by local authority after buildup over two-three months; "
            "insurer declined citing breach of unoccupied-property notification, "
            "inspection and reasonable-precautions conditions after determining the "
            "property had in fact been unoccupied"
        ),
        "Unoccupied Period / Circumstance": (
            "Property unoccupied for an extended period (tenant hadn't paid rent since "
            "February 2014; some evidence of a year or more of unoccupancy) but the "
            "policyholder (landlord) says it only became aware of the unoccupancy when "
            "the local authority contacted it in January 2016; tenant later explained he "
            "had been abroad in India due to a family medical emergency"
        ),
        "Property Type": "Commercial let property (shop with basement)",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": (
            "UKI concluded the property had been unoccupied when the sewage escape "
            "occurred and said P had breached obligations to notify of a change in "
            "occupation, to inspect the unoccupied property every 14 days, and to take "
            "reasonable precautions to prevent damage"
        ),
        "Evidence Dispute": (
            "P: unaware the property was unoccupied until notified by the local authority "
            "in January 2016; had been accommodating the tenant's personal difficulties "
            "over unpaid rent since February 2014 and only instructed solicitors in June "
            "2015 over the rent arrears, not because it knew the property was empty; "
            "director lived some distance away and had no reason to suspect external "
            "signs of vacancy. UKI: tenant hadn't paid rent since February 2014; "
            "neighbours reported the shop had been closed for a year or more; instructing "
            "solicitors over unpaid rent in June 2015 should have prompted notification of "
            "a change in risk. FOS: policy's notification obligation and 'unaffected by "
            "act/omission of a tenant' proviso both require actual awareness of a change "
            "in occupation before the obligation to notify arises; no evidence P knew the "
            "property was unoccupied before January 2016; instructing solicitors over "
            "unpaid rent is not equivalent to knowing a property is unoccupied; without "
            "awareness, the 14-day inspection and general reasonable-precautions "
            "obligations could not reasonably have been triggered either"
        ),
        "Outcome Category": "Upheld",
        "Outcome": (
            "UKI required to accept the claim, reimburse any invoices already paid by P "
            "directly with 8% simple interest from the date of payment to settlement, and "
            "continue providing cover for the repair work"
        ),
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": (
            "A policy condition requiring notification of a change in occupation 'as soon "
            "as they become aware' only bites once the insured has actual knowledge of the "
            "change — a policy provision stating cover 'will not be prejudiced by any act "
            "or omission unknown to or beyond the control of the Insured on the part of a "
            "tenant' reinforces that unawareness of a tenant's own conduct (including "
            "abandoning the property) does not constitute a breach; obligations to inspect "
            "unoccupied premises at least every 14 days and take reasonable precautions "
            "for unoccupied properties cannot fairly be relied upon by an insurer where "
            "the policyholder did not know, and had no reasonable means of knowing, that "
            "the property had become unoccupied; non-payment of rent alone, or "
            "instructing solicitors to recover arrears, does not equate to knowledge that "
            "a property is unoccupied"
        ),
        "Missing Evidence": (
            "Evidence that P had actual knowledge of the unoccupancy before the local "
            "authority's January 2016 contact — absent; UKI relied only on inferences from "
            "non-payment of rent and neighbour reports, not on anything shown to have "
            "reached P"
        ),
        "Ombudsman Reasoning": (
            "Tenant hadn't paid rent since Feb 2014, and neighbours suggested the shop had "
            "been shut for a year or more, but P's director lived some distance away and "
            "had no reason to suspect the property (as opposed to the tenant's rent "
            "account) was actually vacant; solicitor instruction in June 2015 concerned "
            "unpaid rent, not unoccupancy; policy's notification duty and the 'unknown to "
            "or beyond control of the Insured' proviso both turn on actual awareness; "
            "without awareness, the 14-day inspection and reasonable-precautions "
            "conditions specific to vacant/disused locations could not reasonably apply "
            "either; UKI could not rely on any of the three clauses to decline; claim "
            "should be accepted with interest on any invoices already paid"
        ),
        "Workflow Insight": (
            "When investigating suspected non-disclosure of unoccupancy, insurers must "
            "establish that the policyholder had actual knowledge of the change in "
            "occupation, not merely that objective indicators (unpaid rent, neighbour "
            "reports of closure) existed which the insurer itself later uncovered — a "
            "landlord's tolerance of a struggling tenant's rent arrears, or instructing "
            "solicitors to recover unpaid rent, is not on its own evidence of knowledge "
            "that the property has become physically unoccupied; distance between a "
            "landlord's residence and the let property is a relevant factor in assessing "
            "whether the landlord could reasonably have known of unoccupancy"
        ),
        "AI Rule Candidate": (
            "IF policy_notification_duty_is_conditioned_on_insured_becoming_aware_of_change_in_occupation "
            "AND insurer_cannot_show_actual_knowledge_by_policyholder_of_unoccupancy "
            "THEN insurer_cannot_decline_claim_for_breach_of_notification_condition; "
            "IF policyholder_lacked_actual_knowledge_of_unoccupancy "
            "THEN inspection_and_reasonable_precautions_conditions_specific_to_unoccupied_properties_cannot_be_relied_upon_either; "
            "non_payment_of_rent_or_solicitor_instruction_over_arrears_does_not_by_itself_establish_landlord_knowledge_that_property_is_unoccupied"
        ),
        "Source PDF": "DRN2912517.pdf",
    },
    {
        "Case ID": "UNOC-016",
        "FOS Decision ID": "DRN-2973838",
        "Insurer Name": "AmTrust Europe Limited",
        "FOS Decision Date": "22 Oct 2021",
        "Claim Type": (
            "Commercial buildings insurance (let flats) — theft of pipework and fixed "
            "appliances during the void period following an accepted fire damage claim, "
            "before reinstatement works began; insurer declined the theft claim under the "
            "policy's unoccupancy exclusion, regardless of the reason for the unoccupancy"
        ),
        "Unoccupied Period / Circumstance": (
            "Property unoccupied following fire damage that forced tenants to move out; "
            "unoccupancy exceeded the policy's 60-day threshold (defined as empty/"
            "disused/unoccupied/unfurnished/untenanted for 60+ days) while awaiting claim "
            "assessment and reinstatement; theft occurred during this period, before "
            "repairs started"
        ),
        "Property Type": "Residential let property, multiple self-contained flats",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": (
            "Policy excluded theft where the buildings had become 'Unoccupied' (empty/"
            "disused/unfurnished/untenanted for 60+ days) under Endorsement 011b and the "
            "Change of Definition — Unoccupied clause; theft occurred after the 60-day "
            "threshold had been passed while the property remained vacant pending "
            "fire-claim reinstatement"
        ),
        "Evidence Dispute": (
            "Mr R: unoccupancy only arose because of an insured peril (fire) which he "
            "could not control, and was compounded by claim-handling delays; AmTrust "
            "should not rely on the exclusion in these circumstances. AmTrust: "
            "unoccupancy exclusion applies regardless of the underlying cause of the "
            "unoccupancy; offered £100 compensation for acknowledged claim-handling "
            "delays but maintained the theft exclusion applied. FOS: policy wording is "
            "clear and contains no exception for unoccupancy caused by an insured peril; "
            "even absent any avoidable delay, the loss adjuster estimated repairs would "
            "take around four months — well beyond the 60-day threshold — so the property "
            "would have remained unoccupied past the threshold regardless of AmTrust's "
            "handling"
        ),
        "Outcome Category": "Not Upheld",
        "Outcome": (
            "Complaint not upheld — AmTrust correctly declined the theft claim under the "
            "unoccupancy exclusion; £100 already offered for acknowledged claim-handling "
            "delays stands but the theft claim itself is not payable, since even without "
            "those delays the 60-day threshold would still have been exceeded given the "
            "scale of the fire repairs"
        ),
        "Compensation Awarded (£)": 100,
        "Is Core Case": "Yes",
        "Key Policy Clause": (
            "An unoccupancy exclusion tied to a fixed time threshold (e.g. 60 days empty/"
            "disused/unfurnished/untenanted) applies regardless of the reason the property "
            "became unoccupied — including where unoccupancy results from an earlier "
            "insured peril (such as fire) that is entirely outside the policyholder's "
            "control; there is no implied exception for unoccupancy caused by a covered "
            "event unless the policy expressly provides for one; where an insurer has "
            "caused some avoidable claim-handling delay, this will not assist the "
            "policyholder in disputing an unoccupancy exclusion unless the delay can be "
            "shown to have caused the property to cross the relevant unoccupancy "
            "threshold that it would not otherwise have crossed"
        ),
        "Missing Evidence": (
            "Evidence that, absent AmTrust's acknowledged handling delays, reinstatement "
            "could have been completed and the property reoccupied within 60 days of the "
            "fire — absent; the loss adjuster's own repair-time estimate (approximately "
            "four months) showed the threshold would have been exceeded regardless"
        ),
        "Ombudsman Reasoning": (
            "Policy and Endorsement 011b unambiguously excluded theft cover once the "
            "property met the 60-day unoccupancy definition, with no carve-out based on "
            "cause; property became unoccupied due to fire damage, an event outside "
            "either party's control at the time of the theft; AmTrust's acknowledged "
            "handling delays did not causally extend the unoccupancy period beyond what "
            "the fire damage itself would have caused, since repairs were independently "
            "estimated at around four months; £100 already offered for the delays was "
            "proportionate and not increased since no material impact from the delay was "
            "shown"
        ),
        "Workflow Insight": (
            "Claims handlers should confirm whether a fixed-threshold unoccupancy "
            "exclusion contains any carve-out for unoccupancy arising from a prior "
            "insured peril before declining a subsequent claim on that ground — where no "
            "such carve-out exists, the exclusion applies regardless of blamelessness; "
            "where a policyholder alleges that insurer-caused delay extended the "
            "unoccupancy period past the relevant threshold, handlers should obtain and "
            "rely on independent repair-time estimates to assess whether the threshold "
            "would have been crossed in any event"
        ),
        "AI Rule Candidate": (
            "IF unoccupancy_exclusion_is_defined_by_a_fixed_time_threshold_with_no_cause_based_carve_out "
            "THEN exclusion_applies_regardless_of_whether_unoccupancy_arose_from_a_separate_insured_peril; "
            "IF policyholder_alleges_insurer_delay_caused_unoccupancy_threshold_to_be_exceeded "
            "AND independent_repair_estimate_shows_threshold_would_have_been_exceeded_regardless_of_delay "
            "THEN delay_argument_does_not_defeat_the_unoccupancy_exclusion"
        ),
        "Source PDF": "DRN-2973838.pdf",
    },
    {
        "Case ID": "UNOC-017",
        "FOS Decision ID": "DRN-3053282",
        "Insurer Name": "Aviva Insurance Limited",
        "FOS Decision Date": "29 Dec 2021",
        "Claim Type": (
            "Commercial Property Owners insurance — mid-term policy adjustment dispute "
            "(not a declined claim); at renewal the insured disclosed the commercial "
            "rental property was unoccupied, so Aviva applied an Unoccupied Premises "
            "Cover Restriction (fire/lightning/explosion/earthquake only); when the "
            "property became occupied again four months later, Aviva reinstated full "
            "cover but also raised the escape of water and storm/flood excesses to "
            "£2,500 (later reduced EOW excess back to £500 but kept storm/flood at "
            "£2,500 due to an open prior storm claim); policyholder complained the "
            "mid-term excess increase was unfair"
        ),
        "Unoccupied Period / Circumstance": (
            "Commercial rental property unoccupied at the point of annual renewal; "
            "became occupied again approximately four months into the policy year, at "
            "which point full cover (subject to revised excesses) was reinstated"
        ),
        "Property Type": "Commercial rental property",
        "Dispute Type": "Claim Recording / Administrative Dispute",
        "Coverage Decision": "Not Applicable",
        "Rejection Reason": (
            "Not applicable — no claim was declined; the dispute concerned whether a "
            "mid-term increase to policy excesses, applied when cover was upgraded from "
            "the Unoccupied Premises Cover Restriction back to full standard cover, was "
            "permitted and fair under the policy's Alteration of Risk clause"
        ),
        "Evidence Dispute": (
            "Ms B: unoccupancy endorsement could simply have remained dormant and not "
            "applied once the property became occupied again, without any excess "
            "increase; combined with a separate £2,500 loss-of-rent excess, a storm claim "
            "could cost her £5,000 in excesses; if the policy term permitted this, the "
            "term itself was unfair. Aviva: at renewal, cover was restricted to fire/"
            "lightning/explosion/earthquake only because the property was unoccupied, so "
            "EOW and storm/flood risk (and any related claim history) hadn't been priced "
            "in; when occupancy resumed, it reassessed the risk for those newly-"
            "reinstated perils, taking into account an open prior storm claim, and set "
            "higher excesses accordingly; reduced the EOW excess back to standard once it "
            "recognised the open claim was for storm damage only. FOS: the policy's "
            "Alteration of Risk clause expressly permits new terms and conditions "
            "(including premium/excess changes) once notified of a change in risk; moving "
            "from unoccupied to occupied is itself a notifiable change in risk that "
            "engaged this clause; increasing the storm/flood excess for a fresh, "
            "previously-unpriced risk in light of an open storm claim was a legitimate "
            "commercial underwriting decision, not something FOS can second-guess on "
            "quantum"
        ),
        "Outcome Category": "Not Upheld",
        "Outcome": (
            "Complaint not upheld — Aviva was entitled under the policy's Alteration of "
            "Risk clause to reassess and adjust excesses for perils newly reinstated when "
            "the property became occupied again, and did so fairly by relying on the open "
            "storm claim history"
        ),
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Administrative",
        "Key Policy Clause": (
            "An Alteration of Risk clause permitting new terms and conditions (including "
            "premium and excess changes) upon notification of a change in risk extends to "
            "a policyholder's notification that a previously unoccupied insured property "
            "has become occupied again — because perils excluded during the Unoccupied "
            "Premises Cover Restriction period (such as escape of water and storm/flood) "
            "were never priced into the restricted-cover premium, the insurer is entitled "
            "to reassess and reprice those specific perils, including by reference to "
            "claims history, when reinstating full cover; FOS will not adjudicate on the "
            "appropriateness of the excess amount itself, only on whether the policy "
            "permitted the adjustment and whether the process was fair"
        ),
        "Missing Evidence": (
            "Not applicable — the material facts (unoccupancy at renewal, restricted "
            "cover applied, subsequent notification of reoccupation, and an open prior "
            "storm claim) were undisputed; the complaint turned on interpretation of the "
            "Alteration of Risk clause rather than any factual gap"
        ),
        "Ombudsman Reasoning": (
            "Policy renewed on the basis of unoccupancy with a restricted level of cover; "
            "Alteration of Risk clause allows new terms/conditions upon notification of a "
            "change in risk; occupancy resuming is a notifiable change; EOW and storm/"
            "flood cover (and any pricing for prior claims) was not factored in while the "
            "restriction applied, so it was reasonable for Aviva to reassess these "
            "specific perils on reinstatement; EOW excess was corrected once Aviva "
            "recognised the open claim was storm-only; storm/flood excess increase for a "
            "peril newly brought back into scope, informed by claims history, was fair "
            "and within Aviva's commercial discretion"
        ),
        "Workflow Insight": (
            "Underwriters restricting cover for an unoccupied property should record which "
            "specific perils are excluded from pricing during the restriction so that, on "
            "reinstatement following reoccupation, only the newly-reinstated perils are "
            "subject to fresh risk assessment and excess adjustment — clear documentation "
            "of this distinction (as Aviva ultimately demonstrated by correcting the EOW "
            "excess once the storm-only nature of the open claim was recognised) helps "
            "resolve excess disputes without recourse to FOS"
        ),
        "AI Rule Candidate": (
            "IF policyholder_notifies_insurer_that_previously_unoccupied_property_has_become_occupied_again "
            "AND alteration_of_risk_clause_permits_new_terms_on_notified_change_in_risk "
            "THEN insurer_may_reassess_and_reprice_perils_that_were_excluded_or_unpriced_during_the_unoccupied_restricted_cover_period; "
            "excess_or_premium_adjustments_for_perils_newly_reinstated_on_reoccupation_informed_by_open_claims_history_are_not_unfair_merely_because_they_increase_cost_to_policyholder"
        ),
        "Source PDF": "DRN-3053282.pdf",
    },
    {
        "Case ID": "UNOC-018",
        "FOS Decision ID": "DRN-3101941",
        "Insurer Name": "AXA Insurance UK Plc",
        "FOS Decision Date": "9 Nov 2021",
        "Claim Type": (
            "Residential leasehold flat insurance — accidental damage claim for a "
            "gable-end wall partially collapsing after the dormer roof of a long-"
            "unoccupied, derelict neighbouring property caved in; council issued a "
            "Dangerous Building Notice requiring residents to leave; AXA declined the "
            "claim, attributing the wall collapse to gradual water ingress/general wear "
            "rather than a one-off event, even though two of the other three flat owners "
            "in the same building had the same damage accepted by their own insurers"
        ),
        "Unoccupied Period / Circumstance": (
            "Not the insured's own property — the causative unoccupied property was a "
            "neighbouring building that had stood empty and derelict for some years "
            "(subject to two prior fires in 2013 and 2016, and long-standing local media "
            "coverage of its disrepair) before its dormer roof caved in and caused the "
            "shared gable-end wall to partially collapse"
        ),
        "Property Type": (
            "Residential leasehold flat (one of four units in a shared building; damage "
            "caused by an adjoining unoccupied, derelict property)"
        ),
        "Dispute Type": "Causation Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": (
            "AXA considered the wall damage was not a one-off accidental event but "
            "resulted from water ingress accumulating over a number of years, softening "
            "the brickwork and ultimately causing the wall to collapse — treated as "
            "gradual/maintenance-related damage excluded under the policy"
        ),
        "Evidence Dispute": (
            "Mr D and Mr R: two other flat owners in the same building, sharing the same "
            "damage to the same gable-end wall from the same event (the neighbouring "
            "dormer roof collapse), had their claims accepted by their own insurers, one "
            "of whom used the same loss adjuster as AXA; inconsistent for AXA alone to "
            "decline. AXA: maintained the damage was gradual water-ingress deterioration "
            "excluded under the policy, and that other insurers' decisions did not bind "
            "it. FOS (provisional, confirmed in final decision): the two properties were "
            "immediately adjoining, so deterioration of the shared wall would likely not "
            "have been visible until the dormer roof partially collapsed; no evidence of "
            "any internal signs of damage beforehand; a June 2018 survey found the "
            "chimney in fair condition; media reports confirmed the developer of the "
            "derelict property withdrew after the collapse 'damaged neighbouring homes'; "
            "all four flat owners share responsibility for, and are claiming for, the "
            "identical damage to the identical wall arising from the identical event, so "
            "it must be assessed identically; AXA's continued refusal despite two other "
            "insurers (including one using AXA's own loss adjuster) accepting the same "
            "damage as accidental was inconsistent and unfair"
        ),
        "Outcome Category": "Upheld",
        "Outcome": (
            "AXA required to accept and pay Mr D and Mr R's claim in line with the "
            "settlements already made to the other flat owners for the same damage, "
            "under the remaining terms of the policy"
        ),
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": (
            "Where multiple co-owners in the same building each hold separate buildings "
            "insurance policies and all are claiming for the identical damage arising "
            "from the identical single event (here, the collapse of a neighbouring "
            "long-unoccupied, derelict property's roof onto a shared structural wall), an "
            "insurer that treats the damage as gradual/excluded while other insurers "
            "covering the same wall and the same event have accepted the claim as "
            "accidental is acting inconsistently; good industry practice and treating "
            "customers fairly requires consistent treatment of identical damage from an "
            "identical shared event across the different insurers involved, absent a "
            "clear, insurer-specific reason for departing from that consistency; the fact "
            "that a wall was hidden from view (and therefore its deteriorating condition "
            "could not have been observed or maintained by the policyholder) prior to a "
            "sudden triggering event supports treating the resulting damage as a one-off "
            "accidental loss rather than gradual deterioration"
        ),
        "Missing Evidence": (
            "AXA did not produce independent engineering or surveying evidence specific "
            "to Mr D and Mr R's portion of the wall to justify treating their claim "
            "differently from the two accepted claims for the same wall and the same "
            "event; it relied solely on its own opinion that gradual water ingress caused "
            "the wall's softening"
        ),
        "Ombudsman Reasoning": (
            "Site adjacency meant the wall's deteriorating condition would not have been "
            "visible until the neighbouring dormer roof partially collapsed; no internal "
            "signs of damage were shown to exist beforehand; 2018 survey found the "
            "chimney in fair condition; media reports confirmed a developer withdrew from "
            "the derelict site citing damage to neighbouring homes from the collapse; two "
            "of the four flat owners, including one sharing AXA's own loss adjuster, had "
            "the identical damage accepted as accidental by their insurers; AXA "
            "maintained its gradual-damage position without new evidence after the "
            "provisional decision; fair and reasonable outcome required AXA to treat Mr D "
            "and Mr R consistently with the other owners and pay the claim"
        ),
        "Workflow Insight": (
            "Where damage originates from a long-unoccupied, derelict neighbouring "
            "property affecting multiple co-owned or shared-structure units, claims "
            "handlers should proactively check how other insurers covering the same "
            "building or wall have treated equivalent claims for the same triggering "
            "event before relying on a gradual-damage or wear-and-tear exclusion — "
            "inconsistency across insurers assessing identical damage from an identical "
            "cause is a strong indicator of an unfair declinature; adjacency and lack of "
            "visibility of a shared structural element prior to a sudden triggering event "
            "(e.g. a neighbouring roof collapse) should be weighed heavily in favour of "
            "treating resulting damage as a one-off accidental event"
        ),
        "AI Rule Candidate": (
            "IF multiple_co_owners_of_a_shared_structural_element_hold_separate_policies "
            "AND all_are_claiming_for_identical_damage_from_an_identical_single_triggering_event "
            "THEN insurer_declining_on_gradual_damage_grounds_while_other_insurers_covering_the_same_element_and_event_have_accepted_the_claim_is_likely_unfair_absent_specific_contrary_evidence; "
            "IF damage_originates_from_a_long_unoccupied_derelict_neighbouring_property_and_shared_wall_condition_was_not_visible_or_inspectable_before_a_sudden_triggering_event "
            "THEN resulting_damage_should_generally_be_treated_as_a_one_off_accidental_loss_rather_than_gradual_deterioration"
        ),
        "Source PDF": "DRN-3101941.pdf",
    },
    {
        "Case ID": "UNOC-019",
        "FOS Decision ID": "DRN-3113837",
        "Insurer Name": "AXA Insurance UK Plc",
        "FOS Decision Date": "16 Dec 2021",
        "Claim Type": (
            "Home insurance — policyholder notified insurer she was moving out (property "
            "becoming unoccupied); call handler misadvised her that staying at the "
            "property for a couple of nights every 60 days would reset the occupancy "
            "requirement and preserve cover; insurer's underwriting department later "
            "overturned this advice and cancelled the policy because the property was no "
            "longer her permanent home"
        ),
        "Unoccupied Period / Circumstance": (
            "Property became unoccupied as Mrs Y moved to live elsewhere on a permanent "
            "basis; she visited for a 10-day period at one point, incurring travel costs, "
            "believing (based on the misadvice) that periodic overnight stays preserved "
            "her cover"
        ),
        "Property Type": "Residential property (owner relocating; no longer permanent home)",
        "Dispute Type": "Claim Recording / Administrative Dispute",
        "Coverage Decision": "Not Applicable",
        "Rejection Reason": (
            "Not applicable to a coverage decline — AXA's underwriting department "
            "determined the policy did not meet its underwriting rules once it "
            "understood Mrs Y did not intend to return to the property on a permanent "
            "basis, and issued a cancellation notice rather than declining any specific "
            "claim"
        ),
        "Evidence Dispute": (
            "Mrs Y: was told by a call handler that staying at the property for a couple "
            "of nights every 60 days would preserve cover, then was shocked to later "
            "receive a cancellation letter; sought reimbursement of travel costs incurred "
            "visiting the property in reliance on the (incorrect) advice. AXA: accepted "
            "the original advice was a mistake — the property no longer being Mrs Y's "
            "permanent home did not meet its underwriting criteria regardless of periodic "
            "overnight stays; confirmed cover remained in place and would have responded "
            "to any claim up to the cancellation date; waived the cancellation fee, "
            "refunded part of the premium, and (during the FOS process) offered £150 "
            "compensation for failing to offer this when it first responded to the "
            "complaint. FOS: AXA's initial advice was incorrect and it was reasonable for "
            "the underwriting department to correct it once identified; Mrs Y did have "
            "the benefit of cover throughout the period up to cancellation; travel costs "
            "were not reimbursable since she had the benefit of the policy regardless of "
            "whether she needed to claim, and could have chosen alternative arrangements; "
            "£150 was reasonable redress for the misadvice and the inconvenience it caused"
        ),
        "Outcome Category": "Upheld in Part",
        "Outcome": (
            "AXA to pay Mrs Y the £150 compensation it offered, if she wishes to accept "
            "it; no further redress (including travel costs) required"
        ),
        "Compensation Awarded (£)": 150,
        "Is Core Case": "No — Administrative",
        "Key Policy Clause": (
            "Where a call handler gives incorrect advice about what preserves cover for a "
            "property becoming unoccupied (e.g. a periodic-overnight-stay 'reset' of the "
            "occupancy clock that does not reflect the insurer's actual underwriting "
            "rules), the insurer's underwriting department correcting that advice and "
            "issuing a cancellation notice is not itself unfair, provided the "
            "policyholder had the benefit of cover (and any claims would have been "
            "honoured) up to the cancellation date; compensation for the resulting "
            "confusion and inconvenience is appropriate, but travel costs incurred by a "
            "policyholder visiting a property in reliance on incorrect advice are not "
            "recoverable where the policyholder had the benefit of the policy throughout "
            "and had the option of alternative arrangements not requiring travel"
        ),
        "Missing Evidence": (
            "Evidence that Mrs Y attempted to make, or was prevented from making, a claim "
            "during the period cover remained in force under the misadvised terms — "
            "absent; FOS therefore could not consider any substantive claim-handling "
            "issue, only the misadvice and cancellation process itself"
        ),
        "Ombudsman Reasoning": (
            "Call handler's advice about a 60-day overnight-stay reset did not reflect "
            "AXA's actual underwriting position that the property needed to remain Mrs "
            "Y's permanent home; underwriting department was entitled to correct this and "
            "cancel once it understood her true circumstances; cover and claims-handling "
            "remained available up to the cancellation date, so Mrs Y suffered no loss of "
            "substantive cover; £150 (offered during the FOS process, after AXA accepted "
            "it should have provided redress in its complaint response) was fair for the "
            "confusion and inconvenience; travel costs not recoverable since the policy "
            "benefit was received regardless of use, and alternative non-travel "
            "arrangements existed"
        ),
        "Workflow Insight": (
            "Call handlers giving advice about occupancy-preserving conditions (such as "
            "periodic overnight stays resetting an unoccupancy clock) must ensure the "
            "advice matches actual underwriting rules — incorrect advice of this kind "
            "should be identified and corrected as early as possible, ideally before the "
            "policyholder relies on it or incurs costs; when an insurer accepts it gave "
            "incorrect advice, redress should be offered proactively at the complaint-"
            "response stage rather than only once a case reaches the ombudsman service, "
            "to avoid a second layer of complaint about the redress process itself"
        ),
        "AI Rule Candidate": (
            "IF call_handler_advises_that_periodic_overnight_stays_reset_the_unoccupancy_clock "
            "AND this_advice_does_not_reflect_actual_underwriting_rules "
            "THEN advice_is_incorrect_and_insurer_correcting_it_via_underwriting_review_and_cancellation_is_not_unfair_provided_cover_remained_honoured_up_to_cancellation; "
            "travel_costs_incurred_in_reliance_on_incorrect_occupancy_advice_are_not_recoverable_where_policyholder_had_uninterrupted_benefit_of_cover_throughout; "
            "compensation_for_confusion_and_inconvenience_from_misadvice_should_be_offered_at_first_complaint_response_not_only_after_referral_to_fos"
        ),
        "Source PDF": "DRN-3113837.pdf",
    },
    {
        "Case ID": "UNOC-020",
        "FOS Decision ID": "DRN3273371",
        "Insurer Name": "Lloyds Bank General Insurance Limited",
        "FOS Decision Date": "28 Jul 2018",
        "Claim Type": (
            "Home contents insurance — theft claim declined after policyholder moved out "
            "of the insured property (matrimonial separation) but continued visiting "
            "daily/every other day to care for animals and remained on the council tax; "
            "insurer applied the 30-day unoccupied exclusion; also cited lack of "
            "reasonable care and a family member's deliberate act as alternative grounds"
        ),
        "Unoccupied Period / Circumstance": (
            "Policyholder moved out in November 2016 and had not lived at the property "
            "(in the sense of sleeping and washing there most of the week) for roughly "
            "three months by the time of the theft, despite visiting on a daily or "
            "every-other-day basis and remaining on the council tax; estranged family "
            "member (Mr K) restricted her access to remove her belongings"
        ),
        "Property Type": "Residential home (matrimonial property, policyholder relocated but retained access/visiting rights)",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": (
            "Policy excluded theft cover where the home had not been 'lived in' for more "
            "than 30 days; Ms K had not lived in the property since November 2016, "
            "roughly three months before the claim"
        ),
        "Evidence Dispute": (
            "Ms K: visited the property at least every other day (evidenced by ongoing "
            "care of animals kept there) and remained on the property's council tax, so "
            "she did not consider it unoccupied; a court order did not require immediate "
            "removal of all belongings; Mr K prevented her from accessing the property to "
            "retrieve her belongings. Lloyds: three separate exclusions applied — the "
            "30-day unoccupied exclusion, a lack-of-reasonable-care exclusion, and an "
            "exclusion for theft caused by a deliberate act of a family member. FOS: the "
            "relevant test is whether the policyholder was 'living' in the property — "
            "i.e. sleeping and washing there most of the week — not merely visiting "
            "regularly or remaining registered for council tax; on the evidence, Ms K was "
            "living at another property and only visiting the insured property, which "
            "does not defeat the unoccupied exclusion; since the unoccupied exclusion "
            "alone was sufficient to justify declining the claim, the other two "
            "exclusions did not need to be separately addressed"
        ),
        "Outcome Category": "Not Upheld",
        "Outcome": (
            "Complaint not upheld — Lloyds was entitled to rely on the 30-day unoccupied "
            "exclusion to decline the contents theft claim"
        ),
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": (
            "Where a policy's unoccupancy exclusion turns on whether the home has been "
            "'lived in,' the relevant test is whether the policyholder was actually "
            "sleeping and washing at the property most of the week — regular visits "
            "(even daily or every-other-day) for a specific purpose such as caring for "
            "animals, and remaining registered for council tax at the address, do not by "
            "themselves establish that a policyholder was 'living' at the property if "
            "their primary residence has genuinely moved elsewhere; where an unoccupied-"
            "property exclusion alone is sufficient to justify declining a claim, the "
            "ombudsman need not separately determine whether other exclusions relied upon "
            "by the insurer (e.g. lack of reasonable care, deliberate act by a family "
            "member) would also independently apply"
        ),
        "Missing Evidence": (
            "Evidence that Ms K was sleeping and washing at the insured property for the "
            "majority of the relevant period — absent; her evidence went only to "
            "frequency of visits and council tax registration, neither of which "
            "establishes 'living' at the property"
        ),
        "Ombudsman Reasoning": (
            "Ms K had not lived in the property since November 2016, roughly three months "
            "before the theft; visiting daily or every other day to care for animals and "
            "remaining on the council tax record are consistent with regular visits to a "
            "property she no longer lived in, not with actually living there; the 'lived "
            "in' test requires sleeping and washing there most of the week, which was not "
            "met; unoccupied exclusion could reasonably be relied upon; no need to "
            "consider the two further exclusions Lloyds also cited since the unoccupied "
            "exclusion was sufficient on its own"
        ),
        "Workflow Insight": (
            "When assessing whether a policyholder has 'lived in' a property for the "
            "purposes of an unoccupancy exclusion, claims handlers should apply a "
            "sleeping-and-washing-there test rather than treating regular visits, care of "
            "pets/animals kept at the property, or continued council tax registration as "
            "sufficient evidence of continued occupation; where multiple exclusions could "
            "independently justify declining a claim, it is efficient (and sufficient for "
            "a fair outcome) to establish that one applies without needing to separately "
            "evidence the others"
        ),
        "AI Rule Candidate": (
            "IF policyholder_visits_property_regularly_for_a_specific_purpose_such_as_animal_care_or_remains_registered_for_council_tax "
            "BUT_sleeps_and_washes_primarily_at_another_address "
            "THEN property_is_not_lived_in_by_that_policyholder_for_unoccupancy_exclusion_purposes; "
            "IF one_policy_exclusion_is_independently_sufficient_to_justify_declining_a_claim "
            "THEN fos_and_insurer_need_not_separately_establish_additional_exclusions_also_relied_upon"
        ),
        "Source PDF": "DRN3273371.pdf",
    },
]


def validate_case(case: dict) -> None:
    for field, allowed in CONTROLLED_VOCAB.items():
        value = case.get(field, "")
        if value not in allowed:
            raise ValueError(
                f"Case {case['Case ID']}: '{field}' = '{value}' not in controlled vocab {allowed}"
            )


def append_cases(ws, cases: list) -> None:
    next_row = ws.max_row + 1
    for case_idx, case in enumerate(cases):
        row_fill = ROW_FILL_ODD if (next_row % 2 == 1) else ROW_FILL_EVEN
        for col_idx, col_name in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=next_row, column=col_idx, value=case.get(col_name, ""))
            cell.font      = ROW_FONT
            cell.fill      = row_fill
            cell.border    = ROW_BORDER
            cell.alignment = Alignment(
                horizontal="left", vertical="top", wrap_text=True
            )
        ws.row_dimensions[next_row].height = 80
        next_row += 1


def main() -> None:
    for case in NEW_CASES:
        validate_case(case)

    repo_root = os.path.normpath(os.path.join(os.path.dirname(__file__), ".."))
    xlsx_path = os.path.join(
        repo_root, "knowledge", "case-databases",
        "Unoccupied_Property_Case_Database.xlsx"
    )

    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(
            f"Database not found: {xlsx_path}\n"
            "Run create_unoccupied_property_case_db.py first to create it."
        )

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb["Cases"]

    rows_before = ws.max_row - 1  # exclude header
    append_cases(ws, NEW_CASES)
    rows_after = ws.max_row - 1

    wb.save(xlsx_path)

    print(f"Appended : {rows_after - rows_before} cases")
    print(f"Total    : {rows_after} data rows")
    print(f"Last ID  : {ws.cell(row=ws.max_row, column=1).value}")
    print(f"Saved    : {xlsx_path}")


if __name__ == "__main__":
    main()
