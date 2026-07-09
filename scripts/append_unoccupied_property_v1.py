"""
Appends Unoccupied Property Batch 1 — UNOC-001 to UNOC-010 — to
knowledge/case-databases/Unoccupied_Property_Case_Database.xlsx.

Run AFTER create_unoccupied_property_case_db.py.
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
        "Case ID": "UNOC-001",
        "FOS Decision ID": "DRN0145579",
        "Insurer Name": "Hyde Park Insurance Consultants Ltd",
        "FOS Decision Date": "Not stated in document",
        "Claim Type": (
            "Rental property insurance — broker (Hyde Park) failed to notify underwriter "
            "that property became unoccupied after tenant's unexpected death (July/August 2010); "
            "theft occurred April 2011 (~8 months unoccupied); underwriter declined theft claim; "
            "policyholders complained broker should have informed them cover was restricted"
        ),
        "Unoccupied Period / Circumstance": (
            "~8 months unoccupied following unexpected death of tenant July/August 2010; "
            "owners intended to sell; unsuccessful auction October 2010; "
            "sale or refurbishment planned for March 2011"
        ),
        "Property Type": "Residential rental property",
        "Dispute Type": "Broker Conduct Dispute",
        "Coverage Decision": "Not Applicable",
        "Rejection Reason": (
            "Underwriter B declined theft claim — property was unoccupied and had not been "
            "notified; would have restricted cover to exclude theft losses; "
            "broker failed to pass on unoccupancy notification"
        ),
        "Evidence Dispute": (
            "Consumers: notified broker of unoccupancy in August 2010; broker failed to pass on; "
            "would have sought specialist cover if told of limitations; high-crime area with prior "
            "burglary history. Broker: policy required written notification; none received. "
            "FOS: weight of evidence indicates consumers did inform broker (wholesale broker "
            "correspondence); broker erred; but on balance consumers would not have taken "
            "specialist cover (cost £1,101.97 vs £250.20; intending to sell within ~3 months)"
        ),
        "Outcome Category": "Not Upheld",
        "Outcome": (
            "Complaint not upheld — broker's error in failing to notify underwriter did not "
            "prejudice consumers; on balance consumers would not have obtained specialist "
            "unoccupied property insurance given intention to sell and cost differential"
        ),
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Broker Dispute",
        "Key Policy Clause": (
            "A broker who is informed of a change in occupancy status must notify the underwriter "
            "and advise the policyholder of the impact on cover; failure to do so is an error, "
            "but does not automatically give rise to financial loss — FOS applies balance of "
            "probabilities to assess whether the consumer would have taken alternative cover "
            "if properly advised; high specialist insurance cost relative to existing premium "
            "and short-term intention to sell are factors weighed against finding that consumer "
            "would have obtained wider cover"
        ),
        "Missing Evidence": (
            "Clear evidence that consumers would definitely have obtained specialist unoccupied "
            "property insurance had they been informed of the cover restrictions — absent given "
            "intention to sell, cost differential, and short anticipated unoccupancy period"
        ),
        "Ombudsman Reasoning": (
            "Broker aware of unoccupancy change (wholesale broker correspondence); broker failed "
            "to advise written notification was required; broker error established; outcome turns "
            "on whether error caused financial loss; consumers attempted to sell within 3 months, "
            "specialist insurance cost 4x existing premium; on balance consumers would not have "
            "sought specialist cover; claim unlikely to have been paid under underwriter A either "
            "given 8 months unoccupancy; broker error did not cause material loss"
        ),
        "Workflow Insight": (
            "When a policyholder reports a change in occupancy status, brokers must (1) "
            "immediately notify the underwriter and (2) specifically advise the policyholder "
            "of the resulting coverage restrictions and the requirement for written confirmation; "
            "when assessing consequential loss from broker error, apply balance of probabilities "
            "test — consumer's intended plans, cost of alternative cover, and how long property "
            "was likely to remain unoccupied are all relevant factors"
        ),
        "AI Rule Candidate": (
            "IF broker_notified_of_occupancy_change AND broker_fails_to_pass_to_underwriter "
            "THEN broker_error_established; "
            "IF specialist_unoccupied_insurance_cost_significantly_exceeds_existing_premium "
            "AND consumer_is_actively_trying_to_sell "
            "THEN balance_of_probabilities_against_consumer_obtaining_specialist_cover; "
            "broker_error_does_not_automatically_create_financial_liability_without_proof_of_consequential_loss"
        ),
        "Source PDF": "DRN0145579.pdf",
    },
    {
        "Case ID": "UNOC-002",
        "FOS Decision ID": "DRN0191656",
        "Insurer Name": "AXA Insurance UK Plc",
        "FOS Decision Date": "Not stated in document",
        "Claim Type": (
            "Residential landlord insurance — theft and malicious damage claim during extended "
            "unoccupancy (~5 months); AXA rejected claim for failure to notify within 90 days "
            "as required by Empty/Unoccupied Property Condition; secondary complaint about delay "
            "in rejecting claim causing alleged loss of rent"
        ),
        "Unoccupied Period / Circumstance": (
            "~5 months unoccupied (tenancy gap); policyholder agreed at inception property was "
            "not currently vacant and not expected to be vacant for periods exceeding 45 "
            "consecutive days; last property inspection 9 days before theft discovered "
            "(2 days overdue against 7-day requirement)"
        ),
        "Property Type": "Residential landlord property",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": (
            "Failure to notify AXA of unoccupancy within 90 days as required by Empty/Unoccupied "
            "Property Condition (page 32 of policy); AXA stated it would have restricted cover "
            "to fire/lightning/explosion/earthquake/aircraft only if notified"
        ),
        "Evidence Dispute": (
            "Mr H: rented properties are inherently subject to vacancy; condition should have "
            "been more prominent; complied with inspection requirements. AXA: unoccupancy "
            "material to underwriting; inspection missed by 2 days; would have restricted "
            "cover if told of extended unoccupancy. FOS: policy stated theft and malicious "
            "damage covered provided condition complied with — notification requirement was "
            "ambiguous for those perils; Mr H substantially complied with inspection exceptions; "
            "2-day shortfall not causally material; AXA must pay claim"
        ),
        "Outcome Category": "Upheld",
        "Outcome": (
            "AXA must deal with theft/malicious damage claim subject to other policy terms; "
            "pay 8% simple interest from date theft discovered to date of payment; "
            "pay £100 compensation for distress and inconvenience"
        ),
        "Compensation Awarded (£)": 100,
        "Is Core Case": "Yes",
        "Key Policy Clause": (
            "An Empty/Unoccupied Property Condition requiring notification within 90 days is "
            "ambiguous where the policy simultaneously states that theft and malicious damage "
            "cover continues after 21 days unoccupancy provided inspection requirements are met "
            "— a policyholder complying with those requirements could reasonably infer that "
            "notification would not affect those specific perils; an insurer cannot rely on a "
            "notification requirement to exclude cover for perils the policy states continue "
            "subject to inspection compliance; a 2-day shortfall in a 7-day inspection cycle "
            "does not bar a claim where the damage would have occurred at a specific time "
            "regardless of when the inspection took place"
        ),
        "Missing Evidence": (
            "Evidence that earlier inspection would have prevented or mitigated the specific "
            "theft/malicious damage that occurred; AXA did not explain how the 2-day shortfall "
            "was material to the loss"
        ),
        "Ombudsman Reasoning": (
            "Policy stated theft and malicious damage covered after 21 days unoccupancy if "
            "Empty/Unoccupied Property Condition complied with — Mr H could reasonably infer "
            "notification requirement related to other cover, not theft/malicious damage; "
            "AXA's argument it would have restricted cover is not clearly stated in policy; "
            "2-day inspection shortfall immaterial — theft occurred at specific time and damage "
            "would not have been prevented; delay in rejecting claim (7 weeks) did not cause "
            "provable rent loss; £100 D&I appropriate for rejection of valid claim"
        ),
        "Workflow Insight": (
            "Unoccupied property conditions that purport to restrict cover for failure to notify "
            "must be clearly drafted — where the same condition simultaneously states cover "
            "continues subject to inspection compliance, the notification requirement is "
            "ambiguous; insurers must be precise about what cover is restricted and when; "
            "a minor shortfall in inspection requirements should not bar a claim unless the "
            "insurer can demonstrate causal materiality to the loss"
        ),
        "AI Rule Candidate": (
            "IF policy_states_theft_cover_continues_provided_inspection_condition_complied_with "
            "AND policyholder_substantially_complies_with_inspection_requirements "
            "THEN failure_to_separately_notify_insurer_of_unoccupancy_does_not_bar_theft_claim; "
            "IF inspection_shortfall_is_minor AND damage_would_have_occurred_regardless "
            "THEN inspection_breach_is_not_causally_material; "
            "insurer_cannot_rely_on_notification_condition_to_restrict_perils_that_policy_states_continue_subject_to_different_conditions"
        ),
        "Source PDF": "DRN0191656.pdf",
    },
    {
        "Case ID": "UNOC-003",
        "FOS Decision ID": "DRN0204977",
        "Insurer Name": "Royal & Sun Alliance Insurance Plc",
        "FOS Decision Date": "Not stated in document",
        "Claim Type": (
            "Rental property insurance (RSA) — escape of water damage; property unoccupied "
            "~9 months while owner arranged refurbishment for future letting; RSA declined "
            "claim — property unoccupied more than 45 consecutive days; 90-day endorsement "
            "further restricted cover to fire/lightning/explosion/aircraft only"
        ),
        "Unoccupied Period / Circumstance": (
            "~9 months unoccupied from policy inception; owner intended to rent to tenants "
            "but refurbishment ongoing; owner lived next door and inspected property daily"
        ),
        "Property Type": "Residential rental property",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": (
            "Policy excluded escape of water where property unoccupied more than 45 consecutive "
            "days; 90-day endorsement further restricted cover to fire/lightning/explosion/ "
            "aircraft only; property unoccupied throughout policy period"
        ),
        "Evidence Dispute": (
            "Mrs Y: living next door and inspecting daily means property is not truly "
            "'unoccupied'; 90-day restriction should have been drawn to her attention. "
            "RSA: property was unoccupied per policy definition; entitled to rely on exclusion. "
            "FOS: daily visits for inspection do not constitute 'lived in' under policy "
            "definition; Summary of Cover flagged 'significant and unusual exclusions' "
            "including unoccupied losses; RSA entitled to decline"
        ),
        "Outcome Category": "Not Upheld",
        "Outcome": (
            "Complaint not upheld — RSA correctly declined claim; property was unoccupied per "
            "policy definition; escape of water excluded; Summary of Cover adequately flagged "
            "unoccupancy exclusion"
        ),
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": (
            "Frequent visits to an unoccupied property (even daily inspections by an owner "
            "living next door) do not constitute the property being 'lived in' for purposes "
            "of an unoccupancy definition — living in requires actual habitation, not "
            "inspection visits; where a policy Summary of Cover document lists losses "
            "when the home is unoccupied under 'Significant and unusual exclusions or "
            "limitations', this is sufficient to draw a policyholder's attention to "
            "unoccupancy-related cover restrictions; an insurer need not specifically "
            "highlight the 90-day endorsement restricting cover to fire/lightning only "
            "where the primary 45-day exclusion already bars the claim type at issue"
        ),
        "Missing Evidence": (
            "Evidence that the property was actually lived in (as opposed to visited) at any "
            "point during the policy period — absent; Mrs Y's own account confirmed property "
            "was not rented out or lived in at any time between policy inception and damage"
        ),
        "Ombudsman Reasoning": (
            "Policy defined unoccupied as 'not lived in by a person authorised by You for "
            "more than 45 consecutive days' (schedule endorsement); daily visits by Mrs Y "
            "did not constitute living in the property; property therefore unoccupied per "
            "definition for entire policy period; Summary of Cover flagged unoccupancy "
            "exclusion under significant/unusual exclusions; RSA entitled to rely on "
            "45-day exclusion to decline EOW claim; no award"
        ),
        "Workflow Insight": (
            "Claims handlers must confirm actual habitation, not mere visits, when assessing "
            "whether the unoccupancy threshold has been met — daily inspection visits by "
            "an owner-neighbour do not satisfy a 'lived in' definition; the relevant "
            "definition is in the policy schedule endorsement, which prevails for that "
            "policyholder; Summary of Cover disclosure of unoccupancy exclusions as "
            "significant/unusual is sufficient to put policyholders on notice"
        ),
        "AI Rule Candidate": (
            "IF property_not_lived_in_for_more_than_policy_unoccupancy_threshold "
            "AND only_activity_is_inspection_visits "
            "THEN property_is_unoccupied_for_policy_purposes; "
            "daily_inspection_by_owner_next_door_does_not_constitute_living_in; "
            "IF summary_of_cover_lists_unoccupancy_exclusion_under_significant_unusual_exclusions "
            "THEN insurer_has_sufficiently_drawn_attention_to_unoccupancy_restrictions; "
            "policy_schedule_endorsement_definition_of_unoccupied_prevails_over_booklet_definition"
        ),
        "Source PDF": "DRN0204977.pdf",
    },
    {
        "Case ID": "UNOC-004",
        "FOS Decision ID": "DRN0413474",
        "Insurer Name": "U K Insurance Limited",
        "FOS Decision Date": "26 Apr 2018",
        "Claim Type": (
            "Buildings insurance for property owned but not occupied — two parts: "
            "(1) escape of water claim ~2011 — complaint time-barred (referred July 2017, "
            "more than 6 months after August 2015 final response); "
            "(2) premium pricing — £300 increase over 2012–2017 due to sustained unoccupancy"
        ),
        "Unoccupied Period / Circumstance": (
            "Property unoccupied for multiple consecutive years (2012–2017); "
            "ongoing unoccupancy increasing risk premium year on year"
        ),
        "Property Type": "Residential property (owned but not occupied by owner)",
        "Dispute Type": "Claim Recording / Administrative Dispute",
        "Coverage Decision": "Not Applicable",
        "Rejection Reason": (
            "(1) Claim complaint time-barred — final response August 2015, FOS referral "
            "July 2017, outside 6-month limit, no exceptional circumstances; "
            "(2) Premium pricing — UKI demonstrated it charges more for unoccupied "
            "properties consistently; if occupied, premium would have been £350+ cheaper"
        ),
        "Evidence Dispute": (
            "Mrs H: complicated circumstances justified late referral; contacted FOS in 2015. "
            "UKI: final response clearly stated time limit; unoccupied properties present "
            "increased risk. FOS: cannot consider claim element — no exceptional circumstances "
            "for late referral; Mrs H's 2015 FOS contact was pre-final response and did not "
            "restart the time limit; premium pricing fair — UKI demonstrated consistent "
            "application to all unoccupied properties"
        ),
        "Outcome Category": "Not Upheld",
        "Outcome": (
            "Complaint about claim cannot be considered (time-barred, no exceptional "
            "circumstances); premium pricing complaint not upheld — UKI demonstrated "
            "fair and consistent pricing for unoccupied properties"
        ),
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Administrative",
        "Key Policy Clause": (
            "FOS cannot consider a claim complaint referred more than 6 months after the "
            "insurer's final response letter unless there are exceptional circumstances — "
            "complexity of circumstances or existence of multiple parties does not in itself "
            "constitute exceptional circumstances; an insurer may legitimately charge higher "
            "premiums for unoccupied properties, provided it demonstrates it does so "
            "consistently for all unoccupied properties based on increased risk; "
            "prior FOS contact before the insurer issues a final response does not preserve "
            "the right to refer the substantive complaint — the 6-month limit runs from "
            "the final response date"
        ),
        "Missing Evidence": (
            "Evidence of exceptional circumstances preventing Mrs H from referring her "
            "complaint within the 6-month window — absent"
        ),
        "Ombudsman Reasoning": (
            "Final response issued 24 August 2015; 6-month limit expired 24 February 2016; "
            "referred July 2017 — out of time; 2015 FOS contact was pre-final response; "
            "complexity does not establish exceptional circumstances; premium pricing: UKI "
            "showed consistent loading for unoccupied properties; demonstrated £350+ "
            "difference between occupied and unoccupied pricing; fair and reasonable"
        ),
        "Workflow Insight": (
            "When issuing a final response, insurers must clearly state the 6-month FOS "
            "referral deadline — this clock starts the same day for all purposes, including "
            "where policyholders have previously been in contact with FOS; insurers should "
            "be able to provide comparative pricing evidence (occupied vs unoccupied premium) "
            "to justify any unoccupancy loading in the event of a challenge"
        ),
        "AI Rule Candidate": (
            "IF fos_referral_date_is_more_than_6_months_after_final_response_date "
            "AND no_exceptional_circumstances "
            "THEN fos_cannot_consider_complaint; "
            "IF insurer_demonstrates_consistent_unoccupancy_premium_loading_applied_to_all_unoccupied_properties "
            "THEN premium_increase_for_unoccupancy_is_fair; "
            "prior_fos_contact_before_final_response_does_not_preserve_or_extend_the_6_month_referral_time_limit"
        ),
        "Source PDF": "DRN0413474.pdf",
    },
    {
        "Case ID": "UNOC-005",
        "FOS Decision ID": "DRN0594830",
        "Insurer Name": "Towergate Underwriting Group Limited",
        "FOS Decision Date": "19 Apr 2020",
        "Claim Type": (
            "Unoccupied property Legal Expenses Insurance (LEI) add-on — sold non-advised; "
            "policy schedule incorrectly showed 'consumer infringement' heading instead of "
            "'property infringement'; LEI claim declined as legal action not covered under "
            "actual terms; Miss T complained of mis-sale and poor service in failing to "
            "send correct documents"
        ),
        "Unoccupied Period / Circumstance": (
            "Unoccupied property (no detail on period — dispute concerns LEI add-on "
            "documentation error, not the property insurance itself)"
        ),
        "Property Type": "Unoccupied property (type not stated in decision)",
        "Dispute Type": "Claim Recording / Administrative Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": (
            "LEI claim declined — legal action Miss T wished to take was not covered under "
            "'property infringement' section (substantive terms identical under both headings; "
            "only the heading differed); add-on sold non-advised (Miss T's responsibility "
            "to ensure suitability)"
        ),
        "Evidence Dispute": (
            "Miss T: given incorrect information (wrong heading on schedule); would not have "
            "purchased if known it did not cover her; Towergate sent incorrect information "
            "repeatedly. Towergate: add-on sold non-advised; cover identical under both headings; "
            "incorrect heading was an error; recommended Miss T read policy wording. "
            "FOS: non-advised sale — Miss T's responsibility to verify suitability; Towergate "
            "provided information and recommended reading policy wording; incorrect heading did "
            "not affect actual terms; £50 appropriate for confusion caused by error"
        ),
        "Outcome Category": "Upheld in Part",
        "Outcome": (
            "Complaint upheld in part — Towergate must pay £50 compensation for trouble and "
            "upset caused by incorrect policy schedule and difficulty obtaining correct "
            "documents; policy not mis-sold; claim correctly declined"
        ),
        "Compensation Awarded (£)": 50,
        "Is Core Case": "No — Administrative",
        "Key Policy Clause": (
            "When an insurance add-on is sold on a non-advised basis, the consumer bears "
            "responsibility for assessing suitability — the provider must supply clear, "
            "fair and not misleading information but is not required to recommend suitability; "
            "an error in a policy schedule heading that does not change the substantive terms "
            "of cover does not give rise to a mis-sale claim; a policy schedule provides only "
            "a brief overview — policyholders are directed to the policy wording for full terms; "
            "compensation is appropriate where incorrect documentation causes confusion even "
            "where the underlying coverage decision was correct"
        ),
        "Missing Evidence": (
            "Evidence that the incorrect 'consumer infringement' heading (as opposed to "
            "'property infringement') would have resulted in the claim being paid — absent, "
            "as the substantive cover under both headings was identical"
        ),
        "Ombudsman Reasoning": (
            "Add-on sold non-advised — Miss T's responsibility to satisfy herself the policy "
            "was suitable; Towergate's letter advised reading enclosed policy wording; "
            "substantive cover identical under both headings; error in heading did not affect "
            "claim outcome; £50 appropriate for confusion and service issues "
            "(difficulty getting duplicate policy documents)"
        ),
        "Workflow Insight": (
            "Insurers and intermediaries must ensure policy schedules accurately reflect "
            "current product terms when insurers change section headings — failure to update "
            "creates confusion even where substantive cover is unchanged; for non-advised "
            "add-on sales, the provider must clearly direct consumers to full policy wording "
            "and invite questions; compensation is appropriate for administrative errors "
            "causing confusion even when the underlying claim decision was correct"
        ),
        "AI Rule Candidate": (
            "IF policy_schedule_error_does_not_change_substantive_terms_of_cover "
            "THEN error_does_not_constitute_mis_sale_or_affect_claim_outcome; "
            "IF add_on_sold_non_advised AND provider_directed_consumer_to_policy_wording "
            "THEN provider_has_met_information_obligations; "
            "administrative_errors_causing_consumer_confusion_warrant_proportionate_compensation_even_where_coverage_decision_correct"
        ),
        "Source PDF": "DRN0594830.pdf",
    },
    {
        "Case ID": "UNOC-006",
        "FOS Decision ID": "DRN0755022",
        "Insurer Name": "St Andrew's Insurance Plc",
        "FOS Decision Date": "27 Nov 2017",
        "Claim Type": (
            "Residential property insurance — mother died June 2016; daughters arranged "
            "unoccupied property cover with St Andrew's; policy cancelled June 2017 as "
            "property remained unoccupied beyond 12 months from death; daughters sought "
            "compensation for lost premium and inconvenience in finding alternative insurer"
        ),
        "Unoccupied Period / Circumstance": (
            "Property unoccupied from death of owner (mother) June 2016; estate under "
            "administration; remained unsold beyond 12 months; St Andrew's only covers "
            "unoccupied property for maximum 12 months from date of unoccupancy"
        ),
        "Property Type": "Residential property (deceased owner's estate)",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Not Applicable",
        "Rejection Reason": (
            "St Andrew's maximum cover period for unoccupied property is 12 months from "
            "date of unoccupancy; policy cancelled June 2017; Ms L-H was clearly told of "
            "this restriction in July 2016 call and she acknowledged it"
        ),
        "Evidence Dispute": (
            "Ms L-H: no warning of 12-month limit; St Andrew's said policy would renew for "
            "'full year'; no documents confirming 12-month period. St Andrew's: told Ms L-H "
            "in July 2016 call — recorded call reviewed by FOS. FOS: reviewed call — "
            "St Andrew's clearly stated 'we can keep the property showing on cover up to "
            "12 months from the date it is shown as unoccupied'; Ms L-H replied 'Ok that's "
            "fine'; later confirmed policy would run until 'last day of May 2017'; "
            "backdating to June 2016 was reasonable; not mis-sold"
        ),
        "Outcome Category": "Not Upheld",
        "Outcome": (
            "Complaint not upheld — St Andrew's clearly communicated 12-month limit in "
            "July 2016 call; Ms L-H acknowledged it; not mis-sold"
        ),
        "Compensation Awarded (£)": 0,
        "Is Core Case": "No — Administrative",
        "Key Policy Clause": (
            "An insurer may impose a maximum cover period for unoccupied properties "
            "(e.g. 12 months from date of unoccupancy) provided this is clearly communicated "
            "to the policyholder — oral communication during a call acknowledged by the "
            "policyholder is sufficient; it is reasonable for an insurer to backdate the "
            "start of the 12-month period to the actual date the property became unoccupied "
            "rather than the date of first notification; where a policyholder is told of "
            "the maximum cover period and acknowledges it, the insurer is not required to "
            "issue separate written confirmation at renewal for it to be enforceable"
        ),
        "Missing Evidence": (
            "Content of any other calls where Ms L-H alleged St Andrew's gave contradictory "
            "information — dates and recordings of other calls not available"
        ),
        "Ombudsman Reasoning": (
            "Listened to 4 July 2016 call — St Andrew's explicitly stated 12-month maximum "
            "cover from date unoccupied; Ms L-H acknowledged it; confirmed policy would run "
            "to last day of May 2017; renewal in November 2016 carried the same limitation; "
            "backdating to June 2016 was reasonable; not mis-sold; complaint not upheld"
        ),
        "Workflow Insight": (
            "When adding an unoccupied property endorsement following notification of "
            "unoccupancy, the insurer must clearly state any maximum cover period and "
            "backdate it to the actual date unoccupancy began (not the date of notification); "
            "at renewal it is best practice to re-state the expiry date of unoccupied cover "
            "especially where the property remains unsold; call recordings are valuable "
            "evidence in disputes about what was communicated"
        ),
        "AI Rule Candidate": (
            "IF insurer_clearly_communicates_maximum_unoccupied_cover_period_in_recorded_call "
            "AND policyholder_acknowledges_it "
            "THEN insurer_may_cancel_policy_when_period_expires_without_further_notice; "
            "unoccupied_cover_period_should_be_backdated_to_actual_date_of_unoccupancy_not_date_of_notification; "
            "written_confirmation_of_oral_communication_about_cover_limits_is_best_practice_but_not_required_where_policyholder_clearly_acknowledged"
        ),
        "Source PDF": "DRN0755022.pdf",
    },
    {
        "Case ID": "UNOC-007",
        "FOS Decision ID": "DRN0803856",
        "Insurer Name": "UK Insurance Limited",
        "FOS Decision Date": "11 Dec 2015",
        "Claim Type": (
            "Landlord insurance for 3 rental properties — one property (rebuild value "
            "£750,000) became vacant November 2014 with no replacement tenants; "
            "UKI immediately removed property from policy stating it only covers unoccupied "
            "properties with rebuild value ≤£250,000; Mr S had to arrange alternative cover "
            "in 2 days at over £2,000; complained of mis-sale — limitation not disclosed "
            "at point of sale"
        ),
        "Unoccupied Period / Circumstance": (
            "Residential rental property vacant from 1 November 2014 after tenancy ended "
            "with no new tenants lined up; owner informed UKI 28 days after becoming vacant"
        ),
        "Property Type": "Residential landlord property (3-property portfolio policy)",
        "Dispute Type": "Endorsement / Exclusion Challenge",
        "Coverage Decision": "Not Applicable",
        "Rejection Reason": (
            "UKI's underwriting criteria do not cover unoccupied properties with rebuild "
            "value above £250,000; property removed from policy immediately on notification; "
            "no grace period because policyholder failed to notify within 20 days as required "
            "by policy Change of Risk condition"
        ),
        "Evidence Dispute": (
            "Mr S: no mention in policy wording that UKI will not insure unoccupied properties "
            "with rebuild value >£250,000; no mention of 20-day notification requirement; "
            "Change of Risk condition not highlighted at point of sale. UKI: policy booklet "
            "clear that all changes must be notified; higher risk for unoccupied properties; "
            "only removed one property, not cancelled whole policy. FOS: Change of Risk "
            "condition is significant with potentially serious consequences (immediate "
            "cessation) — should have been drawn to Mr S's attention at point of sale; "
            "not in policy summary or Key Facts; online referral to policy wording insufficient; "
            "policy silent on £250,000 rebuild limit for unoccupied yet UKI insured two "
            "other properties above this limit; £100 compensation"
        ),
        "Outcome Category": "Upheld in Part",
        "Outcome": (
            "UKI to pay Mr S £100 compensation for failing to draw Change of Risk condition "
            "to his attention at point of sale and for inconvenience of arranging "
            "immediate alternative cover"
        ),
        "Compensation Awarded (£)": 100,
        "Is Core Case": "Yes",
        "Key Policy Clause": (
            "A 'Change of Risk or Interest' policy condition that causes the policy to cease "
            "immediately upon a material change (with no grace period) is significant with "
            "potentially serious consequences and must be specifically drawn to the "
            "policyholder's attention at the point of sale — referral to the policy booklet "
            "during an online sales process is insufficient; where a policy does not state "
            "that unoccupied properties above a certain rebuild value will not be insured, "
            "and the insurer has accepted such properties at inception, policyholders cannot "
            "reasonably be expected to know of this underwriting criterion; the Change of "
            "Risk condition covers the situation where tenants leave a property with no "
            "occupants lined up"
        ),
        "Missing Evidence": (
            "Evidence that UKI drew the Change of Risk condition to Mr S's attention at "
            "point of sale in a policy summary or Key Facts document — absent"
        ),
        "Ombudsman Reasoning": (
            "Change of Risk condition has immediate effect — no grace period; this is serious "
            "and burdensome; should have been highlighted at point of sale; nothing in policy "
            "summary or Key Facts; online policy wording referral insufficient; policy silent "
            "on £250k rebuild limit for unoccupied yet UKI insured two of Mr S's properties "
            "above this without disclosure; on balance Mr S would have acted differently if "
            "limitations disclosed; UKI reasonably only removed one property; "
            "£100 compensation appropriate"
        ),
        "Workflow Insight": (
            "Insurers must specifically highlight in a policy summary or Key Facts document "
            "any condition that causes the policy to terminate immediately without a grace "
            "period upon a material change; where an insurer imposes a rebuild value threshold "
            "for unoccupied properties, this must be stated clearly in the policy "
            "documentation — hidden underwriting criteria that contradict the apparent scope "
            "of cover give rise to complaints; when removing a property from a multi-property "
            "policy due to unoccupancy, insurers should give reasonable notice"
        ),
        "AI Rule Candidate": (
            "IF change_of_risk_condition_causes_immediate_policy_cessation "
            "AND not_highlighted_in_policy_summary_or_key_facts "
            "THEN condition_is_unenforceable_as_basis_for_refusing_compensation; "
            "IF insurer_has_undisclosed_rebuild_value_threshold_for_unoccupied_cover "
            "AND insured_other_properties_above_threshold "
            "THEN threshold_cannot_be_enforced_without_compensation_for_non_disclosure; "
            "significant_onerous_policy_conditions_must_be_specifically_drawn_to_policyholder_attention_at_point_of_sale"
        ),
        "Source PDF": "DRN0803856.pdf",
    },
    {
        "Case ID": "UNOC-008",
        "FOS Decision ID": "DRN0865293",
        "Insurer Name": "Millennium Insurance Company Limited",
        "FOS Decision Date": "23 Dec 2016",
        "Claim Type": (
            "Buildings and contents landlord policy — let property broken into during "
            "refurbishment after tenancy ended; electrical system stolen and property damaged; "
            "Millennium initially rejected theft element (said no theft cover when unoccupied) "
            "then offered to settle below policy excess; Mr K disputed: property unoccupied "
            "fewer than 30 days; discrepancy between schedule ('cover not available' for theft) "
            "and policy booklet (exclusion only after 30+ days)"
        ),
        "Unoccupied Period / Circumstance": (
            "Property unoccupied fewer than 30 days after tenancy ended; break-in and theft "
            "occurred during refurbishment phase; specific days between tenant departure and "
            "theft discovery confirmed to be fewer than 30 days per policy definition"
        ),
        "Property Type": "Residential landlord property",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": (
            "Policy schedule showed 'cover not available' for theft when property unoccupied; "
            "Millennium's loss adjusters relied on this to reject theft element; also disputed "
            "that some damage pre-dated break-in and some repair costs were excessive"
        ),
        "Evidence Dispute": (
            "Mr K: property unoccupied fewer than 30 days per policy definition; policy "
            "booklet states theft excluded only after 30+ days; loss adjusters offered "
            "settlement based on wrong excess. Millennium: schedule shows 'cover not "
            "available' for theft when unoccupied. FOS: discrepancy between schedule "
            "('cover not available' in excess section) and booklet (excluded after 30 days); "
            "policy booklet governs as it could lead policyholders to believe theft cover "
            "applies for first 30 days; 'cover not available' statement in excesses section "
            "not sufficiently prominent; property unoccupied fewer than 30 days so exclusion "
            "does not apply; damage caused in course of theft treated as theft claim"
        ),
        "Outcome Category": "Upheld",
        "Outcome": (
            "Millennium to resume considering Mr K's claim for theft and damage; "
            "pay £350 compensation for distress and inconvenience; if compensation paid "
            "late, 8% simple interest applies"
        ),
        "Compensation Awarded (£)": 350,
        "Is Core Case": "Yes",
        "Key Policy Clause": (
            "Where there is a discrepancy between what the policy schedule and the policy "
            "booklet state about coverage for theft in an unoccupied property, the policy "
            "booklet prevails where it could lead policyholders to reasonably believe theft "
            "cover applies for the first 30 days of unoccupancy — an exclusion hidden in "
            "the 'excesses' section of the schedule under the heading 'cover not available' "
            "is not sufficiently prominent; damage caused in the course of committing a theft "
            "is properly classified as 'loss or damage caused by theft' not as separate "
            "malicious damage; the applicable unoccupancy period runs from the date the "
            "tenancy ended to the date the theft was discovered"
        ),
        "Missing Evidence": (
            "Precise dates establishing the period of unoccupancy between tenant departure "
            "and the break-in discovery — relevant to confirm the fewer-than-30-day period"
        ),
        "Ombudsman Reasoning": (
            "Policy schedule had 'cover not available' for theft when unoccupied in excess "
            "section — not where an exclusion would be expected; policy booklet stated theft "
            "excluded only after 30+ days; booklet governs; period between tenant moving out "
            "and theft discovery was fewer than 30 days confirmed; exclusion does not apply; "
            "damage during course of theft is theft damage; accidental damage cover also "
            "available with lower excess; Millennium must consider full claim; "
            "£350 D&I for poor handling"
        ),
        "Workflow Insight": (
            "Claims handlers must cross-reference both the policy schedule and policy booklet "
            "when assessing unoccupancy exclusions — where documents conflict, the "
            "interpretation more favourable to the policyholder generally applies; "
            "exclusions of theft cover for unoccupied properties must appear clearly in the "
            "exclusions section, not buried in the excesses/schedule; all damage occasioned "
            "in the course of a break-in should be assessed as theft damage unless there is "
            "clear evidence it is pre-existing or unrelated to the theft event"
        ),
        "AI Rule Candidate": (
            "IF policy_schedule_states_theft_cover_not_available_when_unoccupied "
            "AND policy_booklet_states_theft_excluded_only_after_30_days "
            "THEN policy_booklet_prevails_and_theft_cover_applies_for_first_30_days; "
            "exclusion_of_theft_cover_for_unoccupied_properties_must_appear_in_exclusions_section_not_excesses_section; "
            "damage_caused_during_course_of_theft_is_theft_damage_not_separate_malicious_damage; "
            "IF property_unoccupied_for_fewer_days_than_policy_exclusion_threshold "
            "THEN unoccupancy_exclusion_does_not_apply"
        ),
        "Source PDF": "DRN0865293.pdf",
    },
    {
        "Case ID": "UNOC-009",
        "FOS Decision ID": "DRN0966366",
        "Insurer Name": "Royal & Sun Alliance Insurance Plc",
        "FOS Decision Date": "Not stated in document",
        "Claim Type": (
            "Buildings and contents insurance for a second/holiday home — RSA substantially "
            "increased premium at 2009 renewal due to unoccupancy loading; property used "
            "~58–93 days/year with intervals of ~3 weeks between visits; RSA also had to "
            "refund excess premiums from 2005 due to postcode error; Mr L disputed "
            "unoccupancy loading as RSA's explanations were inconsistent"
        ),
        "Unoccupied Period / Circumstance": (
            "Second home / holiday home — occupied ~58–93 days/year in 2010–2012; "
            "intervals between stays ranging from a few days to maximum 5 weeks; "
            "RSA applied unoccupancy loading when total unoccupancy exceeds 60 days per year"
        ),
        "Property Type": "Residential second home (holiday/occasional use property)",
        "Dispute Type": "Claim Recording / Administrative Dispute",
        "Coverage Decision": "Not Applicable",
        "Rejection Reason": (
            "Not applicable — no claim declined; dispute about premium loading for "
            "unoccupancy; RSA's explanations of the loading criteria were inconsistent "
            "and unclear"
        ),
        "Evidence Dispute": (
            "Mr L: RSA told him property is 'unoccupied' only if not lived in for more than "
            "60 consecutive days — his property never met this; RSA inconsistent in "
            "explanations. RSA: loading varies by total annual unoccupancy days (>60 days "
            "total in year); separate from policy definition; commercially sensitive. "
            "FOS: RSA's explanations were inconsistent and jargon-heavy; on balance RSA's "
            "practice is to load when total unoccupancy >60 days per year; Mr L's property "
            "clearly unoccupied >60 days/year; RSA entitled to increase premium but should "
            "have explained better; postcode error refund + interest; £200 compensation"
        ),
        "Outcome Category": "Upheld in Part",
        "Outcome": (
            "RSA to refund excess premiums paid from 2005 due to postcode error plus interest; "
            "pay £200 compensation for poor communication about premium loading; "
            "premium loading for unoccupancy upheld"
        ),
        "Compensation Awarded (£)": 200,
        "Is Core Case": "No — Administrative",
        "Key Policy Clause": (
            "An insurer may apply a premium loading for properties that are unoccupied for "
            "more than a specified period per year, using a different definition of "
            "'unoccupancy' for pricing purposes than the definition stated in the policy "
            "wording for claims purposes — premium determination is separate from claims "
            "assessment, and commercially sensitive pricing criteria need not be fully "
            "disclosed; however, where an insurer gives inconsistent and unclear explanations "
            "for a premium increase, it warrants compensation even where the underlying "
            "pricing decision was legitimate; an insurer must correct pricing errors (such "
            "as a postcode error) and refund all excess premiums with interest from the "
            "date of the error"
        ),
        "Missing Evidence": (
            "Clear written explanation of RSA's premium loading criteria provided to Mr L "
            "at the time of the 2009 increase — absent; RSA's various contradictory "
            "statements about the basis of the unoccupancy loading"
        ),
        "Ombudsman Reasoning": (
            "RSA's explanations inconsistent (variously claimed loading applied to all "
            "unoccupied properties; also said varies by length; told Mr L unoccupied if "
            ">60 consecutive days; told FOS loading based on total annual unoccupancy days); "
            "on balance RSA's practice is total >60 days/year loading; Mr L's property "
            "clearly qualified; RSA entitled to set premiums for unoccupancy risk; "
            "poor explanation doesn't create financial loss as Mr L could have switched; "
            "postcode error refund + interest already agreed; £200 for poor communication"
        ),
        "Workflow Insight": (
            "Insurers applying unoccupancy premium loadings must be able to provide clear, "
            "consistent and jargon-free explanations to policyholders when challenged — "
            "inconsistent statements about loading criteria can give rise to compensation "
            "awards even where the underlying loading is legitimate; pricing definitions of "
            "'unoccupancy' (e.g. total annual days) may differ from the policy wording "
            "definition — policyholders should be told about this distinction if relevant"
        ),
        "AI Rule Candidate": (
            "IF insurer_applies_premium_loading_for_unoccupancy "
            "AND provides_inconsistent_explanations "
            "THEN compensation_for_poor_communication_appropriate_even_if_loading_is_legitimate; "
            "insurer_may_use_different_unoccupancy_definition_for_pricing_than_for_claims_but_should_disclose_distinction; "
            "IF pricing_error_overcharged_premiums "
            "THEN insurer_must_refund_excess_premiums_plus_interest_from_date_of_error"
        ),
        "Source PDF": "DRN0966366.pdf",
    },
    {
        "Case ID": "UNOC-010",
        "FOS Decision ID": "DRN1400739",
        "Insurer Name": "Saga Services Limited",
        "FOS Decision Date": "Not stated in document",
        "Claim Type": (
            "Home insurance — Mr A (original policyholder) died 2010; Mr C (executor) "
            "informed Saga and continued policy with unoccupied property endorsement; "
            "endorsement required either drain-down or heating at constant temperature "
            "plus weekly checks; Mr C complied; at 2011 renewal Saga silently removed "
            "the endorsement as 'standard procedure'; February 2012 escape of water from "
            "boiler fault; claim declined as policy no longer had escape of water cover "
            "for unoccupied property; £7,500 deducted from property sale price to account "
            "for water damage"
        ),
        "Unoccupied Period / Circumstance": (
            "Property unoccupied from owner's death in 2010; estate under administration; "
            "remained unoccupied through 2011 into 2012; executor kept heating on and "
            "neighbour monitored property in compliance with endorsement requirements; "
            "endorsement silently removed at 2011 renewal"
        ),
        "Property Type": "Residential home (deceased owner's estate)",
        "Dispute Type": "Coverage Dispute",
        "Coverage Decision": "Declined — Full",
        "Rejection Reason": (
            "Saga stated unoccupied property endorsement was only included for one year by "
            "standard procedure; at 2011 renewal endorsement was removed; policy reverted "
            "to standard terms which did not cover escape of water in unoccupied property; "
            "standard policy terms applied at date of loss (February 2012)"
        ),
        "Evidence Dispute": (
            "Mr C: Saga did not draw to his attention that the endorsement would be removed "
            "at renewal; continued to comply with endorsement terms throughout 2011 and into "
            "2012 unaware they no longer applied. Saga: standard procedure to remove "
            "endorsement at first renewal; standard terms applied. FOS: reviewed 2010 and "
            "2011 documents — no mention in renewal documents, covering letter, or Important "
            "Notice of the endorsement removal; change was material and onerous and should "
            "have been specifically highlighted; Mr C continued complying with endorsement "
            "conditions (heating, weekly checks) showing he believed endorsement still "
            "applied and would have sought alternative cover if told otherwise; "
            "Saga's failure caused estate a £7,500 loss"
        ),
        "Outcome Category": "Upheld",
        "Outcome": (
            "Saga to pay £7,500 to estate of Mr A plus 8% simple interest from date "
            "of loss to date of payment"
        ),
        "Compensation Awarded (£)": 0,
        "Is Core Case": "Yes",
        "Key Policy Clause": (
            "Where an insurer removes an endorsement at renewal, this is a material change "
            "in cover that must be specifically highlighted to the policyholder — merely "
            "sending renewal documents with standard terms is insufficient where the previous "
            "policy had bespoke endorsements; the fact that a policyholder continues to "
            "comply with the now-removed endorsement conditions is strong evidence that he "
            "was unaware of the change and would have sought alternative cover if notified; "
            "an insurer's standard internal procedure to remove unoccupied property "
            "endorsements at renewal is not well-known industry practice and cannot be "
            "assumed known by policyholders; where an insurer's failure to notify a material "
            "reduction in cover causes a provable financial loss, the insurer is liable for "
            "that loss"
        ),
        "Missing Evidence": (
            "Evidence that the endorsement removal was communicated to Mr C at the 2011 "
            "renewal — absent from renewal documents, covering letter, and Important Notice"
        ),
        "Ombudsman Reasoning": (
            "Saga informed of unoccupancy at inception and added appropriate endorsement; "
            "'standard procedure' to remove endorsement at first renewal is not well-known "
            "industry practice; 2011 renewal documents contained no notification that "
            "endorsement was being removed; Mr C continued to comply with endorsement "
            "requirements demonstrating he believed they still applied; Saga failed to "
            "specifically draw the change to Mr C's attention; on balance Mr C would have "
            "sought alternative cover including escape of water; £7,500 reduction in sale "
            "price is provable loss caused by Saga's failure"
        ),
        "Workflow Insight": (
            "Insurers must specifically highlight any reduction in cover at renewal — "
            "particularly where an endorsement that provided additional cover is being "
            "removed; where a policyholder is maintaining the property in compliance with "
            "conditions of an endorsement, removing that endorsement without notification "
            "is misleading; the test for consequential liability where insurer fails to "
            "notify a material cover reduction is: would the policyholder on the balance "
            "of probabilities have obtained alternative cover? Where the policyholder was "
            "demonstrably complying with the endorsement conditions, the answer is almost "
            "always yes; sale price reductions to account for uninsured property damage "
            "are recoverable from the insurer that failed to notify the cover change"
        ),
        "AI Rule Candidate": (
            "IF insurer_removes_unoccupied_property_endorsement_at_renewal "
            "AND does_not_specifically_highlight_removal "
            "THEN removal_is_not_effective_against_policyholder_and_insurer_may_be_liable_for_uninsured_losses; "
            "IF policyholder_continues_to_comply_with_removed_endorsement_conditions "
            "THEN strong_evidence_policyholder_was_unaware_of_removal_and_would_have_sought_alternative_cover; "
            "material_reduction_in_cover_at_renewal_must_be_specifically_highlighted_not_merely_included_in_renewal_pack; "
            "standard_internal_procedure_to_remove_endorsements_at_renewal_is_not_sufficient_notice"
        ),
        "Source PDF": "DRN1400739.pdf",
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
            "Run create_unoccupied_property_case_db.py first."
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
