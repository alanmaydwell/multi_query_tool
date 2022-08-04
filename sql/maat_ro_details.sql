select
r.id as "MAAT ID"
,r.case_id
,r.usn
,r.caty_case_type
,r.mcoo_outcome
,r.ioj_result
,r.rder_code
,r.cc_rep_type
,r.cc_rep_decision
,r.rors_status
,r.cmu_id
,r.date_received
,r.ofty_offence_type
FROM
rep_orders r
where
r.id=?