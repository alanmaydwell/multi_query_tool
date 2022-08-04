SELECT
r.id as "MAAT Order ID"
,a.id as "Applicant ID"
,a.first_name
,a.last_name
,a.other_names
,a.dob
,a.gender
,a.ni_number
,a.foreign_id
,a.no_fixed_abode
,r.case_id
,r.cmu_id
,r.supp_account_code
FROM
applicants a
,rep_orders r
WHERE
r.appl_id=a.id
and
r.id = ?