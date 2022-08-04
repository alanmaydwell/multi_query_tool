select
pa.id
,pa.rep_id
,pa.nwor_code
,pa.date_created
,pa.cmu_id
,pa.ass_date
,pa.partner_benefit_claimed
,pa.pcob_confirmation
,pa.result
,pa.past_status
,pa.replaced
from
passport_assessments pa
where
pa.rep_id=?