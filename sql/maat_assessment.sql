select
fa.id as "Assessment ID"
,fa.rep_id
,fa.ass_type
,fa.nwor_code
,fa.date_created
,fa.cmu_id
,fa.init_ass_date
,fa.init_app_partner
,fa.init_tot_aggregated_income
,fa.init_adjusted_income_value
,fa.init_result
,fa.fass_init_status
,fa.full_adjusted_living_allowance
,fa.full_tot_annual_disposable_inc
,fa.full_result
,fa.fass_full_status
,fa.replaced
from financial_assessments fa
where rep_id=?