select
id,
process_code,
erty_error_code,
erco_code
,to_char(date_raised) as "Date Raised"
,hdat_id
,key
,error
,hplo_hub_process_log_id
from errors_raised 
where process_code like ?
and date_raised > ?
order by date_raised desc