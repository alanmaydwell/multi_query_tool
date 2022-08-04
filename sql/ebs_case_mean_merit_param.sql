select
*
from xxccms.xxccms_opa_assessments
where case_party_id in
    (select party_id from
    apps.hz_parties
    where attribute_category = 'CASE'
    and party_name =  ?)