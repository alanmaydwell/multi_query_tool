select
proceeding_id, case_party_id, assessment_status, proceeding_code, matter_type, creation_date, lead_proceeding_ind
from xxccms.xxccms_case_proceedings
where case_party_id in
    (select party_id from
    apps.hz_parties
    where attribute_category = 'CASE'
    and party_name = ?)