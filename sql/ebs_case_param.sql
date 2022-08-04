select
hzp.party_id,
hzp.party_name,
hzp.created_by,
hzp.orig_system_reference,
hzp.creation_date,
hzp.status
from apps.hz_parties hzp
where attribute_category = 'CASE'
and party_name = ?