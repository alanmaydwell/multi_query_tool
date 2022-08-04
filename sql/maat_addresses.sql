SELECT
ad1.id,ad1.line1,ad1.line2,ad1.city,ad1.postcode
,ad2.id,ad2.line1,ad2.line2,ad2.city,ad2.postcode
FROM
applicants a
left outer join rep_orders r on r.appl_id=a.id
left outer join addresses ad1 on ad1.id=a.home_addr_id
left outer join addresses ad2 on ad2.id=r.postal_addr_id
WHERE
r.id = ?