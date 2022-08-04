SELECT
vendor_site_code
,vendor_site_code_alt
,area_code||' '||phone
FROM
po.po_vendor_sites_all
WHERE
vendor_site_code = ?