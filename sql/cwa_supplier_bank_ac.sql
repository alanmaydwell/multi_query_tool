SELECT
vsa.vendor_site_code
,vsa.vat_registration_num
,baa.bank_account_name
,to_char(baa.bank_account_num) as "Bank Account No."
,bb.bank_num as "Sort Code"
,bb.banK_name||', '||bb.bank_branch_name as "Bank Name"
FROM
ap.ap_bank_accounts_all baa
,ap.ap_bank_account_uses_all baua
,apps.po_vendor_sites_all vsa
,ap.ap_bank_branches bb
WHERE
baa.bank_account_id=baua.external_bank_account_id
and baua.vendor_site_id=vsa.vendor_site_id
and bb.bank_branch_id = baa.bank_branch_Id
and baua.end_date is null
and baua.primary_flag='Y'
and vsa.attribute3='Legal Services Supplier (Civil/Crime/Both/Mediator)'
and vsa.vendor_site_code = ?