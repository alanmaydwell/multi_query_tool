SELECT
a.acc_code
,a.vat_no
,a.ref_roll_num as "Bank Account Name"
,to_char(a.bank_account) as "Bank Account No."
,f.sort_code
,f.inst_name as "Bank Name"
FROM
accounts a
,financial_institutions f
WHERE
a.fiin_fin_inst_id=f.fin_inst_id
and a.acco_held_by_type='LGL_SRV_SUPS'
and a.acc_code = ?