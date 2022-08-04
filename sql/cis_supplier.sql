SELECT
sol_off_no
,lss_nm
,nvl(tel_no,' ')
FROM
lgl_srv_sups
WHERE
sol_off_no = ?