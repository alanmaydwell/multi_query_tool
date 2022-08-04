SELECT
hp.process_code,
hp.process_name,
hp.enabled,
hp.date_last_run,
hp.hpls_latest_process_log_status as "Status"
FROM
hub_processes hp
WHERE
hp.process_code in
    (
    'MAAT9'
    ,'MAT10'
    ,'MAT11'
    ,'MAT12'
    ,'EFM20'
    ,'EFM21'
    ,'EFM22'
    ,'EFM10'
    ,'EFM11'
    ,'EFM12')
ORDER BY
process_code