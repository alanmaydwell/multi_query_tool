SELECT
hp.process_code,
hp.process_name,
hp.enabled,
hp.date_last_run,
hp.hpls_latest_process_log_status,
hp.latest_process_log_id,
to_char(hpl.start_time_stamp) as "start_time_stamp"
FROM
hub_processes hp
,hub_process_logs hpl
WHERE
hp.latest_process_log_id=hpl.id(+)
and hp.process_code in
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