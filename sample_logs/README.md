# sample_logs

Synthetic SAS 9.4 log fixtures used to validate `Get-SasLogPerformance_v1.ps1`.

The original `sample_logs/` reference folder was not present in this repository,
so these files were generated to match the specification exactly:

- Windows-1252 encoding with CRLF line endings (headers contain the `©` character)
- `DME_Prepay_Top_Denials_V1_1_1200PM.log` — 34 steps (20 PROC SQL, 8 DATA steps,
  2 IMPORT, 2 PRINT, 1 initialization, 1 session total with real time `35:17.06`)
- `DMEB_SSR_Report_1800PM.log` and `DMEB_SSR_Extract_V2_1100PM.log` — 2-line stub
  logs (only the "Log file opened" line, zero steps)
- `DMEB_POE_Qtrly_Report_0915AM.log` — contains one `ERROR:` and one `WARNING:` line
- `DME_7Z_Denial_Monthly_Reporting_V1_01045AM.log` — session total with an
  `h:mm:ss.ff` real time (`1:02:03.50`)

Replace or augment with real production logs at any time; the parser does not
depend on anything specific to these fixtures.
