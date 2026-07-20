# run_sas_jobs — dataset-driven SAS batch launcher

Launches Windows SAS batch jobs from a control (driver) SAS dataset. By
**default** it submits each job with the **`SYSTASK COMMAND`** statement; a
Boolean flag switches to the classic **`X` command** instead.

## Files
| File | What it is |
|------|-----------|
| `run_sas_jobs.sas` | The `%run_sas_jobs` macro (+ single-job wrapper `%run_sas_job`). |
| `example_run_sas_jobs.sas` | Builds the driver dataset from the requested jobs and calls the macro. |

## Parameters (`%run_sas_jobs`)
| Parameter | Default | Meaning |
|-----------|---------|---------|
| `data=` | *(required)* | Driver dataset — one row per job. |
| `use_x=` | `0` | **Boolean flag.** `0` = `SYSTASK COMMAND` (default), `1` = `X` command. |
| `exepath=` | `D:\SASHome\SASFoundation\9.4\sas.exe` | Full path to the SAS executable. |
| `wait=` | `0` | `0` = launch asynchronously; `1` = wait for each job to finish. |
| `dryrun=` | `0` | `1` = only print the command (safe preview), do not launch. |
| `name_var / sysin_var / log_var` | `jobname / sysin / log` | Column names in the driver dataset. |
| `nosplash_var / nologo_var / icon_var` | `nosplash / nologo / icon` | Numeric flag columns; `1` adds `-NOSPLASH` / `-NOLOGO` / `-ICON`. |
| `dtntoken=` | `{DTN}` | Token in the log path replaced by the timestamp. |

## The `{DTN}` timestamp
Write `{DTN}` wherever the timestamp belongs in the `-LOG` path. At run time the
macro replaces it with `YYYYMMDD_HHMMSS`:

```
...\Logs\PB_Pending_ADR_{DTN}.log   ->   ...\Logs\PB_Pending_ADR_20260719_023040.log
```

## Command that gets built per job
```
"<exepath>" -SYSIN "<sysin>" -LOG "<log-with-timestamp>" [-NOSPLASH] [-NOLOGO] [-ICON]
```

## Usage
```sas
%include "run_sas_jobs.sas";

/* Default: SYSTASK COMMAND, asynchronous */
%run_sas_jobs(data=sasjobs);

/* Use the X command instead */
%run_sas_jobs(data=sasjobs, use_x=1);

/* Preview only (any platform) */
%run_sas_jobs(data=sasjobs, dryrun=1);

/* One-off single job */
%run_sas_job(name=PB_Pending_ADR,
             sysin=\\a70admed.com\R1\CGS\APPS\SAS\PROD\SAS_G\GSIT_PROD\PB\PB_Pending_ADR.sas,
             log=\\a70admed.com\R1\CGS\APPS\SAS\PROD\SAS_G\GSIT_PROD\PB\Logs\PB_Pending_ADR_{DTN}.log);
```

## Notes
- Inner path quotes are doubled (`""`) when building the command so SAS collapses
  them to single quotes and the OS receives correctly quoted, space-safe paths.
- Run `example_run_sas_jobs.sas` with `dryrun=1` to see the exact commands, then
  set `dryrun=0` on the Windows SAS server to launch for real.
