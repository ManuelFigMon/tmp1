/*=====================================================================
  Program Name : Examples_scanFileSystem_v1.sas
  Author       : Manuel Figallo
  Purpose      : Example calls to the %scanFileSystem() SYSTASK wrapper.
  Version      : 1.3.3

  The macro definition lives in sas/Run_scanFileSystem_v1.sas. This file
  only demonstrates how to call it.

  HOW TO USE
    1. Edit the %INCLUDE path below to point at Run_scanFileSystem_v1.sas.
    2. Optionally override PYTHON_EXE / SCRIPT_PATH for your machine.
    3. Un-comment ONE example and submit.

  Every example is commented out by default (wrapped in %* ... ;) so that
  submitting this whole file does nothing by accident.

  LIST PARAMETERS
    input_folder_root, folder_exclusion_list, extract_keyword and
    file_exclusion_list take SEMICOLON-DELIMITED strings wrapped in %str().
    Python splits them back into a list, which is exactly equivalent to
    passing the values as separate command-line arguments.
=====================================================================*/

/*---------------------------------------------------------------------
  1. Load the macro definition.
---------------------------------------------------------------------*/
%include "C:\code\python\cgs_ai\scanFileSystem\sas\Run_scanFileSystem_v1.sas";

/* PowerShell engine (same parameter list, launches ps\scanFileSystem.ps1) */
%*include "C:\code\python\cgs_ai\scanFileSystem\sas\Run_scanFileSystem_PS_v1.sas";

/*---------------------------------------------------------------------
  2. Optional: override the interpreter / script location for this run.
---------------------------------------------------------------------*/
%*let PYTHON_EXE  = C:\code\python\cgs_ai\scanFileSystem\.venv\Scripts\python.exe;
%*let SCRIPT_PATH = C:\code\python\cgs_ai\scanFileSystem\scanFileSystem.py;


/*=====================================================================
  EXAMPLE A - HHH Old_logs + DME Logs, SAS timing profile, Excel out
  ---------------------------------------------------------------------
  SAS equivalent of this working command line:

    python scanFileSystem.py
      --input-folder-root "\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\HHH\Old_Programs\Old_logs"
                          "\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\DME\Logs"
      --output-file-path "C:\code\python\cgs_ai\tests\scanFileSystem\scan.xlsx"
      --metric-profile sas_log
      --extract-keyword "real time" "cpu time"

  The two roots become ONE semicolon-delimited string; likewise the two
  keywords. Produces scan.xlsx with a Files sheet and a StepDetail sheet.
=====================================================================*/
%*
%scanFileSystem(
  input_folder_root=%str(\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\HHH\Old_Programs\Old_logs;\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\DME\Logs),
  output_file_path=C:\code\python\cgs_ai\tests\scanFileSystem\scan.xlsx,
  metric_profile=sas_log,
  extract_keyword=%str(real time;cpu time)
);
*;


/*=====================================================================
  EXAMPLE B - Two roots (HHH + DME), SAS timing profile, Excel out
=====================================================================*/
%*
%scanFileSystem(
  input_folder_root=%str(\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\HHH;\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\DME),
  output_file_path=C:\Logs\scan.xlsx,
  metric_profile=sas_log,
  extract_keyword=%str(real time;cpu time)
);
*;


/*=====================================================================
  EXAMPLE C - Access-database reference sweep (.accdb/.mdb), CSV out
              Pure keyword use case; metric_profile stays none.
=====================================================================*/
%*
%scanFileSystem(
  input_folder_root=%str(\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT),
  output_file_path=C:\Logs\accdb_scan.csv,
  extract_keyword=%str(.accdb;.mdb)
);
*;


/*=====================================================================
  EXAMPLE D - Date-range filter (files modified in H1 2026)
=====================================================================*/
%*
%scanFileSystem(
  input_folder_root=%str(\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT),
  output_file_path=C:\Logs\recent.csv,
  date_from=2026-01-01,
  date_to=2026-06-30,
  date_field=modified
);
*;


/*=====================================================================
  EXAMPLE E - Defaults only (nothing excluded, metric_profile=none)
=====================================================================*/
%*
%scanFileSystem(
  input_folder_root=%str(\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT),
  output_file_path=C:\Logs\scan.csv
);
*;


/*=====================================================================
  EXAMPLE F - No output path: auto-names scan_YYYYMMDD_HHMMSS.csv in the
              SAS session's working directory.
=====================================================================*/
%*
%scanFileSystem(
  input_folder_root=%str(\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT)
);
*;


/*=====================================================================
  EXAMPLE G - Exclude Old/ and Test/ folders, top level only
=====================================================================*/
%*
%scanFileSystem(
  input_folder_root=%str(\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT),
  output_file_path=C:\Logs\filtered.csv,
  folder_exclusion_list=%str(Old;Test),
  include_subdirectories=0,
  metric_profile=sas_log
);
*;


/*=====================================================================
  POWERSHELL ENGINE - %scanFileSystemPS()
  ---------------------------------------------------------------------
  Identical parameter list to %scanFileSystem(); only the macro name
  changes, so you can swap engines without touching the arguments.
  Requires: %include of sas\Run_scanFileSystem_PS_v1.sas (above).

  NOTE .xlsx output from the PowerShell engine needs the ImportExcel
       module (Install-Module ImportExcel -Scope CurrentUser). Without it
       the scan falls back to .csv with a WARNING -- the same way the
       Python engine falls back when openpyxl is missing.
=====================================================================*/

/* PS-A. The HHH Old_logs + DME Logs scan, PowerShell engine */
%*
%scanFileSystemPS(
  input_folder_root=%str(\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\HHH\Old_Programs\Old_logs;\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\DME\Logs),
  output_file_path=C:\code\python\cgs_ai\tests\scanFileSystem\scan_ps.csv,
  metric_profile=sas_log,
  extract_keyword=%str(real time;cpu time)
);
*;

/* PS-B. Exclude Old/ and Test/, top level only */
%*
%scanFileSystemPS(
  input_folder_root=%str(\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT),
  output_file_path=C:\Logs\filtered_ps.csv,
  folder_exclusion_list=%str(Old;Test),
  include_subdirectories=0,
  metric_profile=sas_log
);
*;

/* PS-C. No output path: auto-names scan_YYYYMMDD_HHMMSS.csv */
%*
%scanFileSystemPS(
  input_folder_root=%str(\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\DME\Logs)
);
*;
