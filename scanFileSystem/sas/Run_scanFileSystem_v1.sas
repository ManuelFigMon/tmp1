/*=====================================================================
  Program Name : Run_scanFileSystem_v1.sas
  Author       : Manuel Figallo
  Purpose      : Optional SYSTASK wrapper that launches scanFileSystem.py.
  Version      : 1.3.1

  NOTE: The authoritative parameter documentation lives in the header of
        scanFileSystem.py. This wrapper only mirrors those parameters onto a
        python.exe command line.

  List parameters (input_folder_root, folder_exclusion_list, extract_keyword,
  file_exclusion_list) are passed as SEMICOLON-DELIMITED strings, which the
  Python side splits back into lists.
=====================================================================*/

/*---------------------------------------------------------------------
  CONFIGURATION - edit these two paths for your environment.
---------------------------------------------------------------------*/
%let PYTHON_EXE  = C:\Python311\python.exe;
%let SCRIPT_PATH = C:\Apps\scanFileSystem\scanFileSystem.py;


%macro scanFileSystem(
    input_folder_root      =,          /* REQUIRED; semicolon-delimited   */
    output_file_path       =,          /* REQUIRED; .csv or .xlsx         */
    file_extensions        =,          /* semicolon-delimited             */
    include_subdirectories = 1,        /* 1 = recurse, 0 = top level only */
    folder_exclusion_list  =,          /* semicolon-delimited; default none */
    file_exclusion_list    =,          /* semicolon-delimited             */
    extract_keyword        =,          /* semicolon-delimited             */
    date_from              =,
    date_to                =,
    date_field             =,          /* created | modified | accessed   */
    metric_profile         =           /* none | sas_log                  */
);

    %local _cmd _rc;

    /*-----------------------------------------------------------------
      Mirror the Python required-parameter validation BEFORE launching.
    -----------------------------------------------------------------*/
    %if %superq(input_folder_root) = %then %do;
        %put ERROR: Required parameter INPUT_FOLDER_ROOT is missing or empty.;
        %abort cancel 8;
    %end;
    %if %superq(output_file_path) = %then %do;
        %put ERROR: Required parameter OUTPUT_FILE_PATH is missing or empty.;
        %abort cancel 8;
    %end;

    /*-----------------------------------------------------------------
      Build the command line. UNC paths are quoted.
    -----------------------------------------------------------------*/
    %let _cmd = "&PYTHON_EXE" "&SCRIPT_PATH";
    %let _cmd = &_cmd --input-folder-root "%superq(input_folder_root)";
    %let _cmd = &_cmd --output-file-path "%superq(output_file_path)";

    %if %superq(file_extensions) ne %then
        %let _cmd = &_cmd --file-extensions "%superq(file_extensions)";
    %if &include_subdirectories = 0 %then
        %let _cmd = &_cmd --no-include-subdirectories;
    %if %superq(folder_exclusion_list) ne %then
        %let _cmd = &_cmd --folder-exclusion-list "%superq(folder_exclusion_list)";
    %if %superq(file_exclusion_list) ne %then
        %let _cmd = &_cmd --file-exclusion-list "%superq(file_exclusion_list)";
    %if %superq(extract_keyword) ne %then
        %let _cmd = &_cmd --extract-keyword "%superq(extract_keyword)";
    %if %superq(date_from) ne %then
        %let _cmd = &_cmd --date-from "%superq(date_from)";
    %if %superq(date_to) ne %then
        %let _cmd = &_cmd --date-to "%superq(date_to)";
    %if %superq(date_field) ne %then
        %let _cmd = &_cmd --date-field "%superq(date_field)";
    %if %superq(metric_profile) ne %then
        %let _cmd = &_cmd --metric-profile "%superq(metric_profile)";

    %put NOTE: Launching: &_cmd;

    /*-----------------------------------------------------------------
      Launch and wait. Non-zero return code aborts the SAS job.
    -----------------------------------------------------------------*/
    systask command "&_cmd" taskname=scanpy status=scanrc wait;
    waitfor scanpy;

    %let _rc = &scanrc;
    %put NOTE: scanFileSystem.py returned &_rc..;

    %if &_rc ne 0 %then %do;
        %put ERROR: scanFileSystem.py failed with return code &_rc..;
        %abort cancel &_rc;
    %end;

    /*-----------------------------------------------------------------
      Fallback if SYSTASK is unavailable in your environment:
        x "&_cmd";
    -----------------------------------------------------------------*/

%mend scanFileSystem;


/*=====================================================================
  EXAMPLES  (see README.md for the matching command-line forms)
=====================================================================*/

/* A. Two roots (HHH + DME), SAS timing profile, Excel out */
%*
%scanFileSystem(
  input_folder_root=%str(\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\HHH;\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\DME),
  output_file_path=C:\Logs\scan.xlsx,
  metric_profile=sas_log,
  extract_keyword=%str(real time;cpu time)
);
*;

/* B. Access-database reference sweep (.accdb/.mdb), CSV out */
%*
%scanFileSystem(
  input_folder_root=%str(\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT),
  output_file_path=C:\Logs\accdb_scan.csv,
  extract_keyword=%str(.accdb;.mdb)
);
*;

/* C. Date-range filter (files modified in H1 2026) */
%*
%scanFileSystem(
  input_folder_root=%str(\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT),
  output_file_path=C:\Logs\recent.csv,
  date_from=2026-01-01,
  date_to=2026-06-30,
  date_field=modified
);
*;

/* D. Defaults only */
%*
%scanFileSystem(
  input_folder_root=%str(\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT),
  output_file_path=C:\Logs\scan.csv
);
*;
