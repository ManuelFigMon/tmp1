/*=====================================================================
  Program Name : Run_scanFileSystem_v1.sas
  Author       : Manuel Figallo
  Purpose      : Optional SYSTASK wrapper that launches scanFileSystem.py.
  Version      : 1.3.3

  This file contains the MACRO DEFINITION ONLY. Example calls live in
  sas/Examples_scanFileSystem_v1.sas, which %INCLUDEs this file.

  NOTE: The authoritative parameter documentation lives in the header of
        scanFileSystem.py. This wrapper only mirrors those parameters onto a
        python.exe command line.

  List parameters (input_folder_root, folder_exclusion_list, extract_keyword,
  file_exclusion_list) are passed as SEMICOLON-DELIMITED strings, which the
  Python side splits back into lists. Verified equivalent to passing the
  values as separate command-line arguments.

  Usage:
      %include "<path>\sas\Run_scanFileSystem_v1.sas";
      %scanFileSystem(input_folder_root=..., output_file_path=...);
=====================================================================*/

/*---------------------------------------------------------------------
  CONFIGURATION - edit these two paths for your environment.
  (You may also override them after %INCLUDE-ing this file.)
---------------------------------------------------------------------*/
%global PYTHON_EXE SCRIPT_PATH;
%let PYTHON_EXE  = C:\code\python\cgs_ai\scanFileSystem\.venv\Scripts\python.exe;
%let SCRIPT_PATH = C:\code\python\cgs_ai\scanFileSystem\scanFileSystem.py;


%macro scanFileSystem(
    input_folder_root      =,          /* REQUIRED; semicolon-delimited    */
    output_file_path       =,          /* optional; .csv or .xlsx; omit    */
                                       /* to auto-name scan_<stamp>.csv    */
    file_extensions        =,          /* semicolon-delimited              */
    include_subdirectories = 1,        /* 1 = recurse, 0 = top level only  */
    folder_exclusion_list  =,          /* semicolon-delimited; default none*/
    file_exclusion_list    =,          /* semicolon-delimited              */
    extract_keyword        =,          /* semicolon-delimited              */
    date_from              =,
    date_to                =,
    date_field             =,          /* created | modified | accessed    */
    metric_profile         =           /* none | sas_log                   */
);

    %local _cmd _rc _dq;

    /* A macro-quoted double quote. Using %str(%") keeps the quotes masked
       while the command is being assembled, so they do not unbalance the
       quoted string in the SYSTASK statement below. */
    %let _dq = %str(%");

    /*-----------------------------------------------------------------
      Mirror the Python required-parameter validation BEFORE launching.
      OUTPUT_FILE_PATH is optional (v1.3.2): when omitted, Python writes
      scan_YYYYMMDD_HHMMSS.csv into the working directory.
    -----------------------------------------------------------------*/
    %if %superq(input_folder_root) = %then %do;
        %put ERROR: Required parameter INPUT_FOLDER_ROOT is missing or empty.;
        %abort cancel 8;
    %end;

    /*-----------------------------------------------------------------
      Build the command line. Every path is wrapped in masked double
      quotes so UNC paths and paths containing spaces survive intact.
      %superq() masks semicolons in the list parameters so they do not
      terminate the %LET statements.
    -----------------------------------------------------------------*/
    %let _cmd = &_dq.&PYTHON_EXE.&_dq &_dq.&SCRIPT_PATH.&_dq;
    %let _cmd = &_cmd --input-folder-root &_dq.%superq(input_folder_root)&_dq;

    %if %superq(output_file_path) ne %then
        %let _cmd = &_cmd --output-file-path &_dq.%superq(output_file_path)&_dq;
    %else
        %put NOTE: OUTPUT_FILE_PATH not supplied; Python will auto-name scan_YYYYMMDD_HHMMSS.csv.;

    %if %superq(file_extensions) ne %then
        %let _cmd = &_cmd --file-extensions &_dq.%superq(file_extensions)&_dq;
    %if &include_subdirectories = 0 %then
        %let _cmd = &_cmd --no-include-subdirectories;
    %if %superq(folder_exclusion_list) ne %then
        %let _cmd = &_cmd --folder-exclusion-list &_dq.%superq(folder_exclusion_list)&_dq;
    %if %superq(file_exclusion_list) ne %then
        %let _cmd = &_cmd --file-exclusion-list &_dq.%superq(file_exclusion_list)&_dq;
    %if %superq(extract_keyword) ne %then
        %let _cmd = &_cmd --extract-keyword &_dq.%superq(extract_keyword)&_dq;
    %if %superq(date_from) ne %then
        %let _cmd = &_cmd --date-from &_dq.%superq(date_from)&_dq;
    %if %superq(date_to) ne %then
        %let _cmd = &_cmd --date-to &_dq.%superq(date_to)&_dq;
    %if %superq(date_field) ne %then
        %let _cmd = &_cmd --date-field &_dq.%superq(date_field)&_dq;
    %if %superq(metric_profile) ne %then
        %let _cmd = &_cmd --metric-profile &_dq.%superq(metric_profile)&_dq;

    %put NOTE: Launching: %superq(_cmd);

    /*-----------------------------------------------------------------
      Launch and wait. Non-zero return code aborts the SAS job.
      Python exit codes: 0 = success, 2 = config error, 3 = I/O error.
    -----------------------------------------------------------------*/
    systask command "&_cmd" taskname=scanpy status=scanrc wait;
    waitfor scanpy;

    %let _rc = &scanrc;
    %put NOTE: scanFileSystem.py returned &_rc..;

    %if &_rc ne 0 %then %do;
        %put ERROR: scanFileSystem.py failed with return code &_rc..;
        %put ERROR: 2 = config error (bad/missing parameter), 3 = I/O error.;
        %abort cancel &_rc;
    %end;

    /*-----------------------------------------------------------------
      Fallback if SYSTASK is unavailable in your environment:
        x "&_cmd";
      (X does not return a status, so you lose the return-code check.)
    -----------------------------------------------------------------*/

%mend scanFileSystem;
