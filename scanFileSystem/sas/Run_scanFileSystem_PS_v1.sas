/*=====================================================================
  Program Name : Run_scanFileSystem_PS_v1.sas
  Author       : Manuel Figallo
  Purpose      : SYSTASK wrapper that launches the PowerShell port,
                 ps\scanFileSystem.ps1.
  Version      : 1.3.3

  This file contains the MACRO DEFINITION ONLY. Example calls live in
  sas/Examples_scanFileSystem_v1.sas.

  Same parameter list as %scanFileSystem() (the Python wrapper) -- the only
  difference is what it launches, so you can swap between the two engines by
  changing the macro name and nothing else.

    %scanFileSystem()     -> python.exe  scanFileSystem.py
    %scanFileSystemPS()   -> powershell.exe  ps\scanFileSystem.ps1

  List parameters (input_folder_root, folder_exclusion_list, extract_keyword,
  file_exclusion_list) are passed as SEMICOLON-DELIMITED strings, which the
  PowerShell side splits back into arrays -- verified to produce output
  identical to passing them as a native PowerShell array.

  Usage:
      %include "<path>\sas\Run_scanFileSystem_PS_v1.sas";
      %scanFileSystemPS(input_folder_root=..., output_file_path=...);
=====================================================================*/

/*---------------------------------------------------------------------
  CONFIGURATION - edit for your environment.
  POWERSHELL_EXE : powershell.exe (Windows PowerShell 5.1, always present)
                   or pwsh.exe    (PowerShell 7+, if installed)
---------------------------------------------------------------------*/
%global POWERSHELL_EXE PS_SCRIPT_PATH;
%let POWERSHELL_EXE  = powershell.exe;
%let PS_SCRIPT_PATH  = C:\code\python\cgs_ai\scanFileSystem\ps\scanFileSystem.ps1;


%macro scanFileSystemPS(
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

    /* A macro-quoted double quote: %str(%") keeps the quotes masked while
       the command is assembled so they do not unbalance the quoted string
       in the SYSTASK statement below. */
    %let _dq = %str(%");

    /*-----------------------------------------------------------------
      Mirror the PowerShell required-parameter validation BEFORE launching.
      OUTPUT_FILE_PATH is optional; when omitted the script auto-names
      scan_YYYYMMDD_HHMMSS.csv in the working directory.
    -----------------------------------------------------------------*/
    %if %superq(input_folder_root) = %then %do;
        %put ERROR: Required parameter INPUT_FOLDER_ROOT is missing or empty.;
        %abort cancel 8;
    %end;

    /*-----------------------------------------------------------------
      Build the command line.
        -NoProfile       : skip the user profile (faster, reproducible)
        -NonInteractive  : never prompt -- essential for unattended runs
        -ExecutionPolicy Bypass : run the .ps1 without changing machine policy
        -File            : run the script, then exit with its exit code
      -File must come LAST; everything after the script path is passed to
      the script as its own parameters.
    -----------------------------------------------------------------*/
    %let _cmd = &_dq.&POWERSHELL_EXE.&_dq -NoProfile -NonInteractive;
    %let _cmd = &_cmd -ExecutionPolicy Bypass -File &_dq.&PS_SCRIPT_PATH.&_dq;

    %let _cmd = &_cmd -InputFolderRoot &_dq.%superq(input_folder_root)&_dq;

    %if %superq(output_file_path) ne %then %do;
        %let _cmd = &_cmd -OutputFilePath &_dq.%superq(output_file_path)&_dq;
    %end;
    %else %do;
        /* %str() masks the semicolon in the message text. Without it the ';'
           ends the %put and the rest becomes stray open code (ERROR 180-322). */
        %put %str(NOTE: OUTPUT_FILE_PATH not supplied; the script will auto-name scan_YYYYMMDD_HHMMSS.csv.);
    %end;

    %if %superq(file_extensions) ne %then
        %let _cmd = &_cmd -FileExtensions &_dq.%superq(file_extensions)&_dq;

    /* Pass 1/0. Do NOT pass $true/$false here: with "powershell.exe -File"
       every argument arrives as a string, so the script takes
       -IncludeSubdirectories as text and parses 1/0/true/false/yes/no
       itself. Verified against the script in -File mode. */
    %if &include_subdirectories = 0 %then %do;
        %let _cmd = &_cmd -IncludeSubdirectories 0;
    %end;
    %else %do;
        %let _cmd = &_cmd -IncludeSubdirectories 1;
    %end;

    %if %superq(folder_exclusion_list) ne %then
        %let _cmd = &_cmd -FolderExclusionList &_dq.%superq(folder_exclusion_list)&_dq;
    %if %superq(file_exclusion_list) ne %then
        %let _cmd = &_cmd -FileExclusionList &_dq.%superq(file_exclusion_list)&_dq;
    %if %superq(extract_keyword) ne %then
        %let _cmd = &_cmd -ExtractKeyword &_dq.%superq(extract_keyword)&_dq;
    %if %superq(date_from) ne %then
        %let _cmd = &_cmd -DateFrom &_dq.%superq(date_from)&_dq;
    %if %superq(date_to) ne %then
        %let _cmd = &_cmd -DateTo &_dq.%superq(date_to)&_dq;
    %if %superq(date_field) ne %then
        %let _cmd = &_cmd -DateField &_dq.%superq(date_field)&_dq;
    %if %superq(metric_profile) ne %then
        %let _cmd = &_cmd -MetricProfile &_dq.%superq(metric_profile)&_dq;

    %put NOTE: Launching: %superq(_cmd);

    /*-----------------------------------------------------------------
      Launch and wait. Non-zero return code aborts the SAS job.
      PowerShell exit codes: 0 = success, 2 = config error, 3 = I/O error.
    -----------------------------------------------------------------*/
    systask command "&_cmd" taskname=scanps status=scanpsrc wait;
    waitfor scanps;

    %let _rc = &scanpsrc;
    %put NOTE: scanFileSystem.ps1 returned &_rc..;

    %if &_rc ne 0 %then %do;
        %put ERROR: scanFileSystem.ps1 failed with return code &_rc..;
        %put ERROR- 2 = config error (bad/missing parameter), 3 = I/O error.;
        %abort cancel &_rc;
    %end;

    /*-----------------------------------------------------------------
      Fallback if SYSTASK is unavailable in your environment:
        x "&_cmd";
      (X does not return a status, so you lose the return-code check.)
    -----------------------------------------------------------------*/

%mend scanFileSystemPS;
