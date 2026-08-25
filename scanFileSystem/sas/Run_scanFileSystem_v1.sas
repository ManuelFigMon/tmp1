/*=====================================================================
  Program Name : Run_scanFileSystem_v1.sas
  Author       : Manuel Figallo
  Purpose      : Optional SYSTASK wrapper that launches scanFileSystem.py.
  Version      : 1.3.5

  This file contains the MACRO DEFINITION ONLY. Example calls live in
  sas/Examples_scanFileSystem_v1.sas, which %INCLUDEs this file.

  NOTE: The authoritative parameter documentation lives in the header of
        scanFileSystem.py. This wrapper only mirrors those parameters onto a
        python.exe command line.

  HOW THE COMMAND IS BUILT (changed in v1.3.4)
    v1.3.3 assembled the command into a macro variable using macro-quoted
    double quotes (%str(%")) and passed it straight to SYSTASK. That could
    crash the SAS shell module (sasxshel access violation) because macro
    quoting characters reached the shell layer.

    v1.3.4 instead writes the command to a .bat file in the WORK library
    using a DATA step and symget(), so the quotes are ordinary DATA-step
    characters and NO macro quoting is involved anywhere. SYSTASK then runs
    the .bat. The .bat is echoed to the log, so you can always see and
    re-run by hand exactly what was executed.

  List parameters (input_folder_root, folder_exclusion_list, extract_keyword,
  file_exclusion_list) are passed as SEMICOLON-DELIMITED strings, which the
  Python side splits back into lists. Verified equivalent to passing the
  values as separate command-line arguments.

  REQUIREMENT: SYSTASK needs the XCMD system option. Check with
      %put %sysfunc(getoption(xcmd));
  If it reports NOXCMD, shell access is disabled at your site and NO SAS
  wrapper can launch an external program -- run the scanner from a terminal
  or Windows Task Scheduler instead.

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
    metric_profile         =,          /* none | sas_log                   */
    debug                  = 0         /* 1 = build the .bat, do NOT run it*/
);

    %local _bat _log _rc;

    /*-----------------------------------------------------------------
      Announce the wrapper version FIRST. If you do not see this line in
      the log, you are running an OLD copy of this file -- refresh it
      before debugging anything else. The Python script logs its own
      version too, so the two can be compared.
    -----------------------------------------------------------------*/
    %put NOTE: %nrstr(%scanFileSystem) wrapper version 1.3.5;

    /*-----------------------------------------------------------------
      Refuse to run when shell access is disabled, instead of letting
      SYSTASK fail (or crash) deep inside the SAS shell module.
    -----------------------------------------------------------------*/
    %if %sysfunc(getoption(xcmd)) ne XCMD %then %do;
        %put ERROR: SAS is running with NOXCMD, so SYSTASK cannot launch Python.;
        %put ERROR- Run the scanner from a terminal or Task Scheduler instead.;
        %abort cancel 8;
    %end;

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
      Fail early with a clear message if the script is missing. Otherwise
      the interpreter exits 1 and SAS shows only a bare return code.
    -----------------------------------------------------------------*/
    %if not %sysfunc(fileexist(%superq(SCRIPT_PATH))) %then %do;
        %put ERROR: Script not found: %superq(SCRIPT_PATH);
        %put ERROR- Fix the SCRIPT_PATH macro variable at the top of this file.;
        %abort cancel 8;
    %end;

    %let _log = %sysfunc(pathname(work))\scanpy_output.log;
    %let _bat = %sysfunc(pathname(work))\scanFileSystemPy.bat;

    /*-----------------------------------------------------------------
      Build the command in a DATA step.

      Everything comes through symget(), so the parameter values are DATA
      step CHARACTER DATA -- never macro-resolved text. Double quotes are
      literal characters in the string being written, so there is no macro
      quoting to leak into the shell, and semicolons inside the list
      parameters are just data.
    -----------------------------------------------------------------*/
    data _null_;
        length cmd $32767 piece $32767;
        file "&_bat" lrecl=32767;

        cmd = '"' || strip(symget('PYTHON_EXE')) || '" "'
              || strip(symget('SCRIPT_PATH')) || '"';

        cmd = strip(cmd) || ' --input-folder-root "'
              || strip(symget('input_folder_root')) || '"';

        piece = strip(symget('output_file_path'));
        if piece ne '' then cmd = strip(cmd) || ' --output-file-path "' || piece || '"';

        piece = strip(symget('file_extensions'));
        if piece ne '' then cmd = strip(cmd) || ' --file-extensions "' || piece || '"';

        if strip(symget('include_subdirectories')) = '0'
            then cmd = strip(cmd) || ' --no-include-subdirectories';

        piece = strip(symget('folder_exclusion_list'));
        if piece ne '' then cmd = strip(cmd) || ' --folder-exclusion-list "' || piece || '"';

        piece = strip(symget('file_exclusion_list'));
        if piece ne '' then cmd = strip(cmd) || ' --file-exclusion-list "' || piece || '"';

        piece = strip(symget('extract_keyword'));
        if piece ne '' then cmd = strip(cmd) || ' --extract-keyword "' || piece || '"';

        piece = strip(symget('date_from'));
        if piece ne '' then cmd = strip(cmd) || ' --date-from "' || piece || '"';

        piece = strip(symget('date_to'));
        if piece ne '' then cmd = strip(cmd) || ' --date-to "' || piece || '"';

        piece = strip(symget('date_field'));
        if piece ne '' then cmd = strip(cmd) || ' --date-field "' || piece || '"';

        piece = strip(symget('metric_profile'));
        if piece ne '' then cmd = strip(cmd) || ' --metric-profile "' || piece || '"';

        /* Redirect the interpreter's stdout AND stderr to a log file. All
           progress and error messages go to stderr, so without this the SAS
           log shows only a bare return code with no explanation. */
        cmd = strip(cmd) || ' > "' || strip(symget('_log')) || '" 2>&1';

        put '@echo off';
        len = length(strip(cmd));
        put cmd $varying32767. len;
    run;

    /*-----------------------------------------------------------------
      Show the generated .bat in the log. If SAS ever dies during the
      launch below, this is what you re-run manually in a cmd window to
      prove whether the command itself is sound.
    -----------------------------------------------------------------*/
    %put NOTE: ---- generated command (&_bat) ----;
    data _null_;
        infile "&_bat" truncover;
        input line $32767.;
        put line;
    run;

    %if &debug = 1 %then %do;
        %put NOTE: DEBUG=1 -- the .bat was built but NOT executed.;
        %return;
    %end;

    /*-----------------------------------------------------------------
      Launch and wait.

      The .bat path is wrapped in doubled double quotes, which SAS resolves
      to ONE pair of real quotes around the path -- so a WORK path with
      spaces is handled without any macro quoting.

      Exit codes: 0 = success, 2 = config error, 3 = I/O error.
    -----------------------------------------------------------------*/
    systask command """&_bat""" taskname=scanpy status=scanrc wait;
    waitfor scanpy;

    /*-----------------------------------------------------------------
      Echo everything the interpreter printed. This is what turns a bare
      "return code 1" into an actionable message.
    -----------------------------------------------------------------*/
    %put NOTE: ---- scanFileSystem.py output ----;
    %if %sysfunc(fileexist(&_log)) %then %do;
        data _null_;
            infile "&_log" truncover;
            input line $32767.;
            put line;
        run;
    %end;
    %else %do;
        %put WARNING: No output was captured at &_log..;
        %put WARNING- That means cmd.exe could not even create the log file,;
        %put WARNING- so Python was probably never started. Usual causes:;
        %put WARNING-   * the interpreter is not on PATH (check POWERSHELL_EXE/PYTHON_EXE);
        %put WARNING-   * the WORK path is not writable;
        %put WARNING- Run the .bat shown above by hand in a cmd window.;
    %end;

    %let _rc = &scanrc;
    %put NOTE: scanFileSystem.py returned &_rc..;

    %if &_rc ne 0 %then %do;
        %put ERROR: scanFileSystem.py failed with return code &_rc..;
        %put ERROR- 2 = config error (bad/missing parameter), 3 = I/O error.;
        %put ERROR- 1 = the interpreter itself failed to run the script;
        %put ERROR-     (bad path, blocked script, or a parameter it rejected).;
        %put ERROR-     See the captured output above, and re-run the .bat by hand.;
        %abort cancel &_rc;
    %end;

    /*-----------------------------------------------------------------
      Fallback if SYSTASK is unstable in your environment -- run the same
      .bat with the X statement (no return code is captured):
        x """&_bat""";
    -----------------------------------------------------------------*/

%mend scanFileSystem;
