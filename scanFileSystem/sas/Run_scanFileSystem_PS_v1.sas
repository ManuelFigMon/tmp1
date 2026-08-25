/*=====================================================================
  Program Name : Run_scanFileSystem_PS_v1.sas
  Author       : Manuel Figallo
  Purpose      : SYSTASK wrapper that launches the PowerShell port,
                 ps\scanFileSystem.ps1.
  Version      : 1.3.4

  This file contains the MACRO DEFINITION ONLY. Example calls live in
  sas/Examples_scanFileSystem_v1.sas.

  Same parameter list as %scanFileSystem() (the Python wrapper) -- the only
  difference is what it launches, so you can swap between the two engines by
  changing the macro name and nothing else.

    %scanFileSystem()     -> python.exe  scanFileSystem.py
    %scanFileSystemPS()   -> powershell.exe  ps\scanFileSystem.ps1

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
  PowerShell side splits back into arrays.

  REQUIREMENT: SYSTASK needs the XCMD system option. Check with
      %put %sysfunc(getoption(xcmd));
  If it reports NOXCMD, shell access is disabled at your site and NO SAS
  wrapper can launch an external program -- run the scanner from a terminal
  or Windows Task Scheduler instead.

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
    metric_profile         =,          /* none | sas_log                   */
    debug                  = 0         /* 1 = build the .bat, do NOT run it*/
);

    %local _bat _rc;

    /*-----------------------------------------------------------------
      Refuse to run when shell access is disabled, instead of letting
      SYSTASK fail (or crash) deep inside the SAS shell module.
    -----------------------------------------------------------------*/
    %if %sysfunc(getoption(xcmd)) ne XCMD %then %do;
        %put ERROR: SAS is running with NOXCMD, so SYSTASK cannot launch PowerShell.;
        %put ERROR- Run the scanner from a terminal or Task Scheduler instead.;
        %abort cancel 8;
    %end;

    /*-----------------------------------------------------------------
      Mirror the PowerShell required-parameter validation BEFORE launching.
      OUTPUT_FILE_PATH is optional; when omitted the script auto-names
      scan_YYYYMMDD_HHMMSS.csv in the working directory.
    -----------------------------------------------------------------*/
    %if %superq(input_folder_root) = %then %do;
        %put ERROR: Required parameter INPUT_FOLDER_ROOT is missing or empty.;
        %abort cancel 8;
    %end;

    %let _bat = %sysfunc(pathname(work))\scanFileSystemPS.bat;

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

        cmd = '"' || strip(symget('POWERSHELL_EXE')) || '"'
              || ' -NoProfile -NonInteractive -ExecutionPolicy Bypass'
              || ' -File "' || strip(symget('PS_SCRIPT_PATH')) || '"';

        cmd = strip(cmd) || ' -InputFolderRoot "'
              || strip(symget('input_folder_root')) || '"';

        piece = strip(symget('output_file_path'));
        if piece ne '' then cmd = strip(cmd) || ' -OutputFilePath "' || piece || '"';

        piece = strip(symget('file_extensions'));
        if piece ne '' then cmd = strip(cmd) || ' -FileExtensions "' || piece || '"';

        /* Pass 1/0: with "powershell.exe -File" every argument arrives as a
           string, so the script parses 1/0/true/false/yes/no itself. */
        if strip(symget('include_subdirectories')) = '0'
            then cmd = strip(cmd) || ' -IncludeSubdirectories 0';
            else cmd = strip(cmd) || ' -IncludeSubdirectories 1';

        piece = strip(symget('folder_exclusion_list'));
        if piece ne '' then cmd = strip(cmd) || ' -FolderExclusionList "' || piece || '"';

        piece = strip(symget('file_exclusion_list'));
        if piece ne '' then cmd = strip(cmd) || ' -FileExclusionList "' || piece || '"';

        piece = strip(symget('extract_keyword'));
        if piece ne '' then cmd = strip(cmd) || ' -ExtractKeyword "' || piece || '"';

        piece = strip(symget('date_from'));
        if piece ne '' then cmd = strip(cmd) || ' -DateFrom "' || piece || '"';

        piece = strip(symget('date_to'));
        if piece ne '' then cmd = strip(cmd) || ' -DateTo "' || piece || '"';

        piece = strip(symget('date_field'));
        if piece ne '' then cmd = strip(cmd) || ' -DateField "' || piece || '"';

        piece = strip(symget('metric_profile'));
        if piece ne '' then cmd = strip(cmd) || ' -MetricProfile "' || piece || '"';

        put '@echo off';
        len = length(strip(cmd));
        put cmd $varying32767. len;

        /* Echo to the log so you can copy/paste and run it by hand. */
        put "NOTE: command written to &_bat" ;
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
    systask command """&_bat""" taskname=scanps status=scanpsrc wait;
    waitfor scanps;

    %let _rc = &scanpsrc;
    %put NOTE: scanFileSystem.ps1 returned &_rc..;

    %if &_rc ne 0 %then %do;
        %put ERROR: scanFileSystem.ps1 failed with return code &_rc..;
        %put ERROR- 2 = config error (bad/missing parameter), 3 = I/O error.;
        %abort cancel &_rc;
    %end;

    /*-----------------------------------------------------------------
      Fallback if SYSTASK is unstable in your environment -- run the same
      .bat with the X statement (no return code is captured):
        x """&_bat""";
    -----------------------------------------------------------------*/

%mend scanFileSystemPS;
