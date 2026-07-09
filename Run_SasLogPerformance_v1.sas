/*============================================================================
  Run_SasLogPerformance_v1.sas
  ============================================================================
  Version : 1.0
  Author  : <AUTHOR NAME HERE>
  Date    : 2026-07-09

  PURPOSE
  -------
  Launches the PowerShell SAS log performance crawler
  (Get-SasLogPerformance_v1.ps1) from a SAS session or SAS batch job and
  surfaces the crawler's exit code in the SAS log:

      0 = success
      1 = completed with parse errors (crawler still wrote the output file)
      2 = fatal error (no/partial output - investigate)

  REQUIREMENTS
  ------------
  - SAS must run with XCMD allowed (option XCMD / -allowxcmd). If the session
    is NOXCMD this program writes an ERROR and stops without calling
    PowerShell.
  - The account running SAS needs read access to the .ps1 share and whatever
    the crawler itself reads/writes.

  CONFIG (edit the %let statements below)
  ---------------------------------------
  ps1_dir        Folder containing the PowerShell script.
  ps1_name       Script filename.
  ps1_args       Optional parameter overrides passed to the script, e.g.:
                     %let ps1_args = -OutputPath "E:\Reports\sas_log_metrics.xlsx" -ExcelSheetName "LogMetrics";
                 Leave blank to run with the script defaults.
                 NOTE: to override [bool] parameters (-EnableRobocopy,
                 -EmailEnabled, -IncludeSubdirectories) switch cmd_style to
                 COMMAND below - powershell.exe -File passes booleans as
                 literal text and they will not bind.
  cmd_style      FILE    -> powershell.exe ... -File "script.ps1" <args>
                 COMMAND -> powershell.exe ... -Command "& 'script.ps1' <args>"
                            (required when ps1_args overrides boolean
                            parameters, e.g. -EmailEnabled $false).
                            In COMMAND style quote paths with SINGLE quotes,
                            e.g. %nrstr(-OutputPath 'E:\Reports\metrics.xlsx')
  abort_on_fatal Y -> if the crawler returns 2 (fatal), end the SAS job with
                      "abort return 2" so a scheduler sees the failure.
                      Set to N while testing interactively: abort ends the
                      whole SAS session.
============================================================================*/

/*---------------------------- CONFIG ---------------------------------------*/
%let ps1_dir        = \\a70admed.com\R1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\code\ps\sas_crawler;
%let ps1_name       = Get-SasLogPerformance_v1.ps1;
%let ps1_path       = &ps1_dir.\&ps1_name.;
%let ps1_args       = ;
%let cmd_style      = FILE;
%let abort_on_fatal = N;

/* Do not pop a command window / wait for keypress; wait for the command to
   finish so rc reflects the crawler's real exit code. */
options noxwait xsync;

/*---------------------------- EXECUTION ------------------------------------*/
%macro run_sas_log_crawler;

    /*--- guard: XCMD must be allowed ---------------------------------------*/
    %if %sysfunc(getoption(XCMD)) = NOXCMD %then %do;
        %put ERROR: This SAS session is NOXCMD - operating system commands are disabled.;
        %put ERROR- Ask your SAS admin to start the session with XCMD (-allowxcmd) or;
        %put ERROR- schedule Get-SasLogPerformance_v1.ps1 directly in Windows Task Scheduler.;
        %return;
    %end;

    /*--- guard: script must exist ------------------------------------------*/
    %if %sysfunc(fileexist(&ps1_path.)) = 0 %then %do;
        %put ERROR: PowerShell script not found: &ps1_path.;
        %return;
    %end;

    /*--- build and run the command -----------------------------------------*/
    data _null_;
        length cmd $4000;
        %if %upcase(&cmd_style.) = COMMAND %then %do;
            /* -Command form: booleans like -EmailEnabled $false bind correctly */
            cmd = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command '
                  || '"& ' || "'&ps1_path.' &ps1_args." || '; exit $LASTEXITCODE"';
        %end;
        %else %do;
            /* -File form: fine for string/int/array parameter overrides */
            cmd = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File '
                  || '"' || "&ps1_path." || '"'
                  || " &ps1_args.";
        %end;
        put "NOTE: Launching SAS log performance crawler:";
        put "NOTE- " cmd;

        rc = system(cmd);

        put "NOTE: PowerShell exit code = " rc;
        select (rc);
            when (0) put "NOTE: SAS log crawler completed successfully.";
            when (1) put "WARNING: SAS log crawler completed WITH PARSE ERRORS - "
                         "output file was written; check Get-SasLogPerformance.log "
                         "(next to the output file) for the failing logs.";
            when (2) put "ERROR: SAS log crawler FAILED with a fatal error - "
                         "see Get-SasLogPerformance.log next to the output file.";
            otherwise put "ERROR: Unexpected return code from PowerShell: " rc
                          " (powershell.exe itself may have failed to start).";
        end;
        call symputx('ps_rc', rc);
    run;

    /*--- optionally fail the SAS job so a scheduler sees it -----------------*/
    %if %upcase(&abort_on_fatal.) = Y and &ps_rc. ge 2 %then %do;
        %put ERROR: Aborting SAS job - crawler return code &ps_rc.;
        data _null_;
            abort return 2;
        run;
    %end;

%mend run_sas_log_crawler;

%run_sas_log_crawler;

/*============================================================================
  Console output note: the crawler mirrors everything it prints to a rolling
  text log, Get-SasLogPerformance.log, written next to the output CSV/XLSX -
  open that file to see the full run detail of an unattended execution.
============================================================================*/
