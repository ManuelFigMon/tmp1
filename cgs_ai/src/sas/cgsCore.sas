/*=====================================================================
  Program Name : cgsCore.sas
  Author       : Manuel Figallo
  Purpose      : Core launcher used by every cgs_ai SAS wrapper. Builds a
                 command line in a DATA step, writes it to a .bat, runs it
                 with SYSTASK, and echoes the interpreter output into the
                 SAS log.
  Version      : 1.0beta
  Created      : 2026-08-26

  WHY A DATA STEP AND A .BAT
    An earlier version assembled the command into a macro variable using
    macro-quoted quotes (%str(%")) and passed it straight to SYSTASK. That
    crashed the SAS shell module (sasxshel access violation) because macro
    quoting characters reached the shell layer. Here every value goes
    through symget(), so the quotes are ordinary DATA-step characters and
    NO macro quoting is involved anywhere.

    CRITICAL: a SAS character variable is PADDED to its declared length on
    assignment. Always concatenate strip(piece) -- concatenating the padded
    value overflows `cmd` and silently truncates the trailing quote,
    producing an unterminated command line.

  REQUIREMENT: SYSTASK needs XCMD. Check with
      %put %sysfunc(getoption(xcmd));

  Usage (from a wrapper):
      %cgsBuildArg(name=-OutputFilePath, value=%superq(output_file_path));
      %cgsRun(engine=ps, script=scanFileSystem.ps1, taskname=scanps);
=====================================================================*/

/*---------------------------------------------------------------------
  CONFIGURATION - edit for your environment, or set these from .env.
---------------------------------------------------------------------*/
%global PS_FOLDER_PATH PYTHON_FOLDER_PATH POWERSHELL_EXE PYTHON_EXE CGS_ARG_N;
%let PS_FOLDER_PATH     = \\a70admed.com\R1\CGS\APPS\SAS\UNIT\SAS_G\GSIT_Prod\MANUAL\cgs_ai\src\ps\;
%let PYTHON_FOLDER_PATH = \\a70admed.com\R1\CGS\APPS\SAS\UNIT\SAS_G\GSIT_Prod\MANUAL\cgs_ai\src\py\;
%let POWERSHELL_EXE     = powershell.exe;
%let PYTHON_EXE         = python.exe;


%macro cgsResetArgs;
  /* Start a fresh argument list. Call once per function invocation. */
  %global CGS_ARG_N;
  %let CGS_ARG_N = 0;
%mend cgsResetArgs;


%macro cgsAddArg(name=, value=, always=0);
  /* Record one argument for the pending call.
     name   - the switch, e.g. -OutputFilePath or --output-file-path
     value  - the value; when blank the argument is omitted unless always=1
     always - 1 to emit the switch even with a blank value (flags)

     Arguments are held in INDEXED macro variables -- CGS_ARGNAME1,
     CGS_ARGVAL1, CGS_ARGNAME2, ... -- and never joined into one delimited
     string. List values are themselves ';'-delimited, so any separator
     shared with the values tears them apart: a keyword list of
     "real time;cpu time;accdb;mdb" became four separate command-line
     tokens, and PowerShell bound the strays to whatever positional
     parameter came next.                                                  */
  %global CGS_ARG_N;
  %if %length(%superq(value)) > 0 or &always = 1 %then %do;
    %let CGS_ARG_N = %eval(&CGS_ARG_N + 1);
    %global CGS_ARGNAME&CGS_ARG_N CGS_ARGVAL&CGS_ARG_N;
    %let CGS_ARGNAME&CGS_ARG_N = &name;
    %let CGS_ARGVAL&CGS_ARG_N  = %superq(value);
  %end;
%mend cgsAddArg;


%macro cgsRun(engine=ps, script=, taskname=cgstask, debug=0);
  /* Launch `script` under `engine` (ps|py) with the pending CGS_ARGS.
     Returns the interpreter exit code in the macro variable CGS_RC.
     debug=1 builds and prints the .bat WITHOUT executing it.               */

  %global CGS_RC;
  %local _bat _log _folder _exe _rc;
  %let CGS_RC = ;

  %put NOTE: cgs_ai SAS wrapper version 1.0beta (engine=&engine, script=&script);

  %if %sysfunc(getoption(xcmd)) ne XCMD %then %do;
    %put ERROR: SAS is running with NOXCMD, so SYSTASK cannot launch &engine..;
    %put ERROR- Run the function from a terminal or Task Scheduler instead.;
    %abort cancel 8;
  %end;

  %if &engine = ps %then %do;
    %let _folder = &PS_FOLDER_PATH;
    %let _exe    = &POWERSHELL_EXE;
  %end;
  %else %do;
    %let _folder = &PYTHON_FOLDER_PATH;
    %let _exe    = &PYTHON_EXE;
  %end;

  %if not %sysfunc(fileexist(&_folder.&script)) %then %do;
    %put ERROR: Script not found: &_folder.&script;
    %put ERROR- Fix PS_FOLDER_PATH / PYTHON_FOLDER_PATH at the top of cgsCore.sas.;
    %abort cancel 8;
  %end;

  %let _log = %sysfunc(pathname(work))\&taskname._output.log;
  %let _bat = %sysfunc(pathname(work))\&taskname..bat;

  data _null_;
    /* piece is padded on assignment -- ALWAYS concatenate strip(piece). */
    length cmd $32767 name $256 value $32000;
    file "&_bat" lrecl=32767;

    if upcase(symget('engine')) = 'PS' then
      cmd = '"' || strip(symget('_exe')) || '"'
            || ' -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "'
            || strip(symget('_folder')) || strip(symget('script')) || '"';
    else
      cmd = '"' || strip(symget('_exe')) || '" "'
            || strip(symget('_folder')) || strip(symget('script')) || '"';

    /* Read the arguments from the indexed macro variables. Nothing is
       split here, so a value may safely contain ';' -- which every list
       parameter does. */
    argCount = input(symget('CGS_ARG_N'), best.);
    do i = 1 to argCount;
      name  = strip(symget(cats('CGS_ARGNAME', i)));
      /* A %str() spanning several lines puts CR/LF inside the value; those
         would split the .bat into two lines and truncate the command. */
      value = strip(compress(symget(cats('CGS_ARGVAL', i)), '0D0A'x));
      if value = '' then cmd = strip(cmd) || ' ' || strip(name);
      else               cmd = strip(cmd) || ' ' || strip(name) || ' "' || strip(value) || '"';
    end;

    /* Redirect stdout AND stderr: all progress and error messages go to
       stderr, so without this the SAS log shows only a bare return code. */
    cmd = strip(cmd) || ' > "' || strip(symget('_log')) || '" 2>&1';

    put '@echo off';
    len = length(strip(cmd));
    put cmd $varying32767. len;
  run;

  %put NOTE: ---- generated command (&_bat) ----;
  data _null_;
    infile "&_bat" truncover;
    input line $32767.;
    put line;
  run;

  %if &debug = 1 %then %do;
    %put NOTE: DEBUG=1 -- the .bat was built but NOT executed.;
    %let CGS_RC = 0;
    %return;
  %end;

  systask command """&_bat""" taskname=&taskname status=cgsstat wait;
  waitfor &taskname;

  %put NOTE: ---- &script output ----;
  %if %sysfunc(fileexist(&_log)) %then %do;
    data _null_;
      infile "&_log" truncover;
      input line $32767.;
      put line;
    run;
  %end;
  %else %do;
    %put WARNING: No output was captured at &_log..;
    %put WARNING- cmd.exe could not create the log file, so &engine was probably never started.;
    %put WARNING- Check that the interpreter is on PATH and that WORK is writable.;
  %end;

  %let _rc = &cgsstat;
  %let CGS_RC = &_rc;
  %put NOTE: &script returned &_rc..;

  %if &_rc ne 0 %then %do;
    %put ERROR: &script failed with return code &_rc..;
    %put ERROR- 2 = config error (bad/missing parameter), 3 = I/O error.;
    %put ERROR- 1 = the interpreter itself failed to run the script.;
    %put ERROR- See the captured output above, and re-run the .bat by hand.;
    %abort cancel &_rc;
  %end;
%mend cgsRun;
